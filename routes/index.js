import express from "express";
import graph from "../helpers/graphHelper.js";
import certHelper from "../helpers/certHelper.js";
import dbHelper from "../helpers/dbHelper.js";

const router = express.Router();

// GET /
router.get('/', async function (req, res, next) {
  const users = [];
  const calendars = [];
  const selectedUserId = req.query.userId || process.env.USER_ID || '';
  const selectedCalendarId = req.query.calendarId || '';

  try {
    const client = graph.getGraphClientForApp(req.app.locals.msalClient);
    const response = await client
      .api('/users')
      .select('id,displayName,mail,userPrincipalName')
      .top(25)
      .get();

    if (response && response.value) {
      for (const user of response.value) {
        users.push({
          id: user.id,
          name: user.displayName || user.userPrincipalName || user.mail || user.id,
          email: user.mail || user.userPrincipalName || '',
        });
      }
    }

    if (selectedUserId) {
      const calendarsResponse = await client
        .api(`/users/${selectedUserId}/calendars`)
        .select('id,name,isDefaultCalendar')
        .top(50)
        .get();

      if (calendarsResponse && calendarsResponse.value) {
        for (const calendar of calendarsResponse.value) {
          calendars.push({
            id: calendar.id,
            name: calendar.name || calendar.id,
            isDefaultCalendar: calendar.isDefaultCalendar === true,
          });
        }
      }
    }
  } catch (error) {
    console.log(`Unable to load users/calendars for app-only subscribe: ${error.message}`);
  }

  // Ensure the configured fallback user is always available.
  if (process.env.USER_ID && !users.some((u) => u.id === process.env.USER_ID)) {
    users.unshift({
      id: process.env.USER_ID,
      name: process.env.USER_ID,
      email: '',
    });
  }

  res.render('index', {
    title: 'Microsoft Graph Notifications Sample',
    users,
    calendars,
    selectedUserId,
    selectedCalendarId,
  });
});

router.get('/user-flow', async function (req, res) {
  const selectedUserId = process.env.USER_ID || '';
  let selectedUserName = selectedUserId;
  const calendars = [];
  const subscribedCount = Number.parseInt(req.query.subscribed || '0', 10);
  const failedCount = Number.parseInt(req.query.failed || '0', 10);

  if (selectedUserId) {
    try {
      const client = graph.getGraphClientForApp(req.app.locals.msalClient);
      const selectedUser = await client
        .api(`/users/${selectedUserId}`)
        .select('displayName,mail,userPrincipalName')
        .get();

      selectedUserName =
        selectedUser.displayName ||
        selectedUser.userPrincipalName ||
        selectedUser.mail ||
        selectedUserId;

      const calendarsResponse = await client
        .api(`/users/${selectedUserId}/calendars`)
        .select('id,name,isDefaultCalendar')
        .top(100)
        .get();

      if (calendarsResponse && calendarsResponse.value) {
        for (const calendar of calendarsResponse.value) {
          calendars.push({
            id: calendar.id,
            name: calendar.name || calendar.id,
            isDefaultCalendar: calendar.isDefaultCalendar === true,
          });
        }
      }
    } catch (error) {
      console.log(`Unable to load calendars for user flow: ${error.message}`);
    }
  }

  res.render('user-flow', {
    title: 'User Flow',
    selectedUserId,
    selectedUserName,
    calendars,
    subscribedCount: Number.isNaN(subscribedCount) ? 0 : subscribedCount,
    failedCount: Number.isNaN(failedCount) ? 0 : failedCount,
  });
});

router.post('/user-flow/subscribe', async function (req, res) {
  const selectedUserId = process.env.USER_ID || '';
  const selectedCalendarIds = Array.isArray(req.body.calendarIds)
    ? req.body.calendarIds
    : req.body.calendarIds
      ? [req.body.calendarIds]
      : [];

  if (!selectedUserId) {
    req.flash('error_msg', {
      message: 'USER_ID is not configured.',
      debug: 'Set USER_ID in environment before subscribing.',
    });
    res.redirect('/user-flow');
    return;
  }

  if (!selectedCalendarIds.length) {
    req.flash('error_msg', {
      message: 'No calendars selected.',
      debug: 'Select at least one calendar and click Subscribe.',
    });
    res.redirect('/user-flow');
    return;
  }

  const client = graph.getGraphClientForApp(req.app.locals.msalClient);
  const notificationHost = `https://${req.hostname}`;
  const existingSubscriptions = dbHelper.getSubscriptionsByUserAccountId('APP-ONLY-USER-FLOW');

  for (const existingSub of existingSubscriptions) {
    try {
      await client.api(`/subscriptions/${existingSub}`).delete();
    } catch (err) {
      console.error(err);
    }

    dbHelper.deleteSubscription(existingSub);
  }

  let subscribedCount = 0;
  let failedCount = 0;

  for (const calendarId of selectedCalendarIds) {
    try {
      const subscription = await client.api('/subscriptions').create({
        changeType: 'created',
        notificationUrl: `${notificationHost}/listen`,
        lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
        resource: `users/${selectedUserId}/calendars/${calendarId}/events?$select=id,subject,start,end,organizer,bodyPreview,attendees,location`,
        clientState: process.env.SUBSCRIPTION_CLIENT_STATE,
        includeResourceData: true,
        encryptionCertificate: certHelper.getSerializedCertificate(
          process.env.CERTIFICATE_PATH,
        ),
        encryptionCertificateId: process.env.CERTIFICATE_ID,
        expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
      });

      dbHelper.addSubscription(subscription.id, 'APP-ONLY-USER-FLOW');
      subscribedCount += 1;
    } catch (error) {
      console.log(`Unable to subscribe for calendar ${calendarId}: ${error.message}`);
      failedCount += 1;
    }
  }

  res.redirect(`/user-flow?subscribed=${subscribedCount}&failed=${failedCount}`);
});

router.get('/admin-flow', function (req, res) {
  res.render('admin-flow', {
    title: 'Admin Flow',
  });
});

export default router;
