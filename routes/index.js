import express from "express";
import graph from "../helpers/graphHelper.js";
import certHelper from "../helpers/certHelper.js";
import dbHelper from "../helpers/dbHelper.js";
import subscriptionStateHelper from "../helpers/subscriptionStateHelper.js";

const router = express.Router();

function parseSelectedFields(resource) {
  const rawResource = String(resource || '');
  if (!rawResource) {
    return [];
  }

  const queryIndex = rawResource.indexOf('?');
  if (queryIndex < 0 || queryIndex === rawResource.length - 1) {
    return [];
  }

  const queryString = rawResource.slice(queryIndex + 1);
  const params = new URLSearchParams(queryString);
  const rawSelect = params.get('$select') || params.get('%24select');
  if (!rawSelect) {
    return [];
  }

  return rawSelect
    .split(',')
    .map((field) => field.trim())
    .filter((field) => field.length > 0);
}

function parseCalendarSubscriptionResource(resource) {
  const rawResource = String(resource || '').trim().replace(/^\//, '');
  if (!rawResource) {
    return null;
  }

  const pathOnly = rawResource.split('?')[0];
  const segments = pathOnly.split('/').filter((segment) => segment);

  // Expected: users/{userId}/calendars/{calendarId}/events
  if (segments.length < 5 || segments[0].toLowerCase() !== 'users') {
    return null;
  }

  if (segments[2].toLowerCase() !== 'calendars') {
    return null;
  }

  if (segments[4].toLowerCase() !== 'events') {
    return null;
  }

  return {
    userId: segments[1],
    calendarId: segments[3],
  };
}

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
  const subscribedCalendarIds = new Set();
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

      const existingSubscriptions = dbHelper.getSubscriptionsByUserAccountId('APP-ONLY-USER-FLOW');
      for (const subscriptionId of existingSubscriptions) {
        try {
          const existingSubscription = await client
            .api(`/subscriptions/${subscriptionId}`)
            .get();
          // Normalize: strip optional leading slash so matching works regardless
          // of whether Graph returns "users/..." or "/users/..."
          const resource = String(existingSubscription.resource || '').toLowerCase().replace(/^\//, '');
          const marker = `users/${selectedUserId.toLowerCase()}/calendars/`;

          if (resource.includes(marker)) {
            const afterMarker = resource.split(marker)[1] || '';
            const calendarId = afterMarker.split('/')[0] || '';
            if (calendarId) {
              subscribedCalendarIds.add(calendarId.toLowerCase());
            }
          }
        } catch (subError) {
          console.log(`Unable to load subscription ${subscriptionId}: ${subError.message}`);
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
    subscribedCalendarIds: [...subscribedCalendarIds],
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
  let selectedUserName = selectedUserId;
  const calendarNameById = new Map();

  try {
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
      .select('id,name')
      .top(100)
      .get();

    if (calendarsResponse && calendarsResponse.value) {
      for (const calendar of calendarsResponse.value) {
        calendarNameById.set(
          String(calendar.id || '').toLowerCase(),
          calendar.name || calendar.id,
        );
      }
    }
  } catch (lookupError) {
    console.log(`Unable to load user/calendar names for clientState metadata: ${lookupError.message}`);
  }

  for (const calendarId of selectedCalendarIds) {
    try {
      const calendarName =
        calendarNameById.get(String(calendarId || '').toLowerCase()) ||
        calendarId;

      const subscription = await client.api('/subscriptions').create({
        changeType: 'created',
        notificationUrl: `${notificationHost}/listen`,
        lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
        resource: `users/${selectedUserId}/calendars/${calendarId}/events?$select=id,subject,start,end,organizer,bodyPreview,attendees,location`,
        clientState: subscriptionStateHelper.buildClientState(
          process.env.SUBSCRIPTION_CLIENT_STATE,
          {
            userName: selectedUserName,
            calendarName: calendarName,
          },
        ),
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

router.get('/admin-flow', async function (req, res) {
  const users = [];
  const calendars = [];
  const graphSubscriptions = [];
  const selectedCalendarIds = new Set();
  const selectedUserId = req.query.userId || process.env.USER_ID || '';
  const subscribedCount = Number.parseInt(req.query.subscribed || '0', 10);
  const unsubscribedCount = Number.parseInt(req.query.unsubscribed || '0', 10);
  const failedCount = Number.parseInt(req.query.failed || '0', 10);

  try {
    const client = graph.getGraphClientForApp(req.app.locals.msalClient);
    const usersResponse = await client
      .api('/users')
      .select('id,displayName,mail,userPrincipalName')
      .top(50)
      .get();

    if (usersResponse && usersResponse.value) {
      for (const user of usersResponse.value) {
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
    }

    // Fetch ALL subscriptions from Graph
    const graphSubsResponse = await client.api('/subscriptions').get();

    // Collect every ID the DB knows about
    const subscriptionOwners = ['APP-ONLY', 'APP-ONLY-USER-FLOW'];
    const dbSubscriptionIds = new Map(); // id -> owner
    for (const owner of subscriptionOwners) {
      for (const id of dbHelper.getSubscriptionsByUserAccountId(owner)) {
        dbSubscriptionIds.set(id, owner);
      }
    }

    if (graphSubsResponse && graphSubsResponse.value) {
      for (const sub of graphSubsResponse.value) {
        const inDb = dbSubscriptionIds.has(sub.id);
        let clientStateValue = sub.clientState;

        // Some Graph list responses omit clientState. Fetch full details as fallback.
        if (!clientStateValue) {
          try {
            const details = await client
              .api(`/subscriptions/${sub.id}`)
              .get();
            clientStateValue = details ? details.clientState : '';
          } catch (detailsError) {
            console.log(`Unable to load clientState for subscription ${sub.id}: ${detailsError.message}`);
          }
        }

        const parsedState = subscriptionStateHelper.parseClientState(clientStateValue);
        const stateMetadata = parsedState.metadata || {};
        const resourceInfo = parseCalendarSubscriptionResource(sub.resource);

        if (
          selectedUserId &&
          resourceInfo &&
          resourceInfo.userId.toLowerCase() === selectedUserId.toLowerCase()
        ) {
          selectedCalendarIds.add(String(resourceInfo.calendarId || '').toLowerCase());
        }

        graphSubscriptions.push({
          id: sub.id,
          changeType: sub.changeType || '',
          fields: parseSelectedFields(sub.resource),
          expirationDateTime: sub.expirationDateTime || '',
          notificationUrl: sub.notificationUrl || '',
          owner: inDb ? dbSubscriptionIds.get(sub.id) : '',
          inDb,
          userName: stateMetadata.userName || '',
          calendarName: stateMetadata.calendarName || '',
        });
      }
    }
  } catch (error) {
    console.log(`Unable to load admin flow data: ${error.message}`);
  }

  graphSubscriptions.sort((a, b) => {
    if (!a.expirationDateTime && !b.expirationDateTime) return 0;
    if (!a.expirationDateTime) return 1;
    if (!b.expirationDateTime) return -1;
    return a.expirationDateTime.localeCompare(b.expirationDateTime);
  });

  res.render('admin-flow', {
    title: 'Admin Flow',
    users,
    calendars,
    selectedUserId,
    selectedCalendarIds: [...selectedCalendarIds],
    graphSubscriptions,
    subscribedCount: Number.isNaN(subscribedCount) ? 0 : subscribedCount,
    unsubscribedCount: Number.isNaN(unsubscribedCount) ? 0 : unsubscribedCount,
    failedCount: Number.isNaN(failedCount) ? 0 : failedCount,
  });
});

router.post('/admin-flow/sync-subscriptions', async function (req, res) {
  const selectedUserId = String(req.body.userId || '').trim();
  const selectedCalendarIdsRaw = Array.isArray(req.body.calendarIds)
    ? req.body.calendarIds
    : req.body.calendarIds
      ? [req.body.calendarIds]
      : [];

  if (!selectedUserId) {
    res.redirect('/admin-flow');
    return;
  }

  const selectedByLower = new Map();
  for (const calendarId of selectedCalendarIdsRaw) {
    const normalized = String(calendarId || '').trim();
    if (!normalized) {
      continue;
    }

    selectedByLower.set(normalized.toLowerCase(), normalized);
  }

  const selectedSet = new Set(selectedByLower.keys());

  const client = graph.getGraphClientForApp(req.app.locals.msalClient);
  const notificationHost = `https://${req.hostname}`;

  let selectedUserName = selectedUserId;
  const calendarNameById = new Map();

  try {
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
      .select('id,name')
      .top(200)
      .get();

    if (calendarsResponse && calendarsResponse.value) {
      for (const calendar of calendarsResponse.value) {
        calendarNameById.set(
          String(calendar.id || '').toLowerCase(),
          calendar.name || calendar.id,
        );
      }
    }
  } catch (lookupError) {
    console.log(`Unable to load admin sync metadata for user ${selectedUserId}: ${lookupError.message}`);
  }

  const existingByCalendar = new Map();
  try {
    const graphSubsResponse = await client.api('/subscriptions').get();
    if (graphSubsResponse && graphSubsResponse.value) {
      for (const existingSubscription of graphSubsResponse.value) {
        const resourceInfo = parseCalendarSubscriptionResource(existingSubscription.resource);
        if (!resourceInfo) {
          continue;
        }

        if (resourceInfo.userId.toLowerCase() !== selectedUserId.toLowerCase()) {
          continue;
        }

        const calendarKey = String(resourceInfo.calendarId || '').toLowerCase();
        if (!existingByCalendar.has(calendarKey)) {
          existingByCalendar.set(calendarKey, []);
        }

        existingByCalendar.get(calendarKey).push(existingSubscription.id);
      }
    }
  } catch (existingError) {
    console.log(`Unable to load existing admin subscriptions from Graph: ${existingError.message}`);
  }

  let subscribedCount = 0;
  let unsubscribedCount = 0;
  let failedCount = 0;

  for (const [calendarKey, subscriptionIds] of existingByCalendar) {
    if (selectedSet.has(calendarKey)) {
      continue;
    }

    for (const subscriptionId of subscriptionIds) {
      try {
        await client.api(`/subscriptions/${subscriptionId}`).delete();
      } catch (deleteError) {
        console.log(`Unable to delete subscription ${subscriptionId}: ${deleteError.message}`);
        failedCount += 1;
      }

      dbHelper.deleteSubscription(subscriptionId);
      unsubscribedCount += 1;
    }
  }

  for (const [calendarKey, calendarId] of selectedByLower) {
    if (existingByCalendar.has(calendarKey)) {
      continue;
    }

    try {
      const calendarName = calendarNameById.get(calendarKey) || calendarId;
      const subscription = await client.api('/subscriptions').create({
        changeType: 'created',
        notificationUrl: `${notificationHost}/listen`,
        lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
        resource: `users/${selectedUserId}/calendars/${calendarId}/events?$select=id,subject,start,end,organizer,bodyPreview,attendees,location`,
        clientState: subscriptionStateHelper.buildClientState(
          process.env.SUBSCRIPTION_CLIENT_STATE,
          {
            userName: selectedUserName,
            calendarName: calendarName,
          },
        ),
        includeResourceData: true,
        encryptionCertificate: certHelper.getSerializedCertificate(
          process.env.CERTIFICATE_PATH,
        ),
        encryptionCertificateId: process.env.CERTIFICATE_ID,
        expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
      });

      dbHelper.addSubscription(subscription.id, 'APP-ONLY-USER-FLOW');
      subscribedCount += 1;
    } catch (createError) {
      console.log(`Unable to create admin subscription for calendar ${calendarId}: ${createError.message}`);
      failedCount += 1;
    }
  }

  res.redirect(
    `/admin-flow?userId=${encodeURIComponent(selectedUserId)}&subscribed=${subscribedCount}&unsubscribed=${unsubscribedCount}&failed=${failedCount}`,
  );
});

router.post('/admin-flow/unsubscribe', async function (req, res) {
  const subscriptionId = String(req.body.subscriptionId || '').trim();
  const selectedUserId = String(req.body.userId || '').trim();

  if (!subscriptionId) {
    res.redirect(selectedUserId ? `/admin-flow?userId=${encodeURIComponent(selectedUserId)}` : '/admin-flow');
    return;
  }

  try {
    const client = graph.getGraphClientForApp(req.app.locals.msalClient);
    await client.api(`/subscriptions/${subscriptionId}`).delete();
  } catch (error) {
    console.log(`Unable to delete subscription ${subscriptionId} from Graph: ${error.message}`);
  }

  // Always remove local record to avoid stale IDs in memory.
  dbHelper.deleteSubscription(subscriptionId);

  res.redirect(selectedUserId ? `/admin-flow?userId=${encodeURIComponent(selectedUserId)}` : '/admin-flow');
});

export default router;
