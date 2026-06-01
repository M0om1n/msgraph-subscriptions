// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
const router = express.Router();

import graph from "../helpers/graphHelper.js";
import certHelper from "../helpers/certHelper.js";
import dbHelper from "../helpers/dbHelper.js";

// GET /apponly/subscribe
router.get('/subscribe', async function (req, res) {
  const client = graph.getGraphClientForApp(req.app.locals.msalClient);
  const selectedUserId = req.query.userId || process.env.USER_ID;
  const selectedCalendarId = req.query.calendarId || '';

  // In production, use the current host to receive notifications
  const notificationHost = `https://${req.hostname}`;

  try {
    if (!selectedUserId) {
      throw new Error('No user ID was provided and USER_ID is not configured.');
    }

    const selectedUser = await client
      .api(`/users/${selectedUserId}`)
      .select('id,displayName,mail,userPrincipalName')
      .get();

    let selectedCalendar = null;
    if (selectedCalendarId) {
      selectedCalendar = await client
        .api(`/users/${selectedUserId}/calendars/${selectedCalendarId}`)
        .select('id,name')
        .get();
    }

    const existingSubscriptions = dbHelper.getSubscriptionsByUserAccountId('APP-ONLY');

    // Apps are only allowed one subscription to the /users/{id}/events resource
    // If we already had one, delete it so we can create a new one
    if (existingSubscriptions) {
      for (var existingSub of existingSubscriptions) {
        try {
          await client
            .api(`/subscriptions/${existingSub}`)
            .delete();
        } catch (err) {
          console.error(err);
        }

        dbHelper.deleteSubscription(existingSub);
      }
    }

    const subscribedResource = selectedCalendarId
      ? `users/${selectedUserId}/calendars/${selectedCalendarId}/events?$select=id,subject,start,end,organizer,bodyPreview,attendees,location`
      : `users/${selectedUserId}/events?$select=id,subject,start,end,organizer,bodyPreview,attendees,location`;

    // Create the subscription
    const subscription = await client.api('/subscriptions').create({
      changeType: 'created',
      notificationUrl: `${notificationHost}/listen`,
      lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
      resource: subscribedResource,
      clientState: process.env.SUBSCRIPTION_CLIENT_STATE,
      includeResourceData: true,
      // To get resource data, we must provide a public key that
      // Microsoft Graph will use to encrypt their key
      // See https://learn.microsoft.com/graph/webhooks-with-resource-data#creating-a-subscription
      encryptionCertificate: certHelper.getSerializedCertificate(
        process.env.CERTIFICATE_PATH,
      ),
      encryptionCertificateId: process.env.CERTIFICATE_ID,
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
    });

    // Save the subscription ID in the session
    req.session.subscriptionId = subscription.id;
    req.session.appOnlyUser = {
      id: selectedUser.id,
      name:
        selectedUser.displayName ||
        selectedUser.userPrincipalName ||
        selectedUser.mail ||
        selectedUser.id,
      email: selectedUser.mail || selectedUser.userPrincipalName || '',
    };
    req.session.appOnlyCalendar = selectedCalendar
      ? {
        id: selectedCalendar.id,
        name: selectedCalendar.name || selectedCalendar.id,
      }
      : null;
    console.log(
      `Subscribed to ${selectedCalendarId ? `calendar ${selectedCalendarId}` : 'primary calendar'} for user ${selectedUserId}, subscription ID: ${subscription.id}`,
    );

    // Add subscription to the database
    dbHelper.addSubscription(subscription.id, 'APP-ONLY');

    // Redirect to subscription page
    res.redirect('/watch');
  } catch (error) {
    req.flash('error_msg', {
      message: 'Error subscribing for user calendar event notifications',
      debug: JSON.stringify(error, Object.getOwnPropertyNames(error)),
    });

    res.redirect('/');
  }
});

// GET /apponly/signout
router.get('/signout', async function (req, res) {
  // Delete the subscription from database and Graph
  const subscriptionId = req.session.subscriptionId;
  const msalClient = req.app.locals.msalClient;

  dbHelper.deleteSubscription(subscriptionId);

  const client = graph.getGraphClientForApp(msalClient);

  try {
    await client.api(`/subscriptions/${subscriptionId}`).delete();

    req.session.subscriptionId = null;
    req.session.appOnlyUser = null;
    req.session.appOnlyCalendar = null;
  } catch (graphErr) {
    console.log(`Error deleting subscription from Graph: ${graphErr.message}`);
  }

  res.redirect('/');
});

export default router;
