// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
const router = express.Router();

import graph from "../helpers/graphHelper.js";
import certHelper from "../helpers/certHelper.js";
//const dbHelper = require('../helpers/dbHelper');

//create sertificate if it doesn't exist (testing)
/*
router.get('/subscribe', async function (req, res) {
  const client = graph.getGraphClientForApp(req.app.locals.msalClient);

  // Ensure a certificate exists
  await certHelper.createSelfSignedCertificateIfNotExists(
    process.env.CERTIFICATE_PATH,
    process.env.PRIVATE_KEY_PATH,
    process.env.PRIVATE_KEY_PASSWORD,
  );

  res.redirect('/');
});
*/

// GET /apponly/subscribe
router.get('/subscribe', async function (req, res) {
  const client = graph.getGraphClientForApp(req.app.locals.msalClient);

  // Ensure a certificate exists
  await certHelper.createSelfSignedCertificateIfNotExists(
    process.env.CERTIFICATE_PATH,
    process.env.PRIVATE_KEY_PATH,
    process.env.PRIVATE_KEY_PASSWORD,
  );

  // In production, use the current host to receive notifications
  const notificationHost = `${req.protocol}://${req.hostname}`;

  try {
    /*    
    const existingSubscriptions =
      await dbHelper.getSubscriptionsByUserAccountId('APP-ONLY');

    // Apps are only allowed one subscription to the /teams/getAllMessages resource
    // If we already had one, delete it so we can create a new one
    if (existingSubscriptions) {
      for (var existingSub of existingSubscriptions) {
        try {
          await client
            .api(`/subscriptions/${existingSub.subscriptionId}`)
            .delete();
        } catch (err) {
          console.error(err);
        }

        await dbHelper.deleteSubscription(existingSub.subscriptionId);
      }
    }
    */  

    // Create the subscription
    const subscription = await client.api('/subscriptions').create({
      changeType: 'created',
      notificationUrl: `${notificationHost}/listen`,
      lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
      resource: '/teams/getAllMessages',
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
    console.log(
      `Subscribed to Teams channel messages, subscription ID: ${subscription.id}`,
    );

    // Add subscription to the database
    await dbHelper.addSubscription(subscription.id, 'APP-ONLY');

    // Redirect to subscription page
    res.redirect('/watch');
  } catch (error) {
    req.flash('error_msg', {
      message: 'Error subscribing for Teams channel message notifications',
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

  //await dbHelper.deleteSubscription(subscriptionId);

  const client = graph.getGraphClientForApp(msalClient);

  try {
    await client.api(`/subscriptions/${subscriptionId}`).delete();

    req.session.subscriptionId = null;
  } catch (graphErr) {
    console.log(`Error deleting subscription from Graph: ${graphErr.message}`);
  }

  res.redirect('/');
});

export default router;
