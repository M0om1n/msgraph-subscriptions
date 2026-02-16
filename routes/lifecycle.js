// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
const router = express.Router();

import graph from "../helpers/graphHelper.js";
//const dbHelper = require('../helpers/dbHelper');

// POST /lifecycle
router.post('/', async function (req, res) {
  // This is the notification endpoint Microsoft Graph sends notifications to

  // If there is a validationToken parameter
  // in the query string, this is the endpoint validation
  // request sent by Microsoft Graph. Return the token
  // as plain text with a 200 response
  // https://learn.microsoft.com/graph/webhooks#notification-endpoint-validation
  if (req.query && req.query.validationToken) {
    res.set('Content-Type', 'text/plain');
    res.send(req.query.validationToken);
    return;
  }

  console.log(JSON.stringify(req.body, null, 2));

  for (let i = 0; i < req.body.value.length; i++) {
    const notification = req.body.value[i];

    // Verify the client state matches the expected value
    // and that this is a lifecycle notification
    if (
      notification.clientState === process.env.SUBSCRIPTION_CLIENT_STATE &&
      notification.lifecycleEvent === 'reauthorizationRequired'
    ) {
      // Verify we have a matching subscription record in the database
      /*
      const subscription = await dbHelper.getSubscription(
        notification.subscriptionId,
      );
      */
      //For test purposes, we will assume the subscription is valid
      const subscription = notification.subscriptionId;
      
      if (subscription) {
        // Renew the subscription
        await renewSubscription(subscription, req.app.locals.msalClient);
      }
    }
  }

  res.status(202).end();
});

/**
 * Process a non-encrypted notification
 * @param  {object} subscription - The subscription to renew
 * @param  {IConfidentialClientApplication} msalClient - The MSAL client to retrieve tokens for Graph requests
 */
async function renewSubscription(subscription, msalClient) {
  // Get the Graph client
  const client =
    subscription.userAccountId === 'APP-ONLY'
      ? graph.getGraphClientForApp(msalClient)
      : graph.getGraphClientForUser(msalClient, subscription.userAccountId);

  try {
    // Update the expiration on the subscription
    await client.api(`/subscriptions/${subscription.subscriptionId}`).update({
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
    });
    console.log(`Renewed subscription ${subscription.subscriptionId}`);
  } catch (err) {
    console.log(`Error updating subscription ${subscription.subscriptionId}:`);
    console.error(err);
  }
}

export default router;
