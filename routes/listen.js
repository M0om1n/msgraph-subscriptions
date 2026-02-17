// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
const router = express.Router();

import graph from "../helpers/graphHelper.js";
import ioServer from "../helpers/socketHelper.js";
import certHelper from "../helpers/certHelper.js";
import tokenHelper from "../helpers/tokenHelper.js";
import dbHelper from "../helpers/dbHelper.js";

// POST /listen
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

  // Check for validation tokens, validate them if present
  let areTokensValid = true;
  if (req.body.validationTokens) {
    const appId = process.env.OAUTH_CLIENT_ID;
    const tenantId = process.env.OAUTH_TENANT_ID;
    const validationResults = await Promise.all(
      req.body.validationTokens.map((token) =>
        tokenHelper.isTokenValid(token, appId, tenantId),
      ),
    );

    areTokensValid = validationResults.reduce((x, y) => x && y);
  }

  if (areTokensValid) {
    for (let i = 0; i < req.body.value.length; i++) {
      const notification = req.body.value[i];

      // Verify the client state matches the expected value
      if (notification.clientState == process.env.SUBSCRIPTION_CLIENT_STATE) {
        // Verify we have a matching subscription record in the database
        const subscription = dbHelper.getSubscription(
          notification.subscriptionId
        );
        if (subscription) {
          console.log(`Received notification for subscription ${notification.subscriptionId}`);

          // If notification has encrypted content, process that
          if (notification.encryptedContent) {
            processEncryptedNotification(notification);
          } else {
            await processNotification(
              notification,
              req.app.locals.msalClient,
              subscription.userAccountId,
            );
          }
        }
      }
    }
  }

  res.status(202).end();
});

/**
 * Processes an encrypted notification
 * @param  {object} notification - The notification containing encrypted content
 */
function processEncryptedNotification(notification) {
  // Decrypt the symmetric key sent by Microsoft Graph
  const symmetricKey = certHelper.decryptSymmetricKey(
    notification.encryptedContent.dataKey,
    process.env.PRIVATE_KEY_PATH,
  );

  // Validate the signature on the encrypted content
  const isSignatureValid = certHelper.verifySignature(
    notification.encryptedContent.dataSignature,
    notification.encryptedContent.data,
    symmetricKey,
  );

  if (isSignatureValid) {
    // Decrypt the payload
    const decryptedPayload = certHelper.decryptPayload(
      notification.encryptedContent.data,
      symmetricKey,
    );

    // Send the notification to the Socket.io room
    emitNotification(notification.subscriptionId, {
      type: 'chatMessage',
      resource: JSON.parse(decryptedPayload),
    });
  }
}

/**
 * Process a non-encrypted notification
 * @param  {object} notification - The notification to process
 * @param  {IConfidentialClientApplication} msalClient - The MSAL client to retrieve tokens for Graph requests
 * @param  {string} userAccountId - The user's account ID
 */
async function processNotification(notification, msalClient, userAccountId) {
  // Get the message ID
  const messageId = notification.resourceData.id;

  const client = graph.getGraphClientForUser(msalClient, userAccountId);

  try {
    // Get the message from Graph
    const message = await client
      .api(`/me/messages/${messageId}`)
      .select('subject,id')
      .get();

    // Send the notification to the Socket.io room
    emitNotification(notification.subscriptionId, {
      type: 'message',
      resource: message,
    });
  } catch (err) {
    console.log(`Error getting message with ${messageId}:`);
    console.error(err);
  }
}

/**
 * Sends a notification to a Socket.io room
 * @param  {string} subscriptionId - The subscription ID used to send to the correct room
 * @param  {object} data - The data to send to the room
 */
function emitNotification(subscriptionId, data) {
  ioServer.to(subscriptionId).emit('notification_received', data);
}

export default router;
