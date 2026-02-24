// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
import PostalMime from "postal-mime";

const router = express.Router();

import graph from "../helpers/graphHelper.js";
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

  console.log(`--------- Received notification ---------`);
  console.log(JSON.stringify(req.body, null, 2));
  console.log(`-----------------------------------------`);

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
            await extractBodyAndAttachments(notification, req.app.locals.msalClient);
            processEncryptedNotification(notification, req.app.locals.wss);
          } else {
            await processNotification(
              notification,
              req.app.locals.msalClient,
              subscription.userAccountId,
              req.app.locals.wss
            );
          }
        }
      }
    }
  }

  res.status(202).end();
});

async function extractBodyAndAttachments(notification, msalClient) {
  const client = graph.getGraphClientForApp(msalClient);
  const messageId = notification.resourceData.id;

  try {
    // Get the eml content from Graph
    const eml = await client
      .api(`/users/${process.env.USER_ID}/messages/${messageId}/$value`)
      .get();
    const email = await PostalMime.parse(eml);

    console.log(`Extracted html body: ${email.html}`);

    if (email.attachments.length > 0) {
      console.log(`Attachments:`);
      email.attachments.forEach((attachment) => {
        console.log(`- ${attachment.filename} (${attachment.contentType})`);
      });
    }   

  } catch (err) {
    console.log(`Error getting eml content with ${messageId}:`);
    console.error(err);
  }
}  

/**
 * Processes an encrypted notification
 * @param  {object} notification - The notification containing encrypted content
 */
function processEncryptedNotification(notification, wss) {
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

    // Send the notification to the WebSocket
    emitNotification(notification.subscriptionId, {
      type: 'message',
      resource: JSON.parse(decryptedPayload),
    }, wss);
  }
}

/**
 * Process a non-encrypted notification
 * @param  {object} notification - The notification to process
 * @param  {IConfidentialClientApplication} msalClient - The MSAL client to retrieve tokens for Graph requests
 * @param  {string} userAccountId - The user's account ID
 * @param  {WebSocket.Server} wss - The WebSocket server instance
 */
async function processNotification(notification, msalClient, userAccountId, wss) {
  // Get the message ID
  const messageId = notification.resourceData.id;
  const client = graph.getGraphClientForUser(msalClient, userAccountId);

  try {
    // Get the message from Graph
    const message = await client
      .api(`/me/messages/${messageId}`)
      .select('subject,id')
      .get();

    // Send the notification to the WebSocket
    emitNotification(notification.subscriptionId, {
      type: 'user_message',
      resource: message,
    }, wss);
  } catch (err) {
    console.log(`Error getting message with ${messageId}:`);
    console.error(err);
  }
}

/**
 * Sends a notification
 * @param  {string} subscriptionId - The subscription ID used to send to the correct room
 * @param  {object} data - The data to send to the room
 * @param  {WebSocket.Server} wss - The WebSocket server instance
 */
function emitNotification(subscriptionId, data, wss) {
  console.log(`Emitting notification client ${subscriptionId}: ${JSON.stringify(data)}`);
  // Send the notification to the WebSocket
  wss.clients.forEach((client) => {
    console.log(`Client processing...`); 

    if (client.readyState === WebSocket.OPEN /*&& client.subscriptionId === subscriptionId*/) {
      client.send(JSON.stringify(data));
    }
  });
}

export default router;
