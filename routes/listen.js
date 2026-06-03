// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";

const router = express.Router();

import graph from "../helpers/graphHelper.js";
import certHelper from "../helpers/certHelper.js";
import tokenHelper from "../helpers/tokenHelper.js";
import dbHelper from "../helpers/dbHelper.js";
import subscriptionStateHelper from "../helpers/subscriptionStateHelper.js";

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
      if (subscriptionStateHelper.isExpectedClientState(
        notification.clientState,
        process.env.SUBSCRIPTION_CLIENT_STATE,
      )) {
        // Verify we have a matching subscription record in the database
        const subscription = dbHelper.getSubscription(
          notification.subscriptionId
        );
        if (subscription) {
          console.log(`Received notification for subscription ${notification.subscriptionId}`);

          // If notification has encrypted content, process that
          if (notification.encryptedContent) {
            await processEncryptedNotification(
              notification,
              req.app.locals.wss,
              req.app.locals.msalClient,
            );
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

/**
 * Processes an encrypted notification
 * @param  {object} notification - The notification containing encrypted content
 * @param  {WebSocket.Server} wss - The WebSocket server instance
 * @param  {IConfidentialClientApplication} msalClient - The MSAL client to retrieve app-only tokens for Graph requests
 */
async function processEncryptedNotification(notification, wss, msalClient) {
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

    const decryptedResource = JSON.parse(decryptedPayload);
    const eventWithBody = await getAppEventWithBody(
      notification,
      decryptedResource,
      msalClient,
    );

    // Send the notification to the WebSocket
    emitNotification(notification.subscriptionId, {
      type: 'app_event',
      resource: eventWithBody,
    }, wss);
  }
}

/**
 * Fetches only event body for app-only notifications
 * @param  {object} notification - The webhook notification
 * @param  {object} decryptedResource - The decrypted event payload
 * @param  {IConfidentialClientApplication} msalClient - The MSAL client to retrieve app-only tokens for Graph requests
 * @returns {object} Decrypted payload enriched with body when available
 */
async function getAppEventWithBody(notification, decryptedResource, msalClient) {
  const eventPath = getEventPathFromNotification(
    notification,
    decryptedResource ? decryptedResource.id : null,
  );

  if (!eventPath) {
    return decryptedResource;
  }

  try {
    const client = graph.getGraphClientForApp(msalClient);
    const event = await client
      .api(eventPath)
      .select('body')
      .get();

    return {
      ...decryptedResource,
      body: event && event.body ? event.body : null,
    };
  } catch (err) {
    console.log(`Error getting app-only event from ${eventPath}:`);
    console.error(err);
    return decryptedResource;
  }
}

/**
 * Gets an event API path from a notification
 * @param  {object} notification - The webhook notification
 * @param  {string} eventId - The event ID
 * @returns {string | null} A Graph API path for the event
 */
function getEventPathFromNotification(notification, eventId) {
  const resourceData = notification ? notification.resourceData : null;
  const odataId = resourceData ? resourceData['@odata.id'] : null;
  if (odataId) {
    return normalizeGraphPath(odataId);
  }

  const resource = notification && notification.resource
    ? normalizeGraphPath(notification.resource)
    : null;

  if (!resource) {
    return null;
  }

  const lowerResource = resource.toLowerCase();
  if (lowerResource.includes('/events/')) {
    return resource;
  }

  if (eventId && lowerResource.endsWith('/events')) {
    return `${resource}/${eventId}`;
  }

  return null;
}

/**
 * Normalizes Graph resource paths from webhook payloads
 * @param  {string} rawPath - The raw Graph path
 * @returns {string} A path suitable for Graph client .api()
 */
function normalizeGraphPath(rawPath) {
  const basePath = String(rawPath || '').split('?')[0].trim();
  if (!basePath) {
    return '';
  }

  return basePath.startsWith('/') ? basePath : `/${basePath}`;
}

/**
 * Process a non-encrypted notification
 * @param  {object} notification - The notification to process
 * @param  {IConfidentialClientApplication} msalClient - The MSAL client to retrieve tokens for Graph requests
 * @param  {string} userAccountId - The user's account ID
 * @param  {WebSocket.Server} wss - The WebSocket server instance
 */
async function processNotification(notification, msalClient, userAccountId, wss) {
  // Get the event ID
  const eventId = notification.resourceData.id;
  const client = graph.getGraphClientForUser(msalClient, userAccountId);

  try {
    // Get the event from Graph
    const event = await client
      .api(`/me/events/${eventId}`)
      .select('id,subject,start,end,organizer,body,bodyPreview,attendees,location')
      .get();

    // Send the notification to the WebSocket
    emitNotification(notification.subscriptionId, {
      type: 'user_event',
      resource: event,
    }, wss);
  } catch (err) {
    console.log(`Error getting event with ${eventId}:`);
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
