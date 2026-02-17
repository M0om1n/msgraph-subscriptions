// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
const router = express.Router();

import graph from "../helpers/graphHelper.js";
import dbHelper from "../helpers/dbHelper.js";

// GET /delegated/signin
router.get('/signin', async function (req, res) {
  // Start the authorization code flow by redirecting the
  // browser to Microsoft identity platforms authorization URL
  const urlParameters = {
    scopes: process.env.OAUTH_SCOPES.split(','),
    redirectUri: process.env.OAUTH_REDIRECT_URI,
    prompt: 'select_account',
  };

  const authUrl = await req.app.locals.msalClient.getAuthCodeUrl(urlParameters);
  res.redirect(authUrl);
});

// GET /delegated/callback
router.get('/callback', async function (req, res) {
  // Microsoft identity platform redirects the browser here with the
  // authorization result
  const tokenRequest = {
    code: req.query.code,
    scopes: process.env.OAUTH_SCOPES.split(','),
    redirectUri: process.env.OAUTH_REDIRECT_URI,
  };

  const response = await req.app.locals.msalClient.acquireTokenByCode(tokenRequest);

  // Save the user's homeAccountId in their session
  req.session.userAccountId = response.account.homeAccountId;

  const client = graph.getGraphClientForUser(
    req.app.locals.msalClient,
    req.session.userAccountId,
  );

  try {
    // Get the user's profile from Microsoft Graph
    const user = await client.api('/me').select('displayName, mail').get();

    // Save user's name and email address in the session
    req.session.user = {
      name: user.displayName,
      email: user.mail,
    };

    console.log(`Logged in as ${user.displayName}`);

    // If in production, use the current host to receive notifications
    const notificationHost = `${req.protocol}://${req.hostname}`;

    //console.log(`notificationHost: ${notificationHost}`);

    // Create the subscription
    const subscription = await client.api('/subscriptions').create({
      changeType: 'created',
      notificationUrl: `${notificationHost}/listen`,
      lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
      resource: 'me/mailFolders/inbox/messages',
      clientState: process.env.SUBSCRIPTION_CLIENT_STATE,
      includeResourceData: false,
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
    });

    // Save the subscription ID in the session
    req.session.subscriptionId = subscription.id;
    console.log(`Subscribed to user's inbox, subscription ID: ${subscription.id}`);

    // Add the subscription to the database
    dbHelper.addSubscription(subscription.id, req.session.userAccountId);

    // Redirect to subscription page
    res.redirect('/watch');
  } catch (error) {
    return Promise.reject(new Error('Error completing authentication')); 
  }
});

// GET /delegated/signout
router.get('/signout', async function (req, res) {
  // Delete the subscription from database and Graph
  const subscriptionId = req.session.subscriptionId;
  const msalClient = req.app.locals.msalClient;

  dbHelper.deleteSubscription(subscriptionId);

  const client = graph.getGraphClientForUser(
    msalClient,
    req.session.userAccountId,
  );

  try {
    await client.api(`/subscriptions/${subscriptionId}`).delete();

    req.session.subscriptionId = null;
  } catch (graphErr) {
    console.log(`Error deleting subscription from Graph: ${graphErr.message}`);
  }

  try {
    // Remove user's account from MSAL cache
    const userAccount = await msalClient
      .getTokenCache()
      .getAccountByHomeId(req.session.userAccountId);

    await msalClient.getTokenCache().removeAccount(userAccount);

    req.session.userAccountId = null;
  } catch (msalErr) {
    console.log(`Error removing user from MSAL cache: ${msalErr.message}`);
  }

  res.redirect('/');
});

export default router;
