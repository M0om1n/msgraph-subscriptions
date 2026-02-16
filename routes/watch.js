// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from "express";
const router = express.Router();

import graph from "../helpers/graphHelper.js";

// GET /watch
router.get('/', async function (req, res) {
  const userAccountId = req.session.userAccountId;
  const subscriptionId = req.session.subscriptionId;
  const user = req.session.user;

  if (!subscriptionId) {
    res.redirect('/');
    return;
  }

  // If there is a user account ID in the session, assume
  // we're watching notifications in a user's mailbox
  const userMode = userAccountId ? true : false;

  const client = userMode
    ? graph.getGraphClientForUser(req.app.locals.msalClient, userAccountId)
    : graph.getGraphClientForApp(req.app.locals.msalClient);

  // Get the subscription details to display on the page
  const subscription = await client
    .api(`/subscriptions/${subscriptionId}`)
    .get();

  res.render('watch', {
    subscription: subscription,
    user: user,
    userMode: userMode,
  });
});

export default router;
