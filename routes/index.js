import express from "express";
import graph from "../helpers/graphHelper.js";

const router = express.Router();

// GET /
router.get('/', async function (req, res, next) {
  const users = [];

  try {
    const client = graph.getGraphClientForApp(req.app.locals.msalClient);
    const response = await client
      .api('/users')
      .select('id,displayName,mail,userPrincipalName')
      .top(25)
      .get();

    if (response?.value) {
      for (const user of response.value) {
        users.push({
          id: user.id,
          name: user.displayName || user.userPrincipalName || user.mail || user.id,
          email: user.mail || user.userPrincipalName || '',
        });
      }
    }
  } catch (error) {
    console.log(`Unable to load users for app-only subscribe: ${error.message}`);
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
    selectedUserId: process.env.USER_ID || '',
  });
});

export default router;
