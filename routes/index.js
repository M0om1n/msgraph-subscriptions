import express from "express";
import graph from "../helpers/graphHelper.js";

const router = express.Router();

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

export default router;
