import express from "express";

const router = express.Router();

// GET /
router.get('/', function (req, res, next) {
  res.render('index', { title: 'Microsoft Graph Notifications Sample' });
});

export default router;
