import express from "express";
import createError from "http-errors";
import path from "path";
import { fileURLToPath } from "url";
import cookieParser from "cookie-parser";
import logger from "morgan";
import flash from "connect-flash";
import session from "express-session";
import * as msal from "@azure/msal-node";
import dotenv from "dotenv";

import indexRouter from "./routes/index.js";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

// MSAL config
const msalConfig = {
  auth: {
    clientId: process.env.OAUTH_CLIENT_ID,
    authority: `${process.env.OAUTH_AUTHORITY}/${process.env.OAUTH_TENANT_ID}`,
    clientSecret: process.env.OAUTH_CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(logLevel, message, containsPii) {
        if (!containsPii) console.log(`msal: [${logLevel}] ${message}`);
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Error,
    },
  },
};

// Create msal application object
app.locals.msalClient = new msal.ConfidentialClientApplication(msalConfig);

// Session middleware
app.use(
  session({
    secret: process.env.EXPRESS_SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    unset: 'destroy',
  }),
);

// Flash middleware
app.use(flash());
app.use(function (req, res, next) {
  // Read any flashed errors and save in the response locals
  res.locals.errors = req.flash('error_msg');

  // Check for simple error string and convert to layout's expected format
  const errs = req.flash('error');
  for (const err in errs) {
    res.locals.errors.push({ message: 'An error occurred', debug: err });
  }

  next();
});

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'pug');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', indexRouter);

// catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

const port = process.env.PORT || 3001;

const server = app.listen(port, () => console.log(`App listening on port ${port}!`));

server.keepAliveTimeout = 120 * 1000;
server.headersTimeout = 120 * 1000;
