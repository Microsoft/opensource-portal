//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import querystring from 'querystring';
import { AxiosError } from 'axios';

import { wrapError } from '../utils';
import { getProviders } from '../transitional';

function redactRootPathsFromString(string, path) {
  if (typeof string === 'string' && string.includes && string.split) {
    return string.split(path).join('[app]');
  }
  return string;
}

function redactRootPaths(view) {
  const path = process.cwd();
  if (typeof view === 'object') {
    for (var property in view) {
      if (view.hasOwnProperty(property)) {
        var value = view[property];
        if (typeof value === 'string') {
          view[property] = redactRootPathsFromString(value, path);
        }
      }
    }
  } else if (typeof view === 'string') {
    return redactRootPathsFromString(view, path);
  }
  return view;
}

function containsNewlinesNotHtml(error) {
  if (error && error.message && error.message.includes && error.message.split) {
    var newlines = error.message.split('\n');
    return newlines.length > 3 && !error.message.includes('</');
  }
  return false;
}

const exceptionFieldsOfInterest = [
  'status',
  'statusCode',
  'innerMessage',
];

export default function SiteErrorHandler (err, req, res, next) {
  // CONSIDER: Let's eventually decouple all of our error message improvements to another area to keep the error handler intact.
  const { applicationProfile, config } = getProviders(req);
  var correlationId = req.correlationId;
  var errorStatus = err ? (err.status || err.statusCode) : undefined;
  // Per GitHub: https://developer.github.com/v3/oauth/#bad-verification-code
  // When they offer a code that another GitHub auth server interprets as invalid,
  // the app should retry.
  if ((err.message === 'The code passed is incorrect or expired.' || (err.message === 'Failed to obtain access token' && err.oauthError.message === 'The code passed is incorrect or expired.')) && req.scrubbedUrl.startsWith('/auth/github/')) {
    req.insights.trackMetric({ name: 'GitHubInvalidExpiredCodeRedirect', value: 1 });
    req.insights.trackEvent({ name: 'GitHubInvalidExpiredCodeRetry' });
    return res.redirect(req.scrubbedUrl === '/auth/github/callback/increased-scope?code=*****' ? '/auth/github/increased-scope' : '/auth/github');
  }
  const isGitHubAbuseRateLimit = err && err.message && err.message.includes && err.message.includes('#abuse-rate-limits');
  if (isGitHubAbuseRateLimit) {
    req.insights.trackMetric({ name: 'GitHubAbuseRateLimit', value: 1 });
  }
  if (err.message && err.message.includes && err.message.includes('ETIMEDOUT') && (err.message.includes('192.30.253.116') || err.message.includes('192.30.253.117'))) {
    req.insights.trackMetric({ name: 'GitHubApiTimeout', value: 1 });
    req.insights.trackEvent({ name: 'GitHubApiTimeout' });
    err = wrapError(err, 'The GitHub API is temporarily down. Please try again soon.', false);
  }
  var primaryUserInstance = req.user ? req.user.github : null;
  if (config) {
    if (config.authentication.scheme !== 'github') {
      primaryUserInstance = req.user ? req.user.azure : null;
    }
    var version = config && config.logging && config.logging.version ? config.logging.version : '?';
    if (config.logging.errors && err.status !== 403 && err.skipLog !== true) {
      let appSource = 'unknown';
      if (process.argv.length > 1) {
        appSource = process.argv.slice(1).join(' ');
      }
      const insightsProperties = {
        url: req.scrubbedUrl || req.originalUrl || req.url,
        entrypoint: appSource,
        stk: undefined,
        message: undefined,
      };
      if (req.insights && req.insights.trackException) {
        for (let i = 0; err && i < exceptionFieldsOfInterest.length; i++) {
          const key = exceptionFieldsOfInterest[i];
          const value = err[key];
          if (value && typeof value === 'number') {
            insightsProperties[key] = value.toString();
          } else if (value) {
            insightsProperties[key] = value;
          }
          try {
            // Try and store our own stack representation to compare it with
            // zone aware error stacks that App Insights produces. This is an
            // experiment at this time. (May 2018)
            if (err) {
              insightsProperties.stk = err.stack;
            }
          } catch (stackProblem) { /* ignore */ }
        }
        if (isGitHubAbuseRateLimit) {
          insightsProperties.message = err.message;
          req.insights.trackEvent({
            name: 'GitHubAbuseRateLimitError',
            properties: insightsProperties,
          });
        } else {
          if (err && err['json']) {
            // not tracking jsonErrors for now, they pollute app insights
          } else {
            req.insights.trackException({ exception: err, properties: insightsProperties });
          }
        }
      }
    }
  }
  if (err !== undefined && err.skipLog !== true) {
    console.log('Error: ' + (err && err.message ? err.message : 'Error is undefined.'));
    if (err.stack) {
      console.error(err.stack);
    }
    if (err.innerError) {
      var inner = err.innerError;
      console.log('Inner: ' + inner.message);
      if (inner.stack) {
        console.log(inner.stack);
      }
    }
  }
  // Bubble OAuth errors to the forefront... this is the rate limit scenario.
  if (err && err.oauthError && err.oauthError.statusCode && err.oauthError.statusCode && err.oauthError.data) {
    var detailed = err.message;
    err = err.oauthError;
    err.status = err.statusCode;
    var data = JSON.parse(err.data);
    if (data && data.message) {
      err.message = err.statusCode + ': ' + data.message;
    } else {
      err.message = err.statusCode + ' Unauthorized received. You may have exceeded your GitHub API rate limit or have an invalid auth token at this time.';
    }
    err.detailed = detailed;
  }
  // Don't leak the Redis connection information.
  if (err && err.message && err.message.indexOf('Redis connection') >= 0 && err.message.indexOf('ETIMEDOUT')) {
    err.message = 'The session store was temporarily unavailable. Please try again.';
  }
  if (res.headersSent) {
    console.error('Headers were already sent.');
    return next(err);
  }
  if (err && err.forceSignOut === true && req && req.logout) {
    req.logout();
  }
  var safeMessage = redactRootPaths(err.message);
  const defaultErrorTitle = err && err.skipOops ? 'FYI' : 'Oops';
  const view = {
    message: safeMessage,
    encodedMessage: querystring.escape(safeMessage),
    messageHasNonHtmlNewlines: containsNewlinesNotHtml(err),
    serviceBanner: config && config.serviceMessage ? config.serviceMessage.banner : undefined,
    detailed: err && err.detailed ? redactRootPaths(err.detailed) : undefined,
    encodedDetailed: err && err.detailed ? querystring.escape(redactRootPaths(err.detailed)) : undefined,
    errorFancyLink: err && err.fancyLink ? err.fancyLink : undefined,
    errorFancySecondaryLink: err && err.fancySecondaryLink ? err.fancySecondaryLink : undefined,
    errorStatus: errorStatus,
    skipLog: err.skipLog,
    skipOops: err && err.skipOops ? err.skipOops : false,
    error: {},
    title: err.title || (err.status === 404 ? 'Not Found' : defaultErrorTitle),
    primaryUser: primaryUserInstance,
    user: req.user,
    config: config && config.obfuscatedConfig ? config.obfuscatedConfig : null,
  };

  // Depending on the library in use, we get everything from non-numeric textual status
  // descriptions to status codes as strings and more. Set the status code found in
  // the error if we have it.
  var errStatusAsNumber = null;
  if (err.status) {
    errStatusAsNumber = parseInt(err.status);
  }
  let resCode = errStatusAsNumber || (err.status && typeof (err.status) === 'number' ? err.status : false) || err.statusCode || 500;
  if (err && err.isAxiosError) {
    const axiosError = err as AxiosError;
    if (axiosError?.response?.status) {
      resCode = axiosError.response.status;
    }
  }
  res.status(resCode);

  // Support JSON-based error display for the API route, showing just a small
  // subset of typical view properties to share from the error instance.
  if (err && err.json === true) {
    const safeError = {
      message: safeMessage,
      correlationId: correlationId,
      documentation_url: undefined,
    };
    if (err.documentation_url) {
      safeError.documentation_url = err.documentation_url;
    }
    const fieldsOfInterest = ['serviceBanner', 'detailed'];
    fieldsOfInterest.forEach((fieldName) => {
      if (view[fieldName]) {
        safeError[fieldName] = view[fieldName];
      }
    });
    res.json(safeError);
  } else {
    if (!applicationProfile.customErrorHandlerRender) {
      return res.render('error', view);
    }
    return applicationProfile.customErrorHandlerRender(view, err, req, res, next).then(ok => {
      // done
    }).catch(error => {
      console.error(error);
      res.end();
    });
  }
};
