//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import debug = require('debug');

const dbg = debug('health');

let mapTypeToValue = {
  readiness: 'ready',
  liveness: 'healthy',
};

export default function initializeHealthCheck(app, config) {
  let configuredHealthDelays = {
    readiness: 0,
    liveness: 0,
  };
  function checkHealth(checkType) {
    const started = app.settings.started;
    const startupSeconds = configuredHealthDelays[checkType];
    if (
      configuredHealthDelays[checkType] === undefined ||
      configuredHealthDelays[checkType] === null
    ) {
      throw new Error('Invalid health check type');
    }
    const now = new Date();
    const startupMs = startupSeconds * 1000;
    if (now.getTime() - started.getTime() <= startupMs) {
      // Still in the startup period
      dbg(`Returning ${checkType} OK: within the startup delay window`);
      return true;
    }
    const healthIndicatorKey = mapTypeToValue[checkType];
    const isHealthy = !!provider[healthIndicatorKey];
    const asString = isHealthy ? 'OK' : 'FALSE';
    dbg(`Returning ${checkType} ${asString}`);
    return isHealthy;
  }

  function containerHealthCheck(checkType, validateHeader, req, res, next) {
    const header = config.containers.healthCheck.expectedHeader;
    if (validateHeader && !req.headers[header.name]) {
      dbg(
        `Container ${checkType} health check requested but the ${header.name} header was not present in the HTTP request`
      );
      return next();
    }
    if (validateHeader && !req.headers[header.name] === header.value) {
      dbg(
        `Container ${checkType} health check requested but the ${header.name} header present in the HTTP request did not match the expected, configured value`
      );
      return next();
    }
    let result = null;
    try {
      result = checkHealth(checkType);
    } catch (error) {
      error.statusCode = 500;
      return next(error);
    }
    res.status(result ? 200 : 500).end();
  }

  const provider = {
    readiness: checkHealth.bind(null, 'readiness'),
    liveness: checkHealth.bind(null, 'liveness'),
    ready: false,
    healthy: true,
  };

  if (config && config.containers && config.containers.docker) {
    configuredHealthDelays.readiness =
      config.containers.healthCheck.delay.readiness;
    configuredHealthDelays.liveness =
      config.containers.healthCheck.delay.liveness;

    app.get(
      '/health/readiness',
      containerHealthCheck.bind(null, 'readiness', true)
    );
    app.get(
      '/health/liveness',
      containerHealthCheck.bind(null, 'liveness', true)
    );
    app.get(
      '/health/external',
      containerHealthCheck.bind(
        null,
        'liveness',
        false /* do not validate a header, anyone can hit */
      )
    );
    debug('Health probes listening');
  } else {
    debug('No health probes listening');
  }

  return provider;
}
