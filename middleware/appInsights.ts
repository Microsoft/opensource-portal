//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import wrapOrCreateInsightsConsoleClient from '../lib/insights';

const debug = require('debug')('startup');

import { setup as appInsightsSetup, defaultClient } from 'applicationinsights';
import { IReposApplication } from '../app';
import { IProviders, ReposAppRequest } from '../transitional';

function ignoreKubernetesProbes(envelope /* , context */) {
  if ('RequestData' === envelope.data.baseType) {
    const data = envelope.data;
    if (
      data.baseData.name.startsWith &&
      data.baseData.name.startsWith('GET /health/')
    ) {
      // Do not log any telemetry for k8s and health probes
      return false;
    }
  }
  return true;
}

function filterTelemetry(envelope, context): boolean {
  const { data } = envelope;
  if (
    data &&
    data.baseType === 'RequestData' &&
    data.baseData.responseCode === '401'
  ) {
    // We believe 401 is successful, not a failure
    data.baseData.success = true;
  } else if (
    data &&
    data.baseData &&
    data.baseData.name &&
    data.baseData.responseCode &&
    data.baseData.responseCode === '404'
  ) {
    if (data.baseData.name.startsWith('GET /api/')) {
      // Link lookup APIs return 404 by design, which is a success.
      data.baseData.success = true;
    }
  }
  return true;
}

export default function initializeAppInsights(app: IReposApplication, config) {
  let client = undefined;
  if (!config) {
    // Configuration failure happened ahead of this module
    return;
  }
  const providers = app.settings.providers as IProviders;
  let key: string =
    config.telemetry && config.telemetry.applicationInsightsKey
      ? config.telemetry.applicationInsightsKey
      : null;
  // Override the key with a job-specific one if this is a job execution instead
  if (
    config.telemetry &&
    config.telemetry.jobsApplicationInsightsKey &&
    config.isJobInternal === true
  ) {
    key = config.telemetry.jobsApplicationInsightsKey;
  }
  if (key) {
    const instance = providers.applicationProfile.logDependencies
      ? appInsightsSetup(key)
      : appInsightsSetup(key).setAutoCollectDependencies(false);
    defaultClient.addTelemetryProcessor(ignoreKubernetesProbes);
    defaultClient.addTelemetryProcessor(filterTelemetry);
    instance.start();
    client = defaultClient;
    debug(`insights telmetry will use identifier: ${key.substr(0, 6)}`);
  } else {
    debug(`insights telmetry is not configured with a key`);
  }

  app.use((req: ReposAppRequest, res, next) => {
    // Acknowledge synthetic tests immediately without spending time in more middleware
    if (
      req.headers &&
      req.headers['synthetictest-id'] !== undefined &&
      req.headers['x-ms-user-agent'] !== undefined &&
      req.headers['x-ms-user-agent'].includes('System Center')
    ) {
      return res.status(204).send();
    }

    // Provide application insight event tracking with correlation ID
    const extraProperties = {
      correlationId: req.correlationId,
    };
    req.insights = wrapOrCreateInsightsConsoleClient(extraProperties, client);
    next();
  });

  return wrapOrCreateInsightsConsoleClient({}, client);
}
