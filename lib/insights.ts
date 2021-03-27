//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import Debug from 'debug';
const debug = Debug('appinsights');

// This file was originally designed to wrap the pre-1.0.0 version of applicationinsights,
// and so is less important today.

import type { TelemetryClient } from 'applicationinsights';

function createWrappedClient(
  propertiesToInsert: any,
  client: TelemetryClient
): TelemetryClient {
  let c = client;
  if (client) {
    client.commonProperties = propertiesToInsert;
  } else {
    c = {
      trackEvent: consoleHandler,
      trackException: consoleHandler,
      trackMetric: consoleMetric,
      trackTrace: consoleHandler,
      trackDependency: consoleHandler,
      flush: (options) => {
        options = options || {};
        if (options.callback) {
          return (options.callback as any)();
        }
      },
    } as TelemetryClient;
  }
  return c;
}

const consoleHandler = (eventNameOrProperties) => {
  eventNameOrProperties = eventNameOrProperties || {
    name: 'Unknown event, may be from pre-v1.0.0 applicationinsights',
  };
  let props = '';
  if (eventNameOrProperties && eventNameOrProperties.properties) {
    props = ' ';
    for (let [key, value] of Object.entries(eventNameOrProperties.properties)) {
      props += `${key}=${value} `;
    }
  }
  debug(
    (typeof eventNameOrProperties === 'string'
      ? eventNameOrProperties
      : eventNameOrProperties.name) + props
  );
};
const consoleMetric = (eventNameOrProperties) => {
  if (typeof eventNameOrProperties === 'string') {
    debug(
      `Legacy applicationinsights Metric ${eventNameOrProperties} was not recorded`
    );
  } else {
    eventNameOrProperties = eventNameOrProperties || {
      name: 'UnknownMetric',
      value: 0,
    };
    debug(
      `Metric(${eventNameOrProperties.name}: ${eventNameOrProperties.value}`
    );
  }
};

export default createWrappedClient;
