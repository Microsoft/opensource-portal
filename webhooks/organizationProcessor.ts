//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "log", "dir"] }] */

import _ from 'lodash';

const crypto = require('crypto');
const secureCompare = require('secure-compare');

import { Operations } from '../business';
import { Organization } from '../business';

import Tasks from './tasks';
import { sleep } from '../utils';

interface IValidationError extends Error {
  statusCode?: number;
  computedHash?: string;
}

export abstract class WebhookProcessor {
  abstract filter(data: any): boolean;
  abstract run(
    operations: Operations,
    organization: Organization,
    data: any
  ): Promise<boolean>;
}

export interface IOrganizationWebhookEvent {
  body: any;
  rawBody?: any;
  properties: IGitHubWebhookProperties;
}

export interface IGitHubWebhookProperties {
  delivery: string;
  signature: string;
  event: string;
  started: string; // Date UTC string
}

export interface IProcessOrganizationWebhookOptions {
  operations: Operations;
  organization: Organization;
  event: IOrganizationWebhookEvent;
  acknowledgeValidEvent?: any;
}

export default async function ProcessOrganizationWebhook(
  options: IProcessOrganizationWebhookOptions
): Promise<any> {
  const operations = options.operations;
  if (!operations) {
    throw new Error('No operations instance provided');
  }
  const organization = options.organization;
  const event = options.event;
  if (!organization || !organization.name) {
    throw new Error('Missing organization instance');
  }
  if (!organization.active) {
    console.log(`inactive or unadopted organization ${organization.name}`);
    if (options.acknowledgeValidEvent) {
      options.acknowledgeValidEvent();
    }
    return;
  }
  if (!event) {
    throw new Error('Missing event');
  }
  if (!event.body) {
    throw new Error('Missing event body');
  }
  const body = event.body;
  const rawBody = event.rawBody || JSON.stringify(body);
  const properties = event.properties;
  if (
    !properties ||
    !properties.delivery ||
    !properties.signature ||
    !properties.event
  ) {
    if (options.acknowledgeValidEvent) {
      options.acknowledgeValidEvent();
    }
    throw new Error(
      'Missing event properties - delivery, signature, and/or event'
    );
  }
  // try {
  //   await verifySignatures(properties.signature, organization.webhookSharedSecrets, rawBody);
  // } catch (validationError) {
  // NO LONGER VALIDATING SIG
  // if (validationError) {
  //   if (operations && operations.insights) {
  //     const possibleOrganization = body && body.organization ? body.organization.login : 'unknown-org';
  //     console.warn(`incorrect hook signature - ${possibleOrganization} organization`);
  //     operations.insights.trackMetric({ name: 'WebhookIncorrectSecrets', value: 1 });
  //     operations.insights.trackEvent({
  //       name: 'WebhookIncorrectSecret',
  //       properties: {
  //         org: possibleOrganization,
  //         delivery: properties.delivery,
  //         event: properties.event,
  //         signature: properties.signature,
  //         approximateTime: properties.started.toISOString(),
  //         computedHash: validationError.computedHash,
  //       },
  //     });
  //   }
  //   return callback(validationError);
  // }
  //}

  // In a bus scenario, if a short timeout window is used for queue
  // visibility, a client may want to acknowledge this being a valid
  // event at this time. After this point however there is no
  // guarantee of successful execution.
  if (options.acknowledgeValidEvent) {
    options.acknowledgeValidEvent();
  }
  let interestingEvents = 0;
  const work = Tasks.filter((task) => task.filter(event));
  if (work.length > 0) {
    ++interestingEvents;
    console.log(
      `[* interesting event found: ${event.properties.event} (${work.length} interested tasks)]`
    );
  } else {
    console.log(`[skipping event: ${event.properties.event}]`);
  }

  for (let processor of work) {
    try {
      await processor.run(operations, organization, event);
    } catch (processInitializationError) {
      if (processInitializationError.status === 403) {
        console.log(`403: ${processInitializationError}`);
        if (processInitializationError.headers) {
          const headers = processInitializationError.headers;
          const rateLimit = headers['x-ratelimit-limit'];
          const rateLimitRemaining = headers['x-ratelimit-remaining'];
          const rateLimitReset = headers['x-ratelimit-reset'];
          if (rateLimit !== undefined) {
            console.log(
              `rate limit=${rateLimit}, remaining=${rateLimitRemaining}`
            );
          }
          if (rateLimitReset) {
            const resetValue = Number(rateLimitReset);
            const resetDate = new Date(1000 * resetValue);
            const now = new Date();
            if (resetDate > now) {
              const difference = resetDate.getTime() - now.getTime();
              console.log(
                `[rate limit sleep] This thread will sleep for the remainder of this limit, ${difference}ms, until ${resetDate}`
              );
              await sleep(difference);
              console.log('[resuming from rate limit sleep]');
            }
          }
        }
      } else {
        console.log('Processor ran into an error with an event:');
        console.dir(processInitializationError);
      }
    }
  }
  return interestingEvents;
}

async function verifySignatures(
  signature,
  hookSecrets: string[],
  rawBody
): Promise<void> {
  // To ease local development and simple scenarios, if no shared secrets are
  // configured, they are not required.
  if (!hookSecrets || !hookSecrets.length) {
    return;
  }
  if (!signature) {
    throw new Error('No event signature was provided');
  }
  const computedSignatures = [];
  for (let i = 0; i < hookSecrets.length; i++) {
    const sharedSecret = hookSecrets[i];
    const sha1 = crypto.createHmac('sha1', sharedSecret);
    sha1.update(rawBody, 'utf8');
    const computedHash = 'sha1=' + sha1.digest('hex');
    if (secureCompare(computedHash, signature)) {
      return;
    }
    computedSignatures.push(computedHash);
  }
  const validationError: IValidationError = new Error(
    'The signature could not be verified'
  );
  validationError.statusCode = 401;
  validationError.computedHash = computedSignatures.join(', ');
  throw validationError;
}
