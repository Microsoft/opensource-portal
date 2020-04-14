//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

'use strict';

import throat = require('throat');

// Revoke tokens of users that no longer resolve in the corporate graph and
// delete tokens that have been expired 30 days.

const expiredTokenDeleteThresholdDays = 30;

import { IProviders } from '../../transitional';
import { PersonalAccessToken } from '../../entities/token/token';
import { sleep } from '../../utils';
import { IGraphProvider } from '../../lib/graphProvider';

let insights;

module.exports = function run(config) {
  const app = require('../../app');
  config.skipModules = new Set([
    'web',
  ]);

  app.initializeApplication(config, null, error => {
    if (error) {
      throw error;
    }
    insights = app.settings.appInsightsClient;
    if (!insights) {
      throw new Error('No app insights client available');
    }
    cleanup(config, app).then(done => {
      console.log('done');
      process.exit(0);
    }).catch(error => {
      if (insights) {
        insights.trackException({ exception: error, properties: { name: 'JobCleanupTokensFailure' } });
      }
      throw error;
    });
  });
};

async function lookupCorporateId(graphProvider: IGraphProvider, knownUsers: Map<string, any>, corporateId: string): Promise<any> {
  let entry = knownUsers.get(corporateId);
  if (entry === false) {
    return false;
  } else if (entry) {
    return true;
  }

  try {
    const userDetails = await graphProvider.getUserByIdAsync(corporateId);
    if (!userDetails || !userDetails.userPrincipalName) {
      knownUsers.set(corporateId, false);
      return false;
    }
    knownUsers.set(corporateId, userDetails);
    return true;
  } catch (otherUserError) {
    console.dir(otherUserError);
    throw otherUserError;
  }
}

async function cleanup(config, app) : Promise<void> {
  const providers = app.settings.providers as IProviders;
  const graphProvider = providers.graphProvider;
  const tokenProvider = providers.tokenProvider;

  console.log('reading all tokens');
  const allTokens = await tokenProvider.getAllTokens();
  console.log(`read ${allTokens.length}`);

  insights.trackEvent({ name: 'JobCleanupTokensReadTokens', properties: { tokens: allTokens.length } });

  let errors = 0;

  let revokedUnresolved = 0;
  let deleted = 0;
  let serviceTokens = 0;
  let okUserTokens = 0;

  const parallelUsers = 1;
  const secondsDelayAfterSuccess = 0.25;

  const now = new Date();
  const monthAgo = new Date(now.getTime() - 1000 * 60 * 60 * 24 * expiredTokenDeleteThresholdDays);

  const knownUsers = new Map<string, any>();

  await Promise.all(allTokens.map(throat<void, (token: PersonalAccessToken) => Promise<void>>(async pat => {
    const isGuidMeansADash = pat.corporateId && pat.corporateId.includes('-');
    let wasUser = false;
    if (isGuidMeansADash) {
      wasUser = true;

      const userStatus = await lookupCorporateId(graphProvider, knownUsers, pat.corporateId);
      if (!userStatus && pat.active !== false) {
        pat.active = false;
        console.log(`Revoking key for ${pat.getIdentifier()} - employee ${pat.corporateId} could not be found`);
        try {
          await tokenProvider.updateToken(pat);
          ++revokedUnresolved;
        } catch (tokenUpdateError) {
          console.dir(tokenUpdateError);
          ++errors;
          insights.trackException({ exception: tokenUpdateError });
        }
      }
    } else {
      ++serviceTokens;
    }

    if (pat.isExpired()) {
      const dateExpired = pat.expires;
      if (dateExpired < monthAgo) {
        console.log(`Deleting key for ${pat.getIdentifier()} that expired ${dateExpired}`);
        try {
          await tokenProvider.deleteToken(pat);
          ++deleted;
        } catch (tokenDeleteError) {
          console.dir(tokenDeleteError);
          ++errors;
          insights.trackException({ exception: tokenDeleteError });
        }
      } else {
        console.log(`Expired key, keeping around ${pat.getIdentifier()} that expired ${dateExpired} for user notification purposes`);
      }
    } else if (wasUser) {
      ++okUserTokens;
    }

    await sleep(secondsDelayAfterSuccess * 1000);

  }, parallelUsers)));

  console.log(`deleted: ${deleted}`);
  console.log(`revokedUnresolved: ${revokedUnresolved}`);
  console.log(`okUserTokens: ${okUserTokens}`);
  console.log(`serviceTokens: ${serviceTokens}`);
  console.log();

  insights.trackEvent({ name: 'JobCleanupTokensSuccess', properties: {
    deleted,
    revokedUnresolved,
    okUserTokens,
    serviceTokens,
    errors,
  } });
}
