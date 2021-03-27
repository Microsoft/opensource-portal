//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Strategy as GithubStrategy } from 'passport-github';

import { IProviders } from '../../transitional';
import { Operations } from '../../business';

function githubResponseToSubset(
  app,
  modernAppInUse: boolean,
  accessToken: string,
  refreshToken: string,
  profile,
  done
) {
  const config = app.settings.runtimeConfig;
  const {
    useIncreasedScopeLegacyAppIfNeeded,
  } = getGithubAppConfigurationOptions(config);
  const providers = app.settings.providers as IProviders;
  if (config && config.impersonation && config.impersonation.githubId) {
    const operations = providers.operations as Operations;
    const impersonationId = config.impersonation.githubId;
    const account = operations.getAccount(impersonationId);
    return account
      .getDetails()
      .then((details) => {
        console.warn(
          `GITHUB IMPERSONATION: id=${impersonationId} login=${details.login} name=${details.name}`
        );
        return done(null, {
          github: {
            accessToken: 'fakeaccesstoken',
            displayName: details.name,
            avatarUrl: details.avatar_url,
            id: details.id.toString(),
            username: details.login,
          },
        });
      })
      .catch((err) => {
        return done(err);
      });
  }
  let subset = {
    github: {
      accessToken: accessToken,
      displayName: profile.displayName,
      avatarUrl:
        profile._json && profile._json.avatar_url
          ? profile._json.avatar_url
          : undefined,
      id: profile.id,
      username: profile.username,
      scope: undefined,
    },
  };
  if (modernAppInUse) {
    subset.github.scope = 'githubapp';
  } else if (useIncreasedScopeLegacyAppIfNeeded && !modernAppInUse) {
    subset.github.scope = 'githubapp'; // for local development validation of this interesting scenario only
  }
  return done(null, subset);
}

function githubResponseToIncreasedScopeSubset(
  modernAppInUse: boolean,
  accessToken: string,
  refreshToken: string,
  profile,
  done
) {
  if (modernAppInUse) {
    return done(
      new Error(
        'githubResponseToIncreasedScopeSubset is not compatible with modern apps'
      )
    );
  }
  const subset = {
    githubIncreasedScope: {
      accessToken: accessToken,
      id: profile.id,
      username: profile.username,
    },
  };
  return done(null, subset);
}

export function getGithubAppConfigurationOptions(config) {
  let legacyOAuthApp =
    config.github.oauth2 &&
    config.github.oauth2.clientId &&
    config.github.oauth2.clientSecret
      ? config.github.oauth2
      : null;
  const customerFacingApp =
    config.github.app &&
    config.github.app.ui &&
    config.github.app.ui.clientId &&
    config.github.app.ui.clientSecret
      ? config.github.app.ui
      : null;
  const useCustomerFacingGithubAppIfPresent =
    config.github.oauth2.useCustomerFacingGitHubAppIfPresent === true;
  const useIncreasedScopeLegacyAppIfNeeded =
    config.github.oauth2.useIncreasedScopeCustomerFacingIfNeeded === true;
  if (useCustomerFacingGithubAppIfPresent && customerFacingApp) {
    if (legacyOAuthApp && legacyOAuthApp['callbackUrl']) {
      customerFacingApp['callbackUrl'] = legacyOAuthApp['callbackUrl'];
    }
    legacyOAuthApp = null;
  }
  const modernAppInUse: boolean = customerFacingApp && !legacyOAuthApp;
  const githubAppConfiguration = modernAppInUse
    ? customerFacingApp
    : legacyOAuthApp;
  return {
    legacyOAuthApp,
    customerFacingApp,
    modernAppInUse,
    githubAppConfiguration,
    useIncreasedScopeLegacyAppIfNeeded,
  };
}

export default function createGithubStrategy(app, config) {
  let strategies = {};
  const {
    modernAppInUse,
    githubAppConfiguration,
    useIncreasedScopeLegacyAppIfNeeded,
  } = getGithubAppConfigurationOptions(config);
  if (modernAppInUse) {
    console.log(
      `GitHub App for customer-facing OAuth in use, client ID=${githubAppConfiguration.clientId}`
    );
  } else {
    console.log(
      `Legacy GitHub OAuth app being used for customers, client ID=${githubAppConfiguration.clientId}`
    );
  }
  const writeOrgScopes = ['write:org'];
  const scope =
    useIncreasedScopeLegacyAppIfNeeded && !modernAppInUse ? writeOrgScopes : [];
  if (useIncreasedScopeLegacyAppIfNeeded && !modernAppInUse) {
    console.log(
      `Legacy GitHub OAuth app will use the expanded token with org-write scope`
    );
  }
  // GitHub Passport session setup.
  let githubOptions = {
    clientID: githubAppConfiguration.clientId,
    clientSecret: githubAppConfiguration.clientSecret,
    callbackURL: undefined,
    scope,
    userAgent: 'passport-azure-oss-portal-for-github', // CONSIDER: User agent should be configured.
  };
  if (githubAppConfiguration.callbackUrl) {
    githubOptions.callbackURL = githubAppConfiguration.callbackUrl;
  }
  let githubPassportStrategy = new GithubStrategy(
    githubOptions,
    githubResponseToSubset.bind(null, app, modernAppInUse)
  );
  // Validate the borrow some parameters from the GitHub passport library
  strategies['github'] = githubPassportStrategy;
  // Expanded OAuth-scope GitHub access for org membership writes.
  if (!modernAppInUse) {
    // new GitHub Apps no longer have a separate scope concept
    let expandedGithubScopeStrategy = new GithubStrategy(
      {
        clientID: githubOptions.clientID,
        clientSecret: githubOptions.clientSecret,
        callbackURL: `${githubOptions.callbackURL}/increased-scope`,
        scope: writeOrgScopes,
        userAgent: 'passport-azure-oss-portal-for-github', // CONSIDER: User agent should be configured.
      },
      githubResponseToIncreasedScopeSubset.bind(null, modernAppInUse)
    );
    strategies['expanded-github-scope'] = expandedGithubScopeStrategy;
  }
  return strategies;
}
