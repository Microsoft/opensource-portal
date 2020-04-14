//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

import { IProviders } from "../transitional";
import { Operations } from "../business/operations";

const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;

const serializer = require('./passport/serializer');

const passportGitHub = require('passport-github');

const GitHubStrategy = passportGitHub.Strategy;

function githubResponseToSubset(app, modernAppInUse: boolean, accessToken: string, refreshToken: string, profile, done) {
  const config = app.settings.runtimeConfig;
  const providers = app.settings.providers as IProviders;
  if (config && config.impersonation && config.impersonation.githubId) {
    const operations = providers.operations as Operations;
    const impersonationId = config.impersonation.githubId;

    const account = operations.getAccount(impersonationId);
    return account.getDetails().then(details => {
      console.warn(`GITHUB IMPERSONATION: id=${impersonationId} login=${details.login} name=${details.name}`);
      return done(null, {
        github: {
          accessToken: 'fakeaccesstoken',
          displayName: details.name,
          avatarUrl: details.avatar_url,
          id: details.id.toString(),
          username: details.login,
        },
      });
    }).catch(err => {
      return done(err);
    });
  }
  let subset = {
    github: {
      accessToken: accessToken,
      displayName: profile.displayName,
      avatarUrl: profile._json && profile._json.avatar_url ? profile._json.avatar_url : undefined,
      id: profile.id,
      username: profile.username,
      scope: undefined,
    },
  };
  if (modernAppInUse) {
    subset.github.scope = 'githubapp';
  }
  return done(null, subset);
}

function githubResponseToIncreasedScopeSubset(modernAppInUse: boolean, accessToken: string, refreshToken: string, profile, done) {
  if (modernAppInUse) {
    return done(new Error('githubResponseToIncreasedScopeSubset is not compatible with modern apps'));
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

function activeDirectorySubset(app, iss, sub, profile, done) {
  // CONSIDER: TODO: Hybrid tenant checks.
  // Internal-only code:
  // ----------------------------------------------------------------
  // We've identified users with e-mail addresses in AAD similar to
  // myoutlookaddress#live.com. These are where people have had work
  // shared with them through a service like Office 365; these users
  // are not technically employees with active credentials, and so
  // they should *not* have access. We reject here before the
  // session tokens can be saved.
  // if (username && username.indexOf && username.indexOf('#') >= 0) {
  //   return next(new Error('Your hybrid tenant account, ' + username + ', is not permitted for this resource. Were you invited as an outside collaborator by accident? Please contact us if you have any questions.'));
  // }

  const config = app.settings.runtimeConfig;
  const providers = app.settings.providers as IProviders;
  if (config && config.impersonation && config.impersonation.corporateId) {
    const impersonationCorporateId = config.impersonation.corporateId;
    return providers.graphProvider.getUserById(impersonationCorporateId, (err, impersonationResult) => {
      if (err) {
        return done(err);
      }
      console.warn(`IMPERSONATION: id=${impersonationResult.id} upn=${impersonationResult.userPrincipalName} name=${impersonationResult.displayName}`);
      return done(null, {
        azure: {
          displayName: impersonationResult.displayName,
          oid: impersonationResult.id,
          username: impersonationResult.userPrincipalName,
        },
      });
    });
  }
  const subset = {
    azure: {
      displayName: profile.displayName,
      oid: profile.oid,
      username: profile.upn,
    },
  };
  return done(null, subset);
}

export function getGitHubAppConfigurationOptions(config) {
  let legacyOAuthApp = config.github.oauth2 && config.github.oauth2.clientId && config.github.oauth2.clientSecret ? config.github.oauth2 : null;
  const customerFacingApp = config.github.app && config.github.app.ui && config.github.app.ui.clientId && config.github.app.ui.clientSecret ? config.github.app.ui : null;
  const useCustomerFacingGitHubAppIfPresent = config.github.oauth2.useCustomerFacingGitHubAppIfPresent === true;
  if (useCustomerFacingGitHubAppIfPresent && customerFacingApp) {
    if (legacyOAuthApp && legacyOAuthApp['callbackUrl']) {
      customerFacingApp['callbackUrl'] = legacyOAuthApp['callbackUrl'];
    }
    legacyOAuthApp = null;
  }
  const modernAppInUse = customerFacingApp && !legacyOAuthApp;
  const githubAppConfiguration = modernAppInUse ? customerFacingApp : legacyOAuthApp;
  return { legacyOAuthApp, customerFacingApp, modernAppInUse, githubAppConfiguration };
}

export default function (app, config) {
  if (config.authentication.scheme !== 'github' && config.authentication.scheme !== 'aad') {
    throw new Error(`Unsupported primary authentication scheme type "${config.authentication.scheme}"`);
  }
  const { modernAppInUse, githubAppConfiguration } = getGitHubAppConfigurationOptions(config);
  // NOTE: due to bugs in the GitHub API v3 around user-to-server requests in
  // the new GitHub model, it is better to use an original GitHub OAuth app
  // for user interaction right now until those bugs are corrected. What this
  // does mean is that any GitHub org that should be managed by this portal
  // needs the OAuth app to be authorized as a third-party app for the org or
  // to have the auto-accept invite experience work. (9/24/2019)
  if (modernAppInUse) {
    console.log(`GitHub App for customer-facing OAuth in use, client ID=${githubAppConfiguration.clientId}`);
  } else {
    console.log(`Legacy GitHub OAuth app being used for customers, client ID=${githubAppConfiguration.clientId}`);
  }
  // ----------------------------------------------------------------------------
  // GitHub Passport session setup.
  // ----------------------------------------------------------------------------
  let githubOptions = {
    clientID: githubAppConfiguration.clientId,
    clientSecret: githubAppConfiguration.clientSecret,
    callbackURL: undefined,
    scope: [],
    userAgent: 'passport-azure-oss-portal-for-github' // CONSIDER: User agent should be configured.
  };
  if (githubAppConfiguration.callbackUrl) {
    githubOptions.callbackURL = githubAppConfiguration.callbackUrl
  }
  let githubPassportStrategy = new GitHubStrategy(githubOptions, githubResponseToSubset.bind(null, app, modernAppInUse));
  let aadStrategy = new OIDCStrategy({
    redirectUrl: config.activeDirectory.redirectUrl || `${config.webServer.baseUrl}/auth/azure/callback`,
    allowHttpForRedirectUrl: config.containers.docker || config.webServer.allowHttp,
    realm: config.activeDirectory.tenantId,
    clientID: config.activeDirectory.clientId,
    clientSecret: config.activeDirectory.clientSecret,
    identityMetadata: 'https://login.microsoftonline.com/' + config.activeDirectory.tenantId + '/.well-known/openid-configuration',
    responseType: 'id_token code',
    responseMode: 'form_post',
    // oidcIssuer: config.activeDirectory.issuer,
    // validateIssuer: true,
  }, activeDirectorySubset.bind(null, app));

  // Patching the AAD strategy to intercept a specific state failure message and instead
  // of providing a generic failure message, redirecting (HTTP GET) to the callback page
  // where we can offer a more useful message
  const originalFailWithLog = aadStrategy.failWithLog;
  aadStrategy.failWithLog = function () {
    const args = Array.prototype.slice.call(arguments);
    const messageToIntercept = 'In collectInfoFromReq: invalid state received in the request';
    if (args.length === 1 && typeof(args[0]) === 'string' && args[0] === messageToIntercept) {
      return this.redirect('/auth/azure/callback?failure=invalid');
    } else if (args.length === 1 && typeof(args[0]) === 'string') {
      console.warn(`AAD Failure: ${args[0]}`);
    }
    originalFailWithLog.call(this, args);
  };

  // Validate the borrow some parameters from the GitHub passport library
  if (githubPassportStrategy._oauth2 && githubPassportStrategy._oauth2._authorizeUrl) {
    app.set('runtime/passport/github/authorizeUrl', githubPassportStrategy._oauth2._authorizeUrl);
  } else {
    throw new Error('The GitHub Passport strategy library may have been updated, it no longer contains the expected Authorize URL property within the OAuth2 object.');
  }
  if (githubPassportStrategy._scope && githubPassportStrategy._scopeSeparator) {
    app.set('runtime/passport/github/scope', githubPassportStrategy._scope.join(githubPassportStrategy._scopeSeparator));
  } else {
    throw new Error('The GitHub Passport strategy library may have been updated, it no longer contains the expected Authorize URL property within the OAuth2 object.');
  }

  passport.use('github', githubPassportStrategy);
  passport.use('azure-active-directory', aadStrategy);

  // ----------------------------------------------------------------------------
  // Expanded OAuth-scope GitHub access for org membership writes.
  // ----------------------------------------------------------------------------
  if (!modernAppInUse) { // new GitHub Apps no longer have a separate scope concept
    let expandedGitHubScopeStrategy = new GitHubStrategy({
      clientID: githubOptions.clientID,
      clientSecret: githubOptions.clientSecret,
      callbackURL: `${githubOptions.callbackURL}/increased-scope`,
      scope: ['write:org'],
      userAgent: 'passport-azure-oss-portal-for-github' // CONSIDER: User agent should be configured.
    }, githubResponseToIncreasedScopeSubset.bind(null, modernAppInUse));

    passport.use('expanded-github-scope', expandedGitHubScopeStrategy);
  }

  app.use(passport.initialize());
  app.use(passport.session());

  const serializerOptions = {
    config: config,
    keyResolver: app.get('keyEncryptionKeyResolver'),
  };

  passport.serializeUser(serializer.serialize(serializerOptions));
  passport.deserializeUser(serializer.deserialize(serializerOptions));
  serializer.initialize(serializerOptions, app);

  app.use((req, res, next) => {
    if (req.insights && req.insights.properties && config.authentication.scheme === 'aad' && req.user && req.user.azure) {
      req.insights.properties.aadId = req.user.azure.oid;
    }
    next();
  });

  return passport;
};
