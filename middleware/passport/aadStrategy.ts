//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

const debug = require('debug')('startup');

import { AuthorizationCode } from 'simple-oauth2';

import { IProviders } from '../../transitional';

import { OIDCStrategy } from 'passport-azure-ad';
import { GraphUserType } from '../../lib/graphProvider';

interface IPassportUserWithAAD {
  azure?: IAADUser;
}

interface IAADUser {
  displayName: string;
  oid: string;
  username: string;
  oauthToken?: string;
}

async function login(
  app,
  config,
  client: AuthorizationCode,
  iss,
  sub,
  profile,
  accessToken: string,
  refreshToken: string,
  params
): Promise<IPassportUserWithAAD> {
  const { graphProvider, insights } = app.settings.providers as IProviders;
  const oauthToken = JSON.stringify(params);
  if (config && config.impersonation && config.impersonation.corporateId) {
    // While impersonation for the site interface is possible, the graph API token,
    // rarely used in this app, will use the actual AAD access tokens still.
    const impersonationCorporateId = config.impersonation.corporateId;
    const impersonationResult = await graphProvider.getUserById(
      impersonationCorporateId
    );
    console.warn(
      `IMPERSONATION: id=${impersonationResult.id} upn=${impersonationResult.userPrincipalName} name=${impersonationResult.displayName} graphIsNotImpersonatedAs=${profile.upn}`
    );
    return {
      azure: {
        displayName: impersonationResult.displayName,
        oid: impersonationResult.id,
        username: impersonationResult.userPrincipalName,
        oauthToken,
      },
    };
  }
  if (config.activeDirectory.blockGuestSignIns === true) {
    const lookupResult = await graphProvider.getUserById(profile.oid);
    if (lookupResult && lookupResult.userType === GraphUserType.Guest) {
      const err = new Error(
        `This application does not permit guests. You are currently signed in to Active Directory as: ${lookupResult.userPrincipalName}`
      );
      insights?.trackException({ exception: err });
      throw err;
    }
  }
  return {
    azure: {
      displayName: profile.displayName,
      oid: profile.oid,
      username: profile.upn,
      oauthToken,
    },
  };
}

function activeDirectorySubset(
  app,
  config,
  client: AuthorizationCode,
  iss,
  sub,
  profile,
  accessToken: string,
  refreshToken: string,
  params,
  done
) {
  login(
    app,
    config,
    client,
    iss,
    sub,
    profile,
    accessToken,
    refreshToken,
    params
  )
    .then((profile) => {
      return done(null, profile);
    })
    .catch((error) => {
      return done(error);
    });
}

export default function createAADStrategy(app, config) {
  const {
    redirectUrl,
    tenantId,
    clientId,
    clientSecret,
  } = config.activeDirectory;
  const providers = app.settings.providers as IProviders;
  const aadAuthority = `https://login.microsoftonline.com/${tenantId}/`;
  const aadMetadata = 'v2.0/.well-known/openid-configuration'; // used to use: .well-known/openid-configuration
  const identityMetadata = `${aadAuthority}${aadMetadata}`;
  const originalIdentityMetadata = `${aadAuthority}.well-known/openid-configuration`;
  const authorizePath = 'oauth2/v2.0/authorize';
  const tokenPath = 'oauth2/v2.0/token';
  const aadScopes = 'profile openid user.read';
  const oauth2Client = new AuthorizationCode({
    client: {
      id: clientId,
      secret: clientSecret,
    },
    auth: {
      tokenHost: aadAuthority,
      tokenPath,
      authorizePath,
    },
  });
  debug(`AAD app clientId=${clientId}, redirectUrl=${redirectUrl}`);
  providers.authorizationCodeClient = oauth2Client;
  const aadStrategy = new OIDCStrategy(
    {
      redirectUrl:
        redirectUrl || `${config.webServer.baseUrl}/auth/azure/callback`,
      allowHttpForRedirectUrl:
        config.containers.docker || config.webServer.allowHttp,
      // @ts-ignore
      realm: tenantId,
      clientID: clientId,
      clientSecret,
      identityMetadata: originalIdentityMetadata,
      responseType: 'id_token code',
      responseMode: 'form_post',
      scope: aadScopes.split(' '),
      // cookieSameSite: true, // ???
      // oidcIssuer: config.activeDirectory.issuer,
      // validateIssuer: true,
    },
    activeDirectorySubset.bind(null, app, config, oauth2Client)
  );
  // Patching the AAD strategy to intercept a specific state failure message and instead
  // of providing a generic failure message, redirecting (HTTP GET) to the callback page
  // where we can offer a more useful message
  // @ts-ignore
  const originalFailWithLog = aadStrategy.failWithLog;
  // @ts-ignore
  aadStrategy.failWithLog = function () {
    const args = Array.prototype.slice.call(arguments);
    const messageToIntercept =
      'In collectInfoFromReq: invalid state received in the request';
    if (args.length === 1 && typeof args[0] === 'string') {
      console.warn(
        `AAD Failure: clientId=${clientId}, tenantId=${tenantId}, message=${args[0]}`
      );
    }
    if (
      args.length === 1 &&
      typeof args[0] === 'string' &&
      args[0] === messageToIntercept
    ) {
      return this.redirect('/auth/azure/callback?failure=invalid');
    }
    return originalFailWithLog.call(this, args);
  };
  return { 'azure-active-directory': aadStrategy };
}
