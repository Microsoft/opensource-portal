//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "log", "dir"] }] */

import _ from 'lodash';
const debug = require('debug')('user');

import { getProviders, IAppSession, ReposAppRequest } from '../../transitional';
import {
  ICorporateIdentity,
  IGitHubIdentity,
  IndividualContext,
  GitHubIdentitySource,
} from '../../user';
import { storeOriginalUrlAsReferrer } from '../../utils';

export function requireAuthenticatedUserOrSignInExcluding(
  exclusionPaths: string[],
  req: ReposAppRequest,
  res,
  next
) {
  const baseUrl = req.baseUrl;
  for (let i = 0; i < exclusionPaths.length; i++) {
    if (baseUrl.startsWith(exclusionPaths[i])) {
      console.log(
        `${req.method} ${req.baseUrl} excluded from auth by prefix: ${exclusionPaths[i]}`
      );
      return next();
    }
  }
  return requireAuthenticatedUserOrSignIn(req, res, next);
}

export async function requireAccessTokenClient(
  req: ReposAppRequest,
  res,
  next
) {
  if (req.oauthAccessToken) {
    return next();
  }
  // This code currently assumes you're using AAD.
  const { authorizationCodeClient } = getProviders(req);
  if (!req.user.azure) {
    console.warn('Not an Azure authenticated user');
    return signoutThenSignIn(req, res);
  }
  // Build an OAuth Access Token instance for the request, refreshing as needed
  const { oauthToken } = req.user.azure;
  if (authorizationCodeClient && oauthToken) {
    const hydratedToken = JSON.parse(oauthToken);
    let oauthTokenInstance = authorizationCodeClient.createToken(hydratedToken);
    if (oauthTokenInstance.expired()) {
      oauthTokenInstance = await oauthTokenInstance.refresh();
      const session = req.session as IAppSession;
      session.passport.user.azure.oauthToken = JSON.stringify(
        oauthTokenInstance.token
      );
    }
    req.oauthAccessToken = oauthTokenInstance;
  } else {
    // this is only used during the transition to storing this kind of data
    // console.warn('The user did not have an oauthToken available');
    // return signoutThenSignIn(req, res);
  }
  return next();
}

function signoutThenSignIn(req, res) {
  req.logout();
  return redirectToSignIn(req, res);
}

function redirectToSignIn(req, res) {
  const config = getProviders(req).config;
  storeOriginalUrlAsReferrer(
    req,
    res,
    config.authentication.scheme === 'github' ? '/auth/github' : '/auth/azure',
    'user is not authenticated and needs to authenticate'
  );
}

export function requireAuthenticatedUserOrSignIn(
  req: ReposAppRequest,
  res,
  next
) {
  const config = getProviders(req).config;
  if (req.isAuthenticated()) {
    const expectedAuthenticationProperty =
      config.authentication.scheme === 'github' ? 'github' : 'azure';
    if (req.user && !req.user[expectedAuthenticationProperty]) {
      console.warn(
        `A user session was authenticated but did not have present the property "${expectedAuthenticationProperty}" expected for this type of authentication. Signing them out.`
      );
      return res.redirect('/signout');
    }
    const expectedAuthenticationKey =
      config.authentication.scheme === 'github' ? 'id' : 'oid';
    if (!req.user[expectedAuthenticationProperty][expectedAuthenticationKey]) {
      return next(
        new Error(
          'Invalid information present for the authentication provider.'
        )
      );
    }
    return next();
  }
  return redirectToSignIn(req, res);
}

export function setIdentity(req: ReposAppRequest, res, next) {
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  if (!activeContext) {
    return next(new Error('No context available'));
  }
  const contextName = req.individualContext
    ? 'Individual User Context'
    : 'API Context';

  let requestForAuthentication = req;
  let sourceText = 'AUTHENTICATED SESSION';

  const overwrittenRequestSource = req['userContextOverwriteRequest'];
  if (overwrittenRequestSource) {
    console.warn(
      'userContextOverwriteRequest: *SUBSTITUTING* session identity with another source of data'
    );
    console.dir(overwrittenRequestSource);
    requestForAuthentication = overwrittenRequestSource;
    sourceText = 'OVERWRITTEN SOURCE OF TRUTH';
  }

  let corporateIdentity: ICorporateIdentity = null;
  let gitHubIdentity: IGitHubIdentity = null;

  let s = `${contextName} ${sourceText}: `;
  const user = requestForAuthentication.user;

  if (user.github) {
    s += `github(id=${user.github.id}, username=${user.github.username}) `;
    gitHubIdentity = {
      id: user.github.id,
      username: user.github.username,
      displayName: user.github.displayName,
      avatar: user.github.avatarUrl,
      source: GitHubIdentitySource.Session,
    };
  }
  if (user.azure) {
    s += `azure(oid=${user.azure.oid}, username=${user.azure.username}) `;
    corporateIdentity = {
      id: user.azure.oid,
      username: user.azure.username,
      displayName: user.azure.displayName,
    };
  }
  debug(s);

  // const { insights } = getProviders(req);

  activeContext.corporateIdentity = corporateIdentity;
  activeContext.setSessionBasedGitHubIdentity(gitHubIdentity);

  if (activeContext.webContext) {
    activeContext.webContext.pushBreadcrumb('Organizations');
  }

  if (!corporateIdentity) {
    // JUST TESTING:
    // TODO: when and where does this happen?
    return next(new Error('Not a valid corporate user'));
  }

  return next();
}
