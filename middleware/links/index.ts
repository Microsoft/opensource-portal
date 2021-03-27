//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Within the context, tries to resolve a link if it _can_. It does not force that a user is linked!

import { IndividualContext } from '../../user';
import { ReposAppRequest, IReposError, getProviders } from '../../transitional';
import { wrapError } from '../../utils';

export function RequireLinkMatchesGitHubSessionExceptPrefixedRoute(
  prefix: string
) {
  return requireLinkMatchesGitHubSession.bind(null, prefix);
}

export function RequireLinkMatchesGitHubSession(
  req: ReposAppRequest,
  res,
  next
) {
  return requireLinkMatchesGitHubSession(null, req, res, next);
}

function requireLinkMatchesGitHubSession(
  allowedPrefix: string,
  req: ReposAppRequest,
  res,
  next
) {
  // trying to be equivalent to legacy code in ./usernameConsistency (lightweight)
  const context = req.individualContext;
  if (!context) {
    return next(new Error('Missing context'));
  }
  if (!context.link || !context.link.thirdPartyId) {
    return next();
  }
  const gitHubIdentity = context.getGitHubIdentity();
  const sessionIdentity =
    context.getSessionBasedGitHubIdentity() || gitHubIdentity;
  if (gitHubIdentity && gitHubIdentity.id === sessionIdentity.id) {
    return next();
  }
  if (allowedPrefix && req.path.startsWith(allowedPrefix)) {
    console.log(
      `Mixed GitHub identity issue. Allowed prefix ${allowedPrefix} matches for ${req.path}, allowing downstream route`
    );
    return next();
  }
  let securityError: IReposError = new Error(
    `Your GitHub account identity has changed.`
  );
  securityError.detailed = `When you linked your GitHub account to your corporate identity, you used the GitHub account with the username ${gitHubIdentity.username} (GitHub user ID ${gitHubIdentity.id}), but you are currently signed into GitHub with the username of ${sessionIdentity.username} (GitHub user ID ${sessionIdentity.id}). Please sign out of this site and GitHub and try again.`;
  securityError.fancyLink = {
    title: 'Unlink my account',
    link: '/unlink',
  };
  securityError.fancySecondaryLink = {
    title: 'Sign out of GitHub',
    link: '/signout/github?redirect=github',
  };
  securityError.skipOops = true;

  // TODO_LOW: support multi-account again, if necessary
  const multipleAccountsEnabled =
    sessionIdentity.id &&
    context.webContext['_fake*property_session_enableMultipleAccounts'] ===
      true;
  if (multipleAccountsEnabled) {
    securityError = wrapError(
      null,
      'You are currently signed in to an account on GitHub.com that is different than the one you have selected for your session. Please sign out of GitHub and head back.',
      true
    );
    securityError.fancyLink = {
      title: 'Sign out of GitHub',
      link: '/signout/github?redirect=github',
    };
  }
  return next(securityError);
}

export async function AddLinkToRequest(req, res, next) {
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  const contextName = req.individualContext
    ? 'Individual User Context'
    : 'API Context';
  if (!activeContext) {
    return next(new Error('The middleware requires a context'));
  }
  if (!activeContext.corporateIdentity) {
    return next();
  }
  if (activeContext.link) {
    return next();
  }
  const corporateId = activeContext.corporateIdentity.id;
  if (!corporateId) {
    return next(new Error('No corporate user information'));
  }
  const { linkProvider } = getProviders(req);
  if (!linkProvider) {
    return next(new Error('No link provider'));
  }
  const links = await linkProvider.queryByCorporateId(corporateId);
  if (links.length === 0) {
    return next();
  }
  if (links.length > 1) {
    // TODO: are multiple links selected through a session or web context setting, or ?
    return next(new Error('You cannot have multiple GitHub accounts'));
  }
  const selectedLink = links[0];
  activeContext.link = selectedLink;
  return next();
}
