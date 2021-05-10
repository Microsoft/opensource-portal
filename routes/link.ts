//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn"] }] */

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { ReposAppRequest, IAppSession, SupportedLinkType, ICorporateLink, LinkOperationSource } from '../interfaces';
import { getProviders } from '../transitional';
import { IndividualContext } from '../user';
import { storeOriginalUrlAsReferrer, wrapError } from '../utils';

import validator from 'validator';

import unlinkRoute from './unlink';
import { jsonError } from '../middleware';

interface IRequestWithSession extends ReposAppRequest {
  session: IAppSession;
}

interface IRequestHacked extends ReposAppRequest {
  overrideLinkUserPrincipalName?: any;
}

router.use((req: IRequestHacked, res, next) => {
  const config = getProviders(req).config;;
  if (config && config.github && config.github.links && config.github.links.provider && config.github.links.provider.linkingOfflineMessage) {
    return next(new Error(`Linking is temporarily offline: ${config.github.links.provider.linkingOfflineMessage}`));
  } else {
    return next();
  }
});

router.use('/', asyncHandler(async function (req: ReposAppRequest, res, next) {
  // Make sure both account types are authenticated before showing the link pg [wi 12690]
  const individualContext = req.individualContext;
  if (!individualContext.corporateIdentity || !individualContext.getGitHubIdentity()) {
    req.insights.trackEvent({ name: 'PortalSessionNeedsBothGitHubAndAadUsernames' });
    return res.redirect('/?signin');
  }
  return next();
}));

// TODO: graph provider non-guest check should be middleware and in the link business process

router.use(asyncHandler(async (req: IRequestHacked, res, next) => {
  const individualContext = req.individualContext as IndividualContext;
  const providers = getProviders(req);
  const insights = providers.insights;
  const config = providers.config;
  let validateAndBlockGuests = false;
  if (config && config.activeDirectory && config.activeDirectory.blockGuestUserTypes) {
    validateAndBlockGuests = true;
  }
  // If the app has not been configured to check whether a user is a guest before linking, continue:
  if (!validateAndBlockGuests) {
    return next();
  }
  const aadId = individualContext.corporateIdentity.id;
  // If the app is configured to check guest status, do this now, before linking:
  const graphProvider = providers.graphProvider;
  // REFACTOR: delegate the decision to the auth provider
  if (!graphProvider || !graphProvider.getUserById) {
    return next(new Error('User type validation cannot be performed because there is no graphProvider configured for this type of account'));
  }
  insights.trackEvent({
    name: 'LinkValidateNotGuestStart',
    properties: {
      aadId: aadId,
    },
  });
  try {
    const details = await graphProvider.getUserById(aadId);
    const userType = details.userType;
    const displayName = details.displayName;
    const userPrincipalName = details.userPrincipalName;
    let block = userType as string === 'Guest';
    let blockedRecord = block ? 'BLOCKED' : 'not blocked';
    // If the app is configured to check for guests, but this is a specifically permitted guest user, continue:
    if (config && config.activeDirectoryGuests && config.activeDirectoryGuests.authorizedIds && config.activeDirectoryGuests.authorizedIds.length && config.activeDirectoryGuests.authorizedIds.includes(aadId)) {
      block = false;
      blockedRecord = 'specifically authorized user ' + aadId + ' ' + userPrincipalName;
      req.overrideLinkUserPrincipalName = userPrincipalName;
      return next(new Error('This feature is not currently available. Please reach out to support to re-enable this feature.'));
    }
    insights.trackEvent({
      name: 'LinkValidateNotGuestGraphSuccess',
      properties: {
        aadId: aadId,
        userType: userType,
        displayName: displayName,
        userPrincipalName: userPrincipalName,
        blocked: blockedRecord,
      },
    });
    if (block) {
      insights.trackMetric({ name: 'LinksBlockedForGuests', value: 1 });
      return next(new Error(`This system is not available to guests. You are currently signed in as ${displayName} ${userPrincipalName}. Please sign out or try a private browser window.`));
    }
    const manager = await providers.graphProvider.getManagerById(aadId);
    if (!manager || !manager.userPrincipalName) {
      throw new Error(`You do not have an active manager entry in the directory and so cannot yet link.`);
    }
    return next();
  } catch (graphError) {
    insights.trackException({
      exception: graphError,
      properties: {
        aadId: aadId,
        name: 'LinkValidateNotGuestGraphFailure',
      },
    });
    return next(graphError);
  }
}));

router.get('/', asyncHandler(async function (req: ReposAppRequest, res, next) {
  const individualContext = req.individualContext;
  const link = individualContext.link;
  if (!individualContext.corporateIdentity && !individualContext.getGitHubIdentity()) {
    req.insights.trackEvent({ name: 'PortalSessionNeedsBothGitHubAndAadUsernames' });
    return res.redirect('/?signin');
  }
  if (!individualContext.getGitHubIdentity()) {
    req.insights.trackEvent({ name: 'PortalSessionNeedsGitHubUsername' });
    return res.redirect('/signin/github/');
  }
  if (!link) {
    return await showLinkPage(req, res);
  } else {
    req.insights.trackEvent({ name: 'LinkRouteLinkLocated' });
    let organizations = null;
    try {
      organizations = await individualContext.aggregations.organizations();
    } catch (ignoredError) {
      /* ignore */
    }
    return individualContext.webContext.render({
      view: 'linkConfirmed',
      title: 'You\'re already linked',
      state: {
        organizations,
      }
    });
  }
}));

async function showLinkPage(req: ReposAppRequest, res) {
  const individualContext = req.individualContext as IndividualContext;
  function render(options) {
    individualContext.webContext.render({
      view: 'link',
      title: 'Link GitHub with corporate identity',
      optionalObject: options || {},
    })
  }
  const { config, graphProvider } = getProviders(req);
  if (config.authentication.scheme !== 'aad' || !graphProvider) {
    return render(null);
  }
  const aadId = individualContext.corporateIdentity.id;
  const { operations } = getProviders(req);
  // By design, we want to log the errors but do not want any individual
  // lookup problem to break the underlying experience of letting a user
  // link. This is important if someone is new in the company, they may
  // not be in the graph fully yet.
  const userLinkData = await operations.validateCorporateAccountCanLink(aadId);
  render({
    graphUser: userLinkData.graphEntry,
    isServiceAccountCandidate: userLinkData.type === SupportedLinkType.ServiceAccount,
  });
}

router.get('/enableMultipleAccounts', function (req: IRequestWithSession, res) {
  // LEGACY
  // TODO: is this code still ever really used?
  if (req.user.github) {
    req.session.enableMultipleAccounts = true;
    return res.redirect('/link/cleanup');
  }
  req.insights.trackEvent({ name: 'PortalUserEnabledMultipleAccounts' });
  storeOriginalUrlAsReferrer(req, res, '/auth/github', 'multiple accounts enabled need to auth with GitHub again now');
});

router.post('/', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const individualContext = req.individualContext as IndividualContext;
  try {
    await interactiveLinkUser(false, individualContext, req, res, next);
  } catch (error) {
    return next(error);
  }
}));

export async function interactiveLinkUser(isJson: boolean, individualContext: IndividualContext, req, res, next) {
  const isServiceAccount = req.body.sa === '1';
  const serviceAccountMail = req.body.serviceAccountMail;
  const { operations } = getProviders(req);
  if (isServiceAccount && !validator.isEmail(serviceAccountMail)) {
    const errorMessage = 'Please enter a valid e-mail address for the Service Account maintainer.'
    return next(isJson ? jsonError(errorMessage, 400) : wrapError(null, errorMessage, true));
  }
  let newLinkObject: ICorporateLink = null;
  try {
    newLinkObject = individualContext.createGitHubLinkObject();
  } catch (missingInformationError) {
    return next(missingInformationError);
  }
  if (isServiceAccount) {
    newLinkObject.isServiceAccount = true;
    newLinkObject.serviceAccountMail = serviceAccountMail;
    const address = operations.getOperationsMailAddress();
    const errorMessage = `Service Account linking is not available. Please reach out to ${address} for more information.`;
    return next(isJson ? jsonError(errorMessage, 400) : new Error(errorMessage));
  }
  try {
    await operations.linkAccounts({
      link: newLinkObject,
      operationSource: LinkOperationSource.Portal,
      correlationId: individualContext.webContext?.correlationId || 'N/A',
      skipGitHubValidation: true, // already has been verified in the recent session
    });
    if (isJson) {
      res.status(201);
      return res.end();
    } else {
      return res.redirect('/?onboarding=yes');
    }
  } catch (createError) {
    const errorMessage = `We had trouble linking your corporate and GitHub accounts: ${createError.message}`;
    return next(isJson ? jsonError(errorMessage, 500) : wrapError(createError, errorMessage));
  }
}

router.use('/remove', unlinkRoute);

router.get('/reconnect', function (req: ReposAppRequest, res, next) {
  const config = getProviders(req).config;;
  if (config.authentication.scheme !== 'aad') {
    return next(wrapError(null, 'Account reconnection is only needed for Active Directory authentication applications.', true));
  }
  // If the request comes back to the reconnect page, the authenticated app will
  // actually update the link the next time around.
  const ghi = req.individualContext.getGitHubIdentity();
  const hasToken = !!req.individualContext.webContext.tokens.gitHubReadToken;
  if (ghi && ghi.id && ghi.username && hasToken) {
    req.insights.trackEvent({ name: 'PortalUserReconnected' });
    return res.redirect('/');
  }
  req.insights.trackEvent({ name: 'PortalUserReconnectNeeded' });
  req.individualContext.webContext.render({
    view: 'reconnectGitHub',
    title: 'Please sign in with GitHub',
    state: {
      expectedUsername: ghi.username,
    },
  });
});

export default router;
