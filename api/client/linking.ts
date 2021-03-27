//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';

import { IndividualContext } from '../../user';
import { jsonError } from '../../middleware';
import { getProviders, ReposAppRequest } from '../../transitional';
import { unlinkInteractive } from '../../routes/unlink';
import { interactiveLinkUser } from '../../routes/link';

const router = express.Router();

async function validateLinkOk(req: ReposAppRequest, res, next) {
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  const providers = getProviders(req);
  const insights = providers.insights;
  const config = providers.config;
  let validateAndBlockGuests = false;
  if (
    config &&
    config.activeDirectory &&
    config.activeDirectory.blockGuestUserTypes
  ) {
    validateAndBlockGuests = true;
  }
  // If the app has not been configured to check whether a user is a guest before linking, continue:
  if (!validateAndBlockGuests) {
    return next();
  }
  const aadId = activeContext.corporateIdentity.id;
  // If the app is configured to check guest status, do this now, before linking:
  const graphProvider = providers.graphProvider;
  // REFACTOR: delegate the decision to the auth provider
  if (!graphProvider || !graphProvider.getUserById) {
    return next(jsonError('No configured graph provider', 500));
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
    let block = (userType as string) === 'Guest';
    let blockedRecord = block ? 'BLOCKED' : 'not blocked';
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
      const err = jsonError(
        `This system is not available to guests. You are currently signed in as ${displayName} ${userPrincipalName}. Please sign out or try a private browser window.`,
        400
      );
      insights?.trackException({ exception: err });
      return next(err);
    }
    const manager = await providers.graphProvider.getManagerById(aadId);
    if (!manager || !manager.userPrincipalName) {
      return next(
        jsonError(
          'You do not have an active manager entry in the directory, so cannot yet use this app to link.',
          400
        )
      );
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
    return next(jsonError('Generic graph error', 500));
  }
}

router.delete(
  '/',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    return unlinkInteractive(true, activeContext, req, res, next);
  })
);

router.post(
  '/',
  validateLinkOk,
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    return interactiveLinkUser(true, activeContext, req, res, next);
  })
);

router.use('*', (req: ReposAppRequest, res, next) => {
  return next(jsonError('API or route not found', 404));
});

export default router;
