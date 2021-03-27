//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import {
  ReposAppRequest,
  CreateError,
  UserAlertType,
  getProviders,
} from '../../transitional';

router.use('/:campaignGroupId', (req: ReposAppRequest, res: any, next) => {
  const { config } = getProviders(req);
  const knownCampaignGroups = (config?.campaigns?.groups || '')
    .toLowerCase()
    .split(',');
  req.params.campaignGroupId = req.params.campaignGroupId.toLowerCase();
  const { campaignGroupId } = req.params;
  if (!knownCampaignGroups.includes(campaignGroupId)) {
    return next(
      CreateError.NotFound(
        `The campaign ${campaignGroupId} is not registered in this environment.`
      )
    );
  }
  return next();
});

router.get(
  '/:campaignGroupId/unsubscribe',
  asyncHandler(async (req: ReposAppRequest, res: any, next) => {
    return await modifySubscription(true, req, res, next);
  })
);

router.get(
  '/:campaignGroupId/subscribe',
  asyncHandler(async (req: ReposAppRequest, res: any, next) => {
    return await modifySubscription(false, req, res, next);
  })
);

router.get(
  '/:campaignGroupId',
  asyncHandler(async (req: ReposAppRequest, res: any, next) => {
    const { campaignStateProvider } = getProviders(req);
    if (!campaignStateProvider) {
      return next(
        new Error('This app is not configured for campaign management')
      );
    }
    const { campaignGroupId } = req.params;
    if (!campaignGroupId) {
      return next(new Error('Campaign required'));
    }
    const corporateId = req.individualContext.corporateIdentity.id;
    if (!corporateId) {
      return next(new Error('Corporate authentcation and identity required'));
    }
    const currentState = await campaignStateProvider.getState(
      corporateId,
      campaignGroupId
    );
    return res.json(currentState);
  })
);

async function modifySubscription(
  isUnsubscribing: boolean,
  req: ReposAppRequest,
  res: any,
  next: any
) {
  const { campaignStateProvider } = getProviders(req);
  if (!campaignStateProvider) {
    return next(
      new Error('This app is not configured for campaign management')
    );
  }
  const actionName = isUnsubscribing ? 'unsubscribe' : 'subscribe';
  const { campaignGroupId } = req.params;
  if (!campaignGroupId) {
    return next(new Error(`Campaign required to ${actionName}`));
  }
  const corporateId = req.individualContext.corporateIdentity.id;
  if (!corporateId) {
    return next(new Error('Corporate authentcation and identity required'));
  }
  const currentState = await campaignStateProvider.getState(
    corporateId,
    campaignGroupId
  );
  if (currentState && currentState.optOut && isUnsubscribing) {
    req.individualContext.webContext.saveUserAlert(
      `You've already opted out of the ${campaignGroupId} campaign.`,
      'Opt-out',
      UserAlertType.Success
    );
  } else if (currentState && !currentState.optOut && !isUnsubscribing) {
    req.individualContext.webContext.saveUserAlert(
      `You're not opted out of the ${campaignGroupId} campaign.`,
      'Re-subscribe',
      UserAlertType.Success
    );
  } else {
    if (isUnsubscribing) {
      await campaignStateProvider.optOut(corporateId, campaignGroupId);
      req.individualContext.webContext.saveUserAlert(
        `You've opted out of the ${campaignGroupId} campaign.`,
        'Opt-out',
        UserAlertType.Success
      );
    } else {
      await campaignStateProvider.clearOptOut(corporateId, campaignGroupId);
      req.individualContext.webContext.saveUserAlert(
        `You've subscribed to the ${campaignGroupId} campaign.`,
        'Re-subscribe',
        UserAlertType.Success
      );
    }
  }
  return res.redirect('/');
}

export default router;
