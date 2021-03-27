//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
const router = express.Router();
import asyncHandler from 'express-async-handler';

import { getProviders, ReposAppRequest } from '../../transitional';

interface IUserProfileWarnings {
  company?: string;
  email?: string;
}

router.get(
  '/',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const organization = req.organization;
    const { operations } = getProviders(req);
    const config = operations.config;
    const onboarding = req.query.onboarding;
    const login = req.individualContext.getGitHubIdentity().username;
    let detailed = null;
    try {
      detailed = await operations.getAccountByUsername(login);
    } catch (getAccountError) {
      return next(getAccountError);
    }
    const userProfileWarnings: IUserProfileWarnings = {};
    if (
      !detailed.company ||
      (detailed.company &&
        detailed.company
          .toLowerCase()
          .indexOf(config.brand.companyName.toLowerCase()) < 0)
    ) {
      userProfileWarnings.company = 'color:red';
    }
    if (
      !detailed.email ||
      (detailed.email &&
        detailed.email
          .toLowerCase()
          .indexOf(config.brand.companyName.toLowerCase()) < 0)
    ) {
      userProfileWarnings.email = 'color:red';
    }
    req.individualContext.webContext.render({
      view: 'org/profileReview',
      title: 'Your GitHub Profile',
      state: {
        organization,
        userProfile: detailed,
        userProfileWarnings,
        theirUsername: req.individualContext.getGitHubIdentity().username,
        onboarding,
        showBreadcrumbs: onboarding === undefined,
      },
    });
  })
);

export default router;
