//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
const router = express.Router();

import { getProviders, ReposAppRequest } from '../transitional';

import RoutePeopleSearch from './peopleSearch';
import MiddlewareSystemWidePermissions from '../middleware/github/systemWidePermissions';

router.use(function (req: ReposAppRequest, res, next) {
  req.individualContext.webContext.pushBreadcrumb('People');
  req.reposContext = {
    section: 'people',
    pivotDirectlyToOtherOrg: '/people/', // hack
  };
  next();
});

// Campaign-related redirect to take the user to GitHub
router.get('/github/:login', (req: ReposAppRequest, res, next) => {
  const providers = getProviders(req);
  if (!providers || !providers.campaign) {
    return next();
  }
  return providers.campaign.redirectGitHubMiddleware(req, res, next, () => {
    const login = req.params.login;
    return login ? login : null;
  });
});

router.use(MiddlewareSystemWidePermissions);

router.use(RoutePeopleSearch);

export default router;
