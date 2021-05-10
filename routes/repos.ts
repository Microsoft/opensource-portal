//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
const router: Router = Router();

import { ReposAppRequest } from '../interfaces';
import lowercaser from '../middleware/lowercaser';

import RouteReposPager from './reposPager';

router.use(function (req: ReposAppRequest, res, next) {
  req.individualContext.webContext.pushBreadcrumb('Repositories');
  req.reposContext = {
    section: 'repos',
    pivotDirectlyToOtherOrg: '/repos/', // hack
  };
  req.reposPagerMode = 'orgs';
  next();
});

router.get('/', lowercaser(['sort', 'language', 'type', 'tt']), RouteReposPager);

export default router;
