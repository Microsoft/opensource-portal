//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
const router: Router = Router();

import { ReposAppRequest } from '../../interfaces';

import RoutePeopleSearch from '../peopleSearch';

router.use(function (req:ReposAppRequest, res, next) {
  req.individualContext.webContext.pushBreadcrumb('People');
  req.reposContext = {
    section: 'people',
    organization: req.organization,
    pivotDirectlyToOtherOrg: '/people/', // hack
  };
  next();
});

router.use(RoutePeopleSearch);

export default router;
