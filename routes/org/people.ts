//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import { ReposAppRequest } from '../../transitional';
const router = express.Router();

import RoutePeopleSearch from '../peopleSearch';

router.use(function (req: ReposAppRequest, res, next) {
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
