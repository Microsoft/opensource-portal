//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import RouteRepos from './repos';
import RouteTeams from './teams';
import RoutePeople from './people';
import RouteNewRepoMetadata from './newRepoMetadata';
import { ReposAppRequest } from '../../../interfaces';
import { jsonError } from '../../../middleware';
import getCompanySpecificDeployment from '../../../middleware/companySpecificDeployment';

const router: Router = Router();

const deployment = getCompanySpecificDeployment();
deployment?.routes?.api?.organization?.index && deployment?.routes?.api?.organization?.index(router);

router.get('/', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { organization } = req;
  return res.json(organization.asClientJson());
}));

router.get('/accountDetails', asyncHandler(async (req: ReposAppRequest, res) => {
  const { organization } = req;
  const entity = organization.getEntity();
  if (entity) {
    return res.json(entity);
  }
  const details = await organization.getDetails();
  return res.json(details);
}));

router.use('/repos', RouteRepos);
router.use('/teams', RouteTeams);
router.use('/people', RoutePeople);
router.use('/newRepoMetadata', RouteNewRepoMetadata);

router.use('*', (req, res, next) => {
  return next(jsonError('no API or function available', 404));
});

export default router;
