//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import { apiContextMiddleware, AddLinkToRequest, requireAccessTokenClient, setIdentity, jsonError } from '../../middleware';
import { getProviders } from '../../transitional';

import getCompanySpecificDeployment from '../../middleware/companySpecificDeployment';

import RouteClientNewRepo from './newRepo';

import RouteContext from './context';
import RouteOrganizations from './organizations';
import RouteLinking from './linking';
import RouteSession from './session';
import RouteBanner from './banner';
import RouteCrossOrganizationPeople from './people';
import RouteCrossOrganizationRepos from './repos';
import RouteCrossOrganizationTeams from './teams';
import { ReposAppRequest } from '../../interfaces';

const router: Router = Router()

router.use((req: ReposAppRequest, res, next) => {
  const { config } = getProviders(req);
  if (config?.features?.allowApiClient) {
    return req.isAuthenticated() ? next() : next(jsonError('Session is not authenticated', 401));
  }
  return next(jsonError('Client API features unavailable', 403));
});

router.use(asyncHandler(requireAccessTokenClient));
router.use(apiContextMiddleware);
router.use(setIdentity);
router.use(asyncHandler(AddLinkToRequest));

router.use('/newRepo', RouteClientNewRepo);

router.use('/context', RouteContext);

router.use('/banner', RouteBanner);
router.use('/orgs', RouteOrganizations);
router.use('/link', RouteLinking);
router.use('/signout', RouteSession);
router.use('/people', RouteCrossOrganizationPeople);
router.use('/repos', RouteCrossOrganizationRepos);
router.use('/teams', RouteCrossOrganizationTeams);

const dynamicStartupInstance = getCompanySpecificDeployment();
dynamicStartupInstance?.routes?.api?.index && dynamicStartupInstance?.routes?.api?.index(router);

router.use((req, res, next) => {
  return next(jsonError('The resource or endpoint you are looking for is not there', 404));
});

export default router;
