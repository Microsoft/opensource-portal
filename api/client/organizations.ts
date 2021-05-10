//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import { jsonError } from '../../middleware';
import { ErrorHelper, getProviders } from '../../transitional';
import { ReposAppRequest } from '../../interfaces';

import RouteOrganization from './organization';

const router: Router = Router();

router.get('/', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { operations } = getProviders(req);
  try {
    const orgs = operations.getOrganizations();
    const dd = orgs.map(org => { return org.asClientJson(); });
    return res.json(dd);
  } catch (error) {
    throw jsonError(error, 400);
  }
}));

router.use('/:orgName', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { operations } = getProviders(req);
  const { orgName } = req.params;
  try {
    const org = operations.getOrganization(orgName);
    if (org) {
      req.organization = org;
      return next();
    }
    throw jsonError('managed organization not found', 404);
  } catch (orgNotFoundError) {
    if (ErrorHelper.IsNotFound(orgNotFoundError)) {
      return next(jsonError(orgNotFoundError, 404));
    } else {
      return next(jsonError(orgNotFoundError));
    }
  }
}));

router.use('/:orgName', RouteOrganization);

router.use('*', (req: ReposAppRequest, res, next) => {
  return next(jsonError('orgs API not found', 404));
});

export default router;
