//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';

import {
  AddRepositoryPermissionsToRequest,
  getContextualRepositoryPermissions,
} from '../../../../middleware/github/repoPermissions';
import { jsonError } from '../../../../middleware';
import getCompanySpecificDeployment from '../../../../middleware/companySpecificDeployment';
import { ReposAppRequest } from '../../../../transitional';

const router = express.Router();

router.get(
  '/permissions',
  AddRepositoryPermissionsToRequest,
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const permissions = getContextualRepositoryPermissions(req);
    return res.json(permissions);
  })
);

const deployment = getCompanySpecificDeployment();
deployment?.routes?.api?.context?.organization?.repo &&
  deployment?.routes?.api?.context?.organization?.repo(router);

router.use('*', (req, res, next) => {
  return next(
    jsonError(`no API or ${req.method} function available for repo`, 404)
  );
});

export default router;
