//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import { jsonError } from '../../../middleware/jsonError';
import { ReposAppRequest } from '../../../interfaces';

const router: Router = Router();

router.get('/', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { organization } = req;
  const metadata = organization.getRepositoryCreateMetadata();
  res.json(metadata);
}));

router.get('/byProjectReleaseType', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { organization } = req;
  const options = {
    projectType: req.query.projectType,
  };
  const metadata = organization.getRepositoryCreateMetadata(options);
  res.json(metadata);
}));

router.use('*', (req, res, next) => {
  return next(jsonError('no API or function available within this path', 404));
});

export default router;
