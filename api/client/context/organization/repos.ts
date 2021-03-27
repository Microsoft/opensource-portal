//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
import { OrganizationMembershipState, Repository } from '../../../../business';
import { jsonError } from '../../../../middleware';
import { setContextualRepository } from '../../../../middleware/github/repoPermissions';

import { ReposAppRequest } from '../../../../transitional';
import { IndividualContext } from '../../../../user';
import { createRepositoryFromClient } from '../../newOrgRepo';

import RouteContextualRepo from './repo';

const router = express.Router();

async function validateActiveMembership(req: ReposAppRequest, res, next) {
  const { organization } = req;
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  if (!activeContext.link) {
    return next(
      jsonError(
        'You must be linked and a member of the organization to create and manage repos',
        400
      )
    );
  }
  const membership = await organization.getOperationalMembership(
    activeContext.getGitHubIdentity().username
  );
  if (!membership || membership.state !== OrganizationMembershipState.Active) {
    return next(
      jsonError(
        'You must be a member of the organization to create and manage repos',
        400
      )
    );
  }
  req['knownRequesterMailAddress'] = activeContext.link.corporateMailAddress;
  return next();
}

router.post(
  '/',
  asyncHandler(validateActiveMembership),
  asyncHandler(createRepositoryFromClient)
);

router.use(
  '/:repoName',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { organization } = req;
    const { repoName } = req.params;
    let repository: Repository = null;
    repository = organization.repository(repoName);
    setContextualRepository(req, repository);
    return next();
  })
);

router.use('/:repoName', RouteContextualRepo);

router.use('*', (req, res, next) => {
  return next(jsonError('no API or function available for repos', 404));
});

export default router;
