//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';

import { jsonError } from '../../../middleware/jsonError';
import { ReposAppRequest } from '../../../transitional';
import { Repository } from '../../../business/repository';
import { findRepoCollaboratorsExcludingOwners } from '../../../routes/org/repos';

type RequestWithRepo = ReposAppRequest & {
  repository: Repository;
};

const router = express.Router();

router.get(
  '/',
  asyncHandler(async (req: RequestWithRepo, res, next) => {
    const { repository, organization } = req;
    try {
      const teamPermissions = await repository.getTeamPermissions();
      const owners = await organization.getOwners();
      const {
        collaborators,
        outsideCollaborators,
        memberCollaborators,
      } = await findRepoCollaboratorsExcludingOwners(repository, owners);
      for (let teamPermission of teamPermissions) {
        try {
          teamPermission.resolveTeamMembers();
        } catch (ignoredError) {
          /* ignored */
        }
      }
      return res.json({
        teamPermissions: teamPermissions.map((tp) => tp.asJson()),
        collaborators: collaborators.map((c) => c.asJson()),
        outsideCollaborators: outsideCollaborators.map((oc) => oc.asJson()),
        memberCollaborators: memberCollaborators.map((oc) => oc.asJson()),
      });
    } catch (error) {
      return next(jsonError(error));
    }
  })
);

export default router;
