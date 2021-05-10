//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
import { Team } from '../../../../business';
import { jsonError } from '../../../../middleware';
import { setContextualTeam } from '../../../../middleware/github/teamPermissions';
import { ReposAppRequest } from '../../../../interfaces';

import RouteTeam from './team';

const router: Router = Router();

// CONSIDER: list their teams router.get('/ ')

router.use('/:teamSlug', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { organization } = req;
  const { teamSlug } = req.params;
  let team: Team = null;
  try {
    team = await organization.getTeamFromSlug(teamSlug);
    setContextualTeam(req, team);
  } catch (error) {
    console.dir(error);
    return next(error);
  }
  return next();
}));

router.use('/:teamSlug', RouteTeam);

router.use('*', (req, res, next) => {
  return next(jsonError('no API or function available for repos', 404));
});

export default router;
