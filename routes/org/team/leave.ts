//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { ReposAppRequest, UserAlertType } from '../../../interfaces';
import { Organization } from '../../../business/organization';
import { Team } from '../../../business/team';

interface ILocalRequest extends ReposAppRequest {
  team2?: any;
}

router.post('/', asyncHandler(async (req: ILocalRequest, res, next) => {
  const organization = req.organization as Organization;
  const team2 = req.team2 as Team;
  const username = req.individualContext.link.thirdPartyUsername;
  await team2.removeMembership(username);
  req.individualContext.webContext.saveUserAlert(`You've been successfully removed from ${team2.name}!`, 'Remove', UserAlertType.Success);
  return res.redirect('/' + organization.name + '/teams');
}));

export default router;
