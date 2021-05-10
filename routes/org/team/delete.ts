//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
const router: Router = Router();

import { ReposAppRequest, UserAlertType } from '../../../interfaces';

import MiddlewareTeamAdminRequired from './teamAdminRequired';

interface ILocalRequest extends ReposAppRequest {
  team2?: any;
}

router.post('/', MiddlewareTeamAdminRequired, (req: ILocalRequest, res, next) => {
  const organization = req.organization;
  const team2 = req.team2;
  team2.delete(error => {
    if (error) {
      return next(error);
    }
    req.individualContext.webContext.saveUserAlert(`${team2.name} team deleted`, 'Delete', UserAlertType.Success);
    res.redirect('/' + organization.name + '/teams');
  });
});

export default router;
