//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IReposError } from '../../../transitional';

export default function middlewareTeamAdminRequired(req, res, next) {
  const teamPermissions = req.teamPermissions;
  if (!teamPermissions) {
    return next(new Error('No team permissions information available'));
  }

  if (teamPermissions.allowAdministration === true) {
    return next();
  }

  const err: IReposError = new Error(
    'You do not have permission to administer this team'
  );
  err.status = 403;
  return next(err);
}
