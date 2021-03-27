//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import asyncHandler from 'express-async-handler';

import { ReposAppRequest } from '../../../transitional';
import { IndividualContext } from '../../../user';
import { GitHubRepositoryPermission } from '../../../entities/repositoryMetadata/repositoryMetadata';

export default asyncHandler(async (req: ReposAppRequest, res) => {
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  if (!activeContext.link) {
    return res.json({
      isLinked: false,
      repositories: [],
    });
  }
  let permissions = await activeContext.aggregations.getQueryCacheRepositoryPermissions();
  permissions = permissions.filter((perm) => {
    if (perm.bestComputedPermission !== GitHubRepositoryPermission.Pull) {
      return true;
    }
    let fromBroadAccess = false;
    perm.teamPermissions.map((tp) => {
      if (tp.team.isBroadAccessTeam) {
        fromBroadAccess = true;
      }
    });
    if (fromBroadAccess) {
      return false;
    }
    if (perm.repository.private) {
      return true;
    }
    return false;
  });
  return res.json({
    isLinked: true,
    repositories: permissions.map((perm) => {
      return {
        bestComputedPermission: perm.bestComputedPermission,
        collaboratorPermission: perm.collaboratorPermission,
        repository: perm.repository.asJson(),
        teamPermissions: perm.teamPermissions.map((tp) => tp.asJson()),
        // TODO: would be nice for team permission for repos to also store the team slug in the query cache!
      };
    }),
  });
});
