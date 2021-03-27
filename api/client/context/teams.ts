//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import asyncHandler from 'express-async-handler';

import { ReposAppRequest } from '../../../transitional';
import { IndividualContext } from '../../../user';
import { TeamJsonFormat } from '../../../business';

export default asyncHandler(async (req: ReposAppRequest, res, next) => {
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  if (!activeContext.link) {
    return res.json({
      isLinked: false,
      member: [],
      maintainer: [],
    });
  }
  const permissions = await activeContext.aggregations.getQueryCacheTeams();
  return res.json({
    isLinked: true,
    member: permissions.member.map((t) => t.asJson(TeamJsonFormat.Augmented)),
    maintainer: permissions.maintainer.map((t) =>
      t.asJson(TeamJsonFormat.Augmented)
    ),
  });
});
