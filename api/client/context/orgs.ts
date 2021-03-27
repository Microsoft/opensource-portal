//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import asyncHandler from 'express-async-handler';

import { ReposAppRequest } from '../../../transitional';
import { IndividualContext } from '../../../user';

export default asyncHandler(async (req: ReposAppRequest, res) => {
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  if (!activeContext.link) {
    return res.json({
      member: [],
      admin: [],
      isLinked: false,
    });
  }
  const orgs = await activeContext.aggregations.getQueryCacheOrganizations();
  const data = {
    isLinked: true,
    member: orgs.member.map((org) => {
      return {
        name: org.name,
        id: org.id,
      };
    }),
    admin: orgs.admin.map((org) => {
      return {
        name: org.name,
        id: org.id,
      };
    }),
  };
  return res.json(data);
});
