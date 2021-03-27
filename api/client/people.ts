//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';

import {
  corporateLinkToJson,
  ICorporateLink,
  ICrossOrganizationMembersResult,
  MemberSearch,
  Operations,
  Organization,
} from '../../business';
import { jsonError } from '../../middleware';
import { ReposAppRequest, IProviders } from '../../transitional';
import JsonPager from './jsonPager';
import getCompanySpecificDeployment from '../../middleware/companySpecificDeployment';

import RouteGetPerson from './person';
import { equivalentLegacyPeopleSearch } from './peopleSearch';

const router = express.Router();

const deployment = getCompanySpecificDeployment();
deployment?.routes?.api?.people && deployment.routes.api.people(router);

interface ISimpleAccount {
  login: string;
  avatar_url: string;
  id: number;
}

export interface ICrossOrganizationMemberResponse {
  account: ISimpleAccount;
  link?: ICorporateLink;
  organizations: string[];
}

export interface ICrossOrganizationSearchedMember {
  id: number;
  account: ISimpleAccount;
  link?: ICorporateLink;
  orgs: IOrganizationMembershipAccount;
}

interface IOrganizationMembershipAccount {
  [id: string]: ISimpleAccount;
}

router.get('/:login', RouteGetPerson);

router.get(
  '/',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const pager = new JsonPager<ICrossOrganizationSearchedMember>(req, res);
    try {
      const searcher = await equivalentLegacyPeopleSearch(req);
      const members = (searcher.members as unknown) as ICrossOrganizationSearchedMember[];
      const slice = pager.slice(members);
      return pager.sendJson(
        slice.map((xMember) => {
          const obj = Object.assign(
            {
              link: xMember.link ? corporateLinkToJson(xMember.link) : null,
              id: xMember.id,
              organizations: xMember.orgs
                ? Object.getOwnPropertyNames(xMember.orgs)
                : [],
            },
            xMember.account || { id: xMember.id }
          );
          return obj;
        })
      );
    } catch (repoError) {
      console.dir(repoError);
      return next(jsonError(repoError));
    }
  })
);

router.use('*', (req, res, next) => {
  return next(
    jsonError(
      'no API or function available within this cross-organization people list',
      404
    )
  );
});

export default router;
