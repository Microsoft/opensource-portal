//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import _ from 'lodash';
import asyncHandler from 'express-async-handler';
import express from 'express';

import { getProviders, ReposAppRequest } from '../transitional';
import {
  Operations,
  ICrossOrganizationMembershipByOrganization,
} from '../business';
import { Team } from '../business';
import { UserContext } from '../user/aggregate';

import TeamSearch from '../business/teamSearch';

function sortOrgs(orgs) {
  return _.sortBy(orgs, ['name']);
}

interface IGetTeamsDataResults {
  teams: Team[];
  yourTeamsMap: Map<string, string>;
  totalMemberships: number;
  totalMaintainerships: number;
}

async function getTeamsData(
  singleOrganizationName: string | null,
  operations: Operations,
  userContext: UserContext
): Promise<IGetTeamsDataResults> {
  const options = {
    backgroundRefresh: true,
    maxAgeSeconds: 60 * 10 /* 10 minutes */,
    individualMaxAgeSeconds: 60 * 30 /* 30 minutes */,
  };
  let list: Team[] = null;
  if (singleOrganizationName) {
    const organization = operations.getOrganization(singleOrganizationName);
    list = await organization.getTeams(options);
  } else {
    list = [];
    const crossOrgTeams = await operations.getCrossOrganizationTeams(options);
    const allReducedTeams = Array.from(crossOrgTeams.values());
    allReducedTeams.forEach(
      (reducedTeam: ICrossOrganizationMembershipByOrganization) => {
        const orgs = Object.getOwnPropertyNames(reducedTeam.orgs);
        const firstOrg = orgs[0];
        const organization = operations.getOrganization(firstOrg);
        const entry = organization.teamFromEntity(reducedTeam.orgs[firstOrg]);
        list.push(entry);
      }
    );
  }

  const yourTeamsMap = new Map();
  const overview = await userContext.getAggregatedOverview();
  if (overview.teams && overview.teams.member.length) {
    reduceTeams(overview.teams, 'member', yourTeamsMap);
    reduceTeams(overview.teams, 'maintainer', yourTeamsMap);
  }
  return {
    teams: list,
    yourTeamsMap,
    totalMemberships:
      overview.teams && overview.teams.member
        ? overview.teams.member.length
        : 0,
    totalMaintainerships:
      overview.teams && overview.teams.maintainer
        ? overview.teams.maintainer.length
        : 0,
  };
}

function reduceTeams(collections, property, map) {
  if (!collections) {
    return;
  }
  const values = collections[property];
  values.forEach((team) => {
    map.set(team.id, property);
  });
}

export default asyncHandler(async function (
  req: ReposAppRequest,
  res: express.Response,
  next: express.NextFunction
) {
  const { operations } = getProviders(req);
  const isCrossOrg = req.teamsPagerMode === 'orgs';
  const aggregations = req.individualContext.aggregations;
  const orgName = isCrossOrg ? null : req.organization.name.toLowerCase();
  const {
    teams,
    yourTeamsMap,
    totalMemberships,
    totalMaintainerships,
  } = await getTeamsData(
    isCrossOrg ? null : orgName.toLowerCase(),
    operations,
    aggregations
  );
  const page = req.query.page_number ? Number(req.query.page_number) : 1;
  let phrase = req.query.q;

  let set = req.query.set;
  if (set !== 'all' && set !== 'available' && set !== 'your') {
    set = 'all';
  }

  const filters = [];
  if (phrase) {
    filters.push({
      type: 'phrase',
      value: phrase,
      displayPrefix: 'matching',
    });
  }

  const search = new TeamSearch(teams, {
    phrase: phrase,
    set: set,
    yourTeamsMap: yourTeamsMap,
  });
  await search.search(null, page, req.query.sort);

  const onboardingOrJoining = req.query.joining || req.query.onboarding;
  req.individualContext.webContext.render({
    view: 'teams/',
    title: 'Teams',
    state: {
      organizations: isCrossOrg
        ? sortOrgs(operations.getOrganizations(operations.organizationNames))
        : undefined,
      organization: isCrossOrg ? undefined : req.organization,
      search,
      filters,
      query: {
        phrase,
        set,
      },
      yourTeamsMap,
      totalMemberships,
      onboarding: req.query.onboarding,
      totalMaintainerships,
      errorAsWarning: null, // warning /* if an error occurs that is not fatal, we may want to display information about it */,
      onboardingOrJoining,
    },
  });
});
