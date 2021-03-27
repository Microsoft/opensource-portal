//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import asyncHandler from 'express-async-handler';
import express from 'express';
import _ from 'lodash';

import { IReposAppWithTeam, getProviders } from '../transitional';
import { Operations } from '../business';
import { Repository } from '../business';
import { Team, GitHubRepositoryType } from '../business';
import { RepositorySearch } from '../business/repoSearch';
import QueryCache from '../business/queryCache';
import { Organization } from '../business/organization';
import { IPersonalizedUserAggregateRepositoryPermission } from '../business/graphManager';
import { IRequestTeamPermissions } from '../middleware/github/teamPermissions';
import { UserContext } from '../user/aggregate';
import { daysInMilliseconds } from '../utils';
import { TeamRepositoryPermission } from '../business/teamRepositoryPermission';

interface IGetReposAndOptionalTeamPermissionsResponse {
  reposData: Repository[];
  ageInformation?: any;
  userRepos?: IPersonalizedUserAggregateRepositoryPermission[];
  specificTeamRepos?: TeamRepositoryPermission[];
}

function sortOrgs(orgs) {
  return _.sortBy(orgs, ['name']);
}

async function getRepos(
  organizationId: number,
  operations: Operations,
  queryCache: QueryCache
): Promise<Repository[]> {
  if (organizationId) {
    if (queryCache && queryCache.supportsRepositories) {
      return (
        await queryCache.organizationRepositories(organizationId.toString())
      ).map((wrapper) => wrapper.repository);
    } else {
      return operations.getOrganizationById(organizationId).getRepositories();
    }
  } else {
    if (queryCache && queryCache.supportsRepositories) {
      return (await queryCache.allRepositories()).map(
        (wrapper) => wrapper.repository
      );
    }
    return operations.getRepos();
  }
}

async function getReposAndOptionalTeamPermissions(
  organizationId: number,
  operations: Operations,
  queryCache: QueryCache,
  teamsType: string | null | undefined,
  team2: Team,
  specificTeamRepos,
  userContext: UserContext
): Promise<IGetReposAndOptionalTeamPermissionsResponse> {
  // REMOVED: previously age information was avialable via getRepos(orgName, operations, (error, reposData, ageInformation). Was it really useful?
  const reposData = await getRepos(organizationId, operations, queryCache);
  if (!teamsType || teamsType === 'all') {
    // Retrieve the repositories for this specific repo, along with permissions information
    // NOTE: This means that for now the filtering of permissions won't work in specific team mode
    if (team2 && specificTeamRepos) {
      const repoOptions = {
        type: GitHubRepositoryType.Sources,
      };
      return {
        reposData,
        specificTeamRepos: await team2.getRepositories(repoOptions),
      };
    } else {
      return { reposData };
    }
  }
  const userRepos = await userContext.repositoryPermissions();
  return { reposData, userRepos };
}

export default asyncHandler(async function (
  req: IReposAppWithTeam,
  res: express.Response,
  next: express.NextFunction
) {
  const providers = getProviders(req);
  const operations = providers.operations;
  const queryCache = providers.queryCache;
  const individualContext = req.individualContext;
  const isCrossOrg = req.reposPagerMode === 'orgs';
  let teamsType = req.query.tt as string;
  const organization = req.organization as Organization;
  const organizationId = isCrossOrg ? null : organization.id;
  // Filter by team repositories, only in sub-team views
  const specificTeamPermissions = req.teamPermissions as IRequestTeamPermissions;
  const team2 = req.team2 as Team;
  let specificTeamId = team2 ? team2.id : null;
  const {
    reposData,
    userRepos,
    specificTeamRepos,
  } = await getReposAndOptionalTeamPermissions(
    organizationId,
    operations,
    queryCache,
    teamsType,
    team2,
    specificTeamId,
    individualContext.aggregations
  );

  const page = req.query.page_number ? Number(req.query.page_number) : 1;

  let phrase = req.query.q as string;

  // TODO: Validate the type
  let type = req.query.type as string;
  if (
    type !== 'public' &&
    type !== 'private' &&
    type !== 'source' &&
    type !== 'fork' /*&& type !== 'mirrors' - we do not do mirror stuff */
  ) {
    type = null;
  }

  let metadataType = req.query.mt as string;
  if (
    metadataType !== 'with-metadata' &&
    metadataType !== 'without-metadata' &&
    metadataType !== 'administrator-locked' &&
    metadataType !== 'locked' &&
    metadataType !== 'unlocked'
  ) {
    metadataType = null;
  }

  const createdSinceValue = req.query.cs ? Number(req.query.cs) : null;
  let createdSince = null;
  if (createdSinceValue) {
    createdSince = new Date(
      new Date().getTime() - daysInMilliseconds(createdSinceValue)
    );
  }

  let showIds = req.query.showids === '1';

  let teamsSubType = null;
  if (
    teamsType !== 'myread' &&
    teamsType !== 'mywrite' &&
    teamsType !== 'myadmin'
  ) {
    teamsType = null;
  } else if (
    teamsType === 'myread' ||
    teamsType === 'mywrite' ||
    teamsType === 'myadmin'
  ) {
    teamsSubType = teamsType.substr(2);
    teamsType = 'my';
  }
  // TODO: Validate the language value is in the Linguist list
  let language = req.query.language as string;

  const filters = [];
  if (type) {
    filters.push({
      type: 'type',
      value: type,
      displayValue: type === 'fork' ? 'forked' : type,
      displaySuffix: 'repositories',
    });
  }
  if (phrase) {
    filters.push({
      type: 'phrase',
      value: phrase,
      displayPrefix: 'matching',
    });
  }
  if (language) {
    filters.push({
      type: 'language',
      value: language,
      displayPrefix: 'written in',
    });
  }
  if (teamsType) {
    let ttValue = teamsType === 'my' ? 'my ' + teamsSubType : teamsType;
    filters.push({
      type: 'tt',
      value: ttValue,
      displayPrefix: 'and',
      displaySuffix: 'team permissions',
    });
  }
  if (createdSince) {
    filters.push({
      type: 'cs',
      value: `${createdSinceValue} days`,
      displayPrefix: 'created within',
    });
  }
  if (metadataType) {
    const mtValue = metadataType.replace('-', ' ');
    filters.push({
      type: 'mt',
      value: mtValue,
    });
  }

  const search = new RepositorySearch(reposData, {
    phrase,
    language,
    type,
    teamsType,
    metadataType,
    specificTeamRepos,
    specificTeamPermissions,
    createdSince,
    teamsSubType,
    userRepos,
    graphManager: operations.graphManager,
    repositoryMetadataProvider: operations.providers.repositoryMetadataProvider,
  });

  await search.search(page, req.query.sort as string);

  // await Promise.all(search.repos.map(repo => repo.getDetails()));

  req.individualContext.webContext.render({
    view: 'repos/',
    title: 'Repos',
    state: {
      organizations: isCrossOrg
        ? sortOrgs(operations.getOrganizations(operations.organizationNames))
        : undefined,
      organization: isCrossOrg ? undefined : req.organization,
      search,
      filters,
      query: {
        phrase,
        type,
        language,
        tt: teamsType ? req.query.tt : null,
      },
      reposDataAgeInformation: null, // ageInformation, // TODO: can 'ageInformation' be recovered?
      specificTeamPermissions,
      specificTeam: team2,
      teamUrl: req.teamUrl,
      showIds,
    },
  });
});
