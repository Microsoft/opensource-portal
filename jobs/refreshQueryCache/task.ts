//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import throat from 'throat';
import { shuffle } from 'lodash';

import { IReposJob, IReposJobResult } from '../../app';
import {
  ICacheOptions,
  IPagedCacheOptions,
  permissionsObjectToValue,
} from '../../transitional';
import {
  Team,
  TeamMember,
  GitHubTeamRole,
  Operations,
  Organization,
  OrganizationMembershipRoleQuery,
  OrganizationMembershipRole,
} from '../../business';
import {
  Repository,
  IGetCollaboratorsOptions,
  GitHubCollaboratorAffiliationQuery,
  GitHubCollaboratorType,
} from '../../business';
import { Collaborator } from '../../business';
import { TeamPermission } from '../../business';
import { OrganizationMember } from '../../business';
import { sleep, addArrayToSet } from '../../utils';
import QueryCache, {
  IQueryCacheOrganizationMembership,
  IQueryCacheTeam,
  IQueryCacheRepository,
  IQueryCacheTeamRepositoryPermission,
  IQueryCacheRepositoryCollaborator,
  QueryCacheOperation,
  IQueryCacheTeamMembership,
} from '../../business/queryCache';

interface IConsistencyStats {
  new: number;
  update: number;
  delete: number;
}

interface IRefreshOrganizationResults {
  organizationName: string;
  refreshSet: string;
  started: Date;
  finished: Date;
  consistencyStats: IConsistencyStats;
}

const slowRequestCacheOptions: IPagedCacheOptions = {
  maxAgeSeconds: 60, // 1m
  backgroundRefresh: false,
  pageRequestDelay: 1050, // just over a second
};

const minuteAgoCache: ICacheOptions = {
  maxAgeSeconds: 60,
  backgroundRefresh: false,
};

const sleepBetweenSteps = 100; // mostly impacts repo collaborator views

// TODO: we mark when the slow walk starts for the process, use it to set the max age value when requesting entities (?)
// TODO: make max age seconds accept a function that could make it dynamic and functional

// removed during refactor:
// insights.trackMetric({ name: 'JobRefreshQueryCacheMinutes', value: minutes });

async function refreshOrganization(
  organizationIndex: number,
  operations: Operations,
  refreshSet: string,
  queryCache: QueryCache,
  organization: Organization
): Promise<IRefreshOrganizationResults> {
  const result: IRefreshOrganizationResults = {
    organizationName: organization.name,
    refreshSet: refreshSet,
    started: new Date(),
    finished: null,
    consistencyStats: {
      delete: 0,
      new: 0,
      update: 0,
    },
  };

  let organizationDetails = null;
  try {
    organizationDetails = await organization.getDetails();
  } catch (organizationError) {
    console.log(
      `Organization get details error: ${organizationError} for org ${organization.name}`
    );
    console.dir(organizationError);
    return;
  }
  const organizationId = organizationDetails.id.toString();
  console.log(
    `refreshing ${organization.name} (id=${organizationId}) organization...`
  );

  if (refreshSet === 'all' || refreshSet === 'organizations') {
    try {
      const organizationAdmins = await organization.getMembers({
        ...slowRequestCacheOptions,
        role: OrganizationMembershipRoleQuery.Admin,
      });
      updateConsistencyStats(
        result.consistencyStats,
        await cacheOrganizationMembers(
          queryCache,
          organizationId,
          organizationAdmins,
          OrganizationMembershipRole.Admin
        )
      );
      const memberIds = new Set<string>(
        organizationAdmins.map((admin) => admin.id.toString())
      );
      await sleep(sleepBetweenSteps);

      const organizationMembers = await organization.getMembers({
        ...slowRequestCacheOptions,
        role: OrganizationMembershipRoleQuery.Member,
      });
      console.log(
        `${organizationIndex}: organization ${organization.name} has ${organizationAdmins.length} admins and ${organizationMembers.length} members`
      );
      updateConsistencyStats(
        result.consistencyStats,
        await cacheOrganizationMembers(
          queryCache,
          organizationId,
          organizationMembers,
          OrganizationMembershipRole.Member
        )
      );
      organizationMembers.map((member) => memberIds.add(member.id.toString()));
      await sleep(sleepBetweenSteps);

      // Cleanup any former members
      const cachedMembers = await queryCache.organizationMembers(
        organizationId
      );
      const potentialFormerMembers: IQueryCacheOrganizationMembership[] = [];
      cachedMembers.map((cm) => {
        if (!memberIds.has(cm.userId)) {
          potentialFormerMembers.push(cm);
        }
      });
      updateConsistencyStats(
        result.consistencyStats,
        await cleanupFormerMembers(
          operations,
          queryCache,
          organization,
          potentialFormerMembers
        )
      );
    } catch (orgMembersError) {
      console.log(
        `refresh for organization ${organization.name} members was interrupted by an error`
      );
      console.dir(orgMembersError);
    }
  }

  if (refreshSet === 'all' || refreshSet === 'teams') {
    try {
      const teams = await organization.getTeams(slowRequestCacheOptions);
      console.log(
        `${organizationIndex}: organization ${organization.name} has ${teams.length} teams`
      );
      const knownTeams = new Set(teams.map((team) => team.id.toString()));
      await sleep(sleepBetweenSteps);

      for (let i = 0; i < teams.length; i++) {
        try {
          const team = teams[i];
          const teamDetailsData = await team.getDetails(minuteAgoCache);
          updateConsistencyStats(
            result.consistencyStats,
            await queryCache.addOrUpdateTeam(
              organizationId,
              team.id.toString(),
              teamDetailsData
            )
          );

          const teamMaintainers = await team.getMaintainers(
            slowRequestCacheOptions
          );
          const maintainers = new Set<number>(
            teamMaintainers.map((maintainer) => maintainer.id)
          );
          updateConsistencyStats(
            result.consistencyStats,
            await cacheTeamMembers(
              queryCache,
              organizationId,
              team,
              teamMaintainers,
              GitHubTeamRole.Maintainer
            )
          );

          const teamMembers = await team.getMembers(slowRequestCacheOptions);
          const knownTeamMembers = new Set(
            teamMembers.map((member) => member.id.toString())
          );
          const nonMaintainerMembers = teamMembers.filter(
            (member) => !maintainers.has(member.id)
          );
          updateConsistencyStats(
            result.consistencyStats,
            await cacheTeamMembers(
              queryCache,
              organizationId,
              team,
              nonMaintainerMembers,
              GitHubTeamRole.Member
            )
          );
          console.log(
            `${organizationIndex}: team ${i + 1}/${teams.length}: ${
              team.name
            } from org ${organization.name} has ${
              teamMaintainers.length
            } maintainers and ${nonMaintainerMembers.length} members`
          );

          // Cleanup any removed team members
          const cachedTeamMembers = await queryCache.teamMembers(
            team.id.toString()
          );
          const removedMembers: IQueryCacheTeamMembership[] = [];
          cachedTeamMembers.map((ctm) => {
            if (!knownTeamMembers.has(ctm.userId)) {
              removedMembers.push(ctm);
            }
          });
          updateConsistencyStats(
            result.consistencyStats,
            await cleanupRemovedTeamMembers(queryCache, team, removedMembers)
          );
          await sleep(sleepBetweenSteps);
        } catch (teamError) {
          console.log(
            `issue processing team ${teams[i].id} in org ${organization.name}`
          );
          console.dir(teamError);
          await sleep(sleepBetweenSteps);
        }
      }

      // Cleanup any removed teams
      const cachedTeams = await queryCache.organizationTeams(organizationId);
      const potentialFormerTeams: IQueryCacheTeam[] = [];
      cachedTeams.map((ct) => {
        if (!knownTeams.has(ct.team.id.toString())) {
          potentialFormerTeams.push(ct);
        }
      });
      updateConsistencyStats(
        result.consistencyStats,
        await cleanupFormerTeams(queryCache, organization, potentialFormerTeams)
      );
    } catch (refreshTeamsError) {
      console.log(`error while refreshing teams in ${organization.name} org`);
      console.dir(refreshTeamsError);
    }
  }

  if (
    refreshSet === 'all' ||
    refreshSet === 'collaborators' ||
    refreshSet === 'permissions'
  ) {
    const repositories = await organization.getRepositories(
      slowRequestCacheOptions
    );
    console.log(
      `${organizationIndex}: ${repositories.length} repositories in ${organization.name}`
    );
    const repoIds = new Set(repositories.map((repo) => repo.id.toString()));

    for (let i = 0; i < repositories.length; i++) {
      try {
        const repository = repositories[i];
        const repoDetailsData = await repository.getDetails(minuteAgoCache);
        updateConsistencyStats(
          result.consistencyStats,
          await queryCache.addOrUpdateRepository(
            organizationId,
            repository.id.toString(),
            repoDetailsData
          )
        );
        await sleep(sleepBetweenSteps);

        if (refreshSet === 'all' || refreshSet === 'permissions') {
          const repoTeamPermissions = await repository.getTeamPermissions(
            slowRequestCacheOptions
          );
          updateConsistencyStats(
            result.consistencyStats,
            await cacheRepositoryTeams(
              queryCache,
              repository,
              repoTeamPermissions
            )
          );
          const knownTeamPermissions = new Set(
            repoTeamPermissions.map((rtp) => rtp.team.id.toString())
          );
          await sleep(sleepBetweenSteps);
          // Cleanup any removed team permissions
          const cachedTeamPermissions = await queryCache.repositoryTeamPermissions(
            repository.id.toString()
          );
          const removedPermissions: IQueryCacheTeamRepositoryPermission[] = [];
          cachedTeamPermissions.map((ctp) => {
            if (!knownTeamPermissions.has(ctp.team.id.toString())) {
              removedPermissions.push(ctp);
            }
          });
          updateConsistencyStats(
            result.consistencyStats,
            await cleanupRemovedTeamPermissions(
              queryCache,
              repository,
              removedPermissions
            )
          );
        }

        if (refreshSet === 'all' || refreshSet === 'collaborators') {
          const outsideOptions: IGetCollaboratorsOptions = {
            ...slowRequestCacheOptions,
            affiliation: GitHubCollaboratorAffiliationQuery.Outside,
          };
          const outsideRepoCollaborators = await repository.getCollaborators(
            outsideOptions
          );
          const collaboratorIds = new Set(
            outsideRepoCollaborators.map((orc) => orc.id.toString())
          );
          updateConsistencyStats(
            result.consistencyStats,
            await cacheRepositoryCollaborators(
              queryCache,
              organizationId,
              repository,
              outsideRepoCollaborators,
              GitHubCollaboratorType.Outside
            )
          );
          const outsideSet = new Set<number>(
            outsideRepoCollaborators.map((outsider) => outsider.id)
          );
          const directOptions: IGetCollaboratorsOptions = {
            ...slowRequestCacheOptions,
            affiliation: GitHubCollaboratorAffiliationQuery.Direct,
          };
          const directRepoCollaborators = await repository.getCollaborators(
            directOptions
          );
          directRepoCollaborators.map((drc) =>
            collaboratorIds.add(drc.id.toString())
          );
          const insideDirectCollaborators = directRepoCollaborators.filter(
            (collaborator) => !outsideSet.has(collaborator.id)
          );
          // technically 'direct' is just those that are not outside collaborators
          updateConsistencyStats(
            result.consistencyStats,
            await cacheRepositoryCollaborators(
              queryCache,
              organizationId,
              repository,
              insideDirectCollaborators,
              GitHubCollaboratorType.Direct
            )
          );
          const cachedRepositoryCollaborators = await queryCache.repositoryCollaborators(
            repository.id.toString()
          );
          const formerCollaborators: IQueryCacheRepositoryCollaborator[] = [];
          cachedRepositoryCollaborators.map((crc) => {
            if (!collaboratorIds.has(crc.userId)) {
              formerCollaborators.push(crc);
            }
          });
          updateConsistencyStats(
            result.consistencyStats,
            await cleanupFormerCollaborators(
              queryCache,
              repository,
              formerCollaborators
            )
          );
          await sleep(sleepBetweenSteps);
        }
        console.log(
          `${organizationIndex}: repository ${i + 1}/${repositories.length}: ${
            repository.full_name
          } repository`
        );
      } catch (refreshRepositoryError) {
        console.dir(refreshRepositoryError);
        await sleep(sleepBetweenSteps);
      }
    }
    // Cleanup any deleted repos
    try {
      const cachedRepositories = await queryCache.organizationRepositories(
        organizationId
      );
      const deletedRepositories: IQueryCacheRepository[] = [];
      cachedRepositories.map((r) => {
        if (!repoIds.has(r.repository.id.toString())) {
          deletedRepositories.push(r);
        }
      });
      updateConsistencyStats(
        result.consistencyStats,
        await cleanupDeletedRepositories(
          queryCache,
          organization,
          deletedRepositories
        )
      );
    } catch (cleanError) {
      console.dir(cleanError);
      await sleep(sleepBetweenSteps);
    }
  }
  result.finished = new Date();
  return result;
}

async function cacheRepositoryTeams(
  queryCache: QueryCache,
  repository: Repository,
  repoTeamPermissions: TeamPermission[]
): Promise<QueryCacheOperation[]> {
  const ops = [];
  const organizationId = repository.organization.id.toString();
  const repositoryId = repository.id.toString();
  for (let teamPermission of repoTeamPermissions) {
    const teamId = teamPermission.team.id.toString();
    const permission = teamPermission.permission;
    const isPrivate = repository.private as boolean;
    const repoName = repository.name as string;
    ops.push(
      await queryCache.addOrUpdateTeamsPermission(
        organizationId,
        repositoryId,
        isPrivate,
        repoName,
        teamId,
        permission
      )
    );
  }
  return ops.filter((exists) => exists);
}

async function cacheTeamMembers(
  queryCache: QueryCache,
  organizationId: string,
  team: Team,
  members: TeamMember[],
  typeOfRole: GitHubTeamRole
): Promise<QueryCacheOperation[]> {
  const ops = [];
  const teamId = team.id.toString();
  for (let member of members) {
    const userId = member.id.toString();
    const login = member.login;
    const avatar = member.avatar_url;
    ops.push(
      await queryCache.addOrUpdateTeamMember(
        organizationId,
        teamId,
        userId,
        typeOfRole,
        login,
        avatar
      )
    );
  }
  return ops.filter((exists) => exists);
}

async function cacheOrganizationMembers(
  queryCache: QueryCache,
  organizationId: string,
  members: OrganizationMember[],
  memberRole: OrganizationMembershipRole
): Promise<QueryCacheOperation[]> {
  const ops = [];
  for (let member of members) {
    const userId = member.id.toString();
    ops.push(
      await queryCache.addOrUpdateOrganizationMember(
        organizationId,
        memberRole,
        userId
      )
    );
  }
  return ops.filter((exists) => exists);
}

async function cleanupRemovedTeamMembers(
  queryCache: QueryCache,
  team: Team,
  removedMembers: IQueryCacheTeamMembership[]
): Promise<QueryCacheOperation[]> {
  const ops = [];
  for (const { team, userId, login } of removedMembers) {
    try {
      const userInTeam = await team.getMembership(login, minuteAgoCache);
      if (!userInTeam) {
        ops.push(
          await queryCache.removeTeamMember(
            team.organization.id.toString(),
            team.id.toString(),
            userId
          )
        );
        console.log(
          `permission for user login=${login} user id=${userId} to team ${team.id} removed from query cache`
        );
      }
    } catch (removalError) {
      console.log(
        `error while trying to cleanup potential former team members from ${team.organization.name} org`
      );
      console.dir(removalError);
    }
  }
  return ops.filter((real) => real);
}

async function cleanupRemovedTeamPermissions(
  queryCache: QueryCache,
  repository: Repository,
  removedPermissions: IQueryCacheTeamRepositoryPermission[]
): Promise<QueryCacheOperation[]> {
  const ops = [];
  for (const { team, repository } of removedPermissions) {
    try {
      if (!repository.name) {
        await repository.getDetails(); // make sure the repo name is known
      }
      const teamManagesRepository = await repository.checkTeamManages(
        team.id.toString(),
        minuteAgoCache
      );
      if (!teamManagesRepository) {
        ops.push(
          await queryCache.removeRepositoryTeam(
            repository.organization.id.toString(),
            repository.id.toString(),
            team.id.toString()
          )
        );
        console.log(
          `permission for team ${team.id} removed from repository id=${repository.id} query cache`
        );
      }
    } catch (removalError) {
      console.log(
        `error while trying to cleanup potential former repo permissions from ${repository.organization.name} org -> ${removalError}`
      );
      console.dir(removalError);
    }
  }
  return ops.filter((real) => real);
}

async function cleanupDeletedRepositories(
  queryCache: QueryCache,
  organization: Organization,
  deletedRepositories: IQueryCacheRepository[]
): Promise<QueryCacheOperation[]> {
  const ops = [];
  const organizationId = organization.id.toString();
  try {
    for (const { repository } of deletedRepositories) {
      if (await repository.isDeleted()) {
        ops.push(
          await queryCache.removeRepository(
            organizationId,
            repository.id.toString()
          )
        );
        console.log(
          `former organization=${organizationId} repository id=${repository.id} removed from query cache`
        );
      }
    }
  } catch (removingDeletedRepositoryError) {
    console.log(
      `error while trying to cleanup potential former repos from ${organization.name} org`
    );
    console.dir(removingDeletedRepositoryError);
  }
  return ops.filter((real) => real);
}

async function cleanupFormerCollaborators(
  queryCache: QueryCache,
  repository: Repository,
  formerCollaborators: IQueryCacheRepositoryCollaborator[]
): Promise<QueryCacheOperation[]> {
  const ops = [];
  const repositoryId = repository.id.toString();
  for (const { userId, cacheEntity } of formerCollaborators) {
    try {
      const login = cacheEntity.login;
      const isCollaborator = await repository.checkCollaborator(
        login,
        minuteAgoCache
      );
      if (!isCollaborator) {
        ops.push(
          await queryCache.removeRepositoryCollaborator(
            repository.organization.id.toString(),
            repositoryId,
            userId
          )
        );
        console.log(
          `removed collaborator ${login} from repository id=${repository.id} query cache`
        );
      }
    } catch (removalError) {
      console.log(
        `error while trying to cleanup potential former collaborators from ${repository.name} in ${repository.organization.name} org`
      );
      console.dir(removalError);
    }
  }
  return ops.filter((real) => real);
}

async function cleanupFormerTeams(
  queryCache: QueryCache,
  organization: Organization,
  potentialFormerTeams: IQueryCacheTeam[]
): Promise<QueryCacheOperation[]> {
  const ops = [];
  const organizationId = organization.id.toString();
  try {
    for (const { team } of potentialFormerTeams) {
      if (await team.isDeleted()) {
        ops.push(
          await queryCache.removeOrganizationTeam(
            organizationId,
            team.id.toString()
          )
        );
        console.log(
          `former organization=${organizationId} team id=${team.id} removed from query cache`
        );
      }
    }
  } catch (removingFormerTeamsError) {
    console.log(
      `error while trying to cleanup potential former teams from ${organization.name} org teams`
    );
    console.dir(removingFormerTeamsError);
  }
  return ops.filter((real) => real);
}

async function cleanupFormerMembers(
  operations: Operations,
  queryCache: QueryCache,
  organization: Organization,
  potentialFormerMembers: IQueryCacheOrganizationMembership[]
): Promise<QueryCacheOperation[]> {
  const ops = [];
  const organizationId = organization.id.toString();
  try {
    for (const { userId } of potentialFormerMembers) {
      const account = operations.getAccount(userId);
      let confirmedFormer = false;
      if ((await account.isDeleted()) === true) {
        confirmedFormer = true;
      } else {
        const login = account.login;
        const operationalMembership = await organization.getOperationalMembership(
          login
        );
        if (!operationalMembership) {
          confirmedFormer = true;
        } else {
          console.log(
            `while looking to cleanup a former member ID ${userId} with login ${login}, operational membership indicated a status of role=${operationalMembership.role}, state=${operationalMembership.state}, it will be kept`
          );
          console.log();
        }
      }
      if (confirmedFormer) {
        ops.push(
          await queryCache.removeOrganizationMember(organizationId, userId)
        );
        console.log(
          `former organization=${organizationId} member id=${userId} removed from query cache`
        );
      }
    }
  } catch (removingFormerMembersError) {
    console.log(
      `error while trying to cleanup potential former members from ${organization.name} org members`
    );
    console.dir(removingFormerMembersError);
  }
  return ops.filter((real) => real);
}

async function cacheRepositoryCollaborators(
  queryCache: QueryCache,
  organizationId: string,
  repository: Repository,
  repoCollaborators: Collaborator[],
  typeOfCollaborator: GitHubCollaboratorType
): Promise<QueryCacheOperation[]> {
  const operations = [];
  const repositoryId = repository.id.toString();
  for (let collaborator of repoCollaborators) {
    const permission = permissionsObjectToValue(collaborator.permissions);
    operations.push(
      await queryCache.addOrUpdateCollaborator(
        organizationId,
        repositoryId,
        repository,
        repository.name,
        collaborator.id.toString(),
        collaborator.login,
        collaborator.avatar_url,
        permission,
        typeOfCollaborator
      )
    );
  }
  return operations.filter((real) => real);
}

export default async function refresh({
  providers,
  args,
}: IReposJob): Promise<IReposJobResult> {
  const operations = providers.operations as Operations;
  const insights = providers.insights;
  const repositoryCacheProvider = providers.repositoryCacheProvider;
  const queryCache = providers.queryCache;
  const teamCacheProvider = providers.teamCacheProvider;
  const teamMemberCacheProvider = providers.teamMemberCacheProvider;
  const repositoryCollaboratorCacheProvider =
    providers.repositoryCollaboratorCacheProvider;
  const repositoryTeamCacheProvider = providers.repositoryTeamCacheProvider;
  if (!repositoryCacheProvider) {
    throw new Error('repositoryCacheProvider required');
  }
  if (!repositoryCollaboratorCacheProvider) {
    throw new Error('repositoryCollaboratorCacheProvider required');
  }
  if (!teamCacheProvider) {
    throw new Error('teamCacheProvider required');
  }
  if (!teamMemberCacheProvider) {
    throw new Error('teamMemberCacheProvider required');
  }
  if (!repositoryTeamCacheProvider) {
    throw new Error('repositoryTeamCacheProvider required');
  }
  let refreshSet = 'all';
  if (args.length > 0) {
    switch (args[0]) {
      case 'teams':
        refreshSet = 'teams';
        break;
      case 'collaborators':
        refreshSet = 'collaborators';
        break;
      case 'permissions':
        refreshSet = 'permissions';
        break;
      case 'organizations':
        refreshSet = 'organizations';
        break;
      case 'all':
        refreshSet = 'all';
        break;
      default:
        throw new Error(`unsupported mode ${args[0]}`);
    }
  }
  const parallelWorkCount = 1;
  const orgs = shuffle(Array.from(operations.organizations.values()));
  const currentlyKnownOrgIds = new Set<string>(
    orgs.map((org) => String(org.id))
  );
  let organizationWorkerCount = 0;
  const allUpStats: IConsistencyStats = {
    delete: 0,
    new: 0,
    update: 0,
  };
  let processedOrgs = 0;
  const staticOrgs = orgs.filter((org) => org.hasDynamicSettings === false);
  const dynamicOrgs = orgs.filter((org) => org.hasDynamicSettings === true);
  async function organizationProcessed(
    organization: Organization
  ): Promise<void> {
    await sleep(sleepBetweenSteps);
    console.log(
      `organization ${++organizationWorkerCount}/${orgs.length}: refreshing ${
        organization.name
      }`
    );
    const orgResult = await refreshOrganization(
      organizationWorkerCount,
      operations,
      refreshSet,
      queryCache,
      organization
    );
    if (orgResult) {
      const resultsAsLog = { ...orgResult, ...orgResult.consistencyStats };
      delete resultsAsLog.consistencyStats;
      allUpStats['delete'] += orgResult.consistencyStats['delete'];
      allUpStats['update'] += orgResult.consistencyStats['update'];
      allUpStats['new'] += orgResult.consistencyStats['new'];
      insights.trackEvent({
        name: 'QueryCacheOrganizationConsistencyResults',
        properties: (resultsAsLog as any) as { [key: string]: string },
      });

      console.log('--------------------------------------------------');
      console.log(`${organization.name} processed - eventual consistency`);
      console.log(`${++processedOrgs} organizations visited in this group`);
      console.log(`Added: ${orgResult.consistencyStats['new']}`);
      console.log(`Removed: ${orgResult.consistencyStats['delete']}`);
      console.log(`Updated: ${orgResult.consistencyStats['update']}`);
      console.log('--------------------------------------------------');
    } else {
      console.log('--------------------------------------------------');
      console.log(`${organization.name} failed processing`);
      console.log(`${++processedOrgs} organizations visited in this group`);
      console.log('--------------------------------------------------');
    }
  }
  const isFastOK = refreshSet === 'teams' || refreshSet === 'organizations';
  const parallelDynamicCount = isFastOK ? dynamicOrgs.length / 2 : 5;
  console.log(
    `Parallel dynamic organization count: ${parallelDynamicCount} for ${dynamicOrgs.length} configured dynamic orgs`
  );
  try {
    console.log(
      `processing ${dynamicOrgs.length} dynamic orgs, ${parallelDynamicCount} at a time`
    );
    const dynamicThrottle = throat(parallelDynamicCount);
    await Promise.all(
      dynamicOrgs.map((org: Organization) =>
        dynamicThrottle(organizationProcessed.bind(null, org))
      )
    );
    processedOrgs = 0;
    console.log(
      `processing ${staticOrgs.length} static orgs, ${parallelWorkCount} at a time to avoid rate-limiting`
    );
    const staticThrottle = throat(parallelWorkCount);
    await Promise.all(
      staticOrgs.map((org: Organization) =>
        staticThrottle(organizationProcessed.bind(null, org))
      )
    );
  } catch (dynamicError) {
    console.dir(dynamicError);
  }

  let removedOrganizations = 0;
  if (
    queryCache &&
    queryCache.supportsOrganizationMembership &&
    (refreshSet === 'all' || refreshSet === 'organizations')
  ) {
    const knownOrgIds = new Set<string>();
    addArrayToSet(
      knownOrgIds,
      await queryCache.organizationMemberCacheOrganizationIds()
    );
    addArrayToSet(
      knownOrgIds,
      await queryCache.repositoryCacheOrganizationIds()
    );
    addArrayToSet(
      knownOrgIds,
      await queryCache.repositoryCollaboratorCacheOrganizationIds()
    );
    addArrayToSet(
      knownOrgIds,
      await queryCache.repositoryTeamOrganizationIds()
    );
    addArrayToSet(knownOrgIds, await queryCache.teamOrganizationIds());
    const unknownOrgs = Array.from(knownOrgIds.values()).filter(
      (id) => !currentlyKnownOrgIds.has(id)
    );
    if (unknownOrgs.length > 0) {
      for (const id of unknownOrgs) {
        try {
          console.log(`Unknown former organization ID: ${id}`);
          await queryCache.removeOrganizationById(id);
          ++removedOrganizations;
        } catch (processingIndividualOrgError) {
          console.dir(processingIndividualOrgError);
        }
      }
    }
  }

  console.log('--------------------------------------------------');
  console.log('All organizations processed, all-up results:');
  console.log(`Added:        ${allUpStats['new']}`);
  console.log(`Removed:      ${allUpStats['delete']}`);
  console.log(`Updated:      ${allUpStats['update']}`);
  console.log(`Removed orgs: ${removedOrganizations}`);
  console.log('--------------------------------------------------');
  insights.trackEvent({
    name: 'JobRefreshQueryCacheSuccess',
    properties: {
      allUpNew: allUpStats['new'].toString(),
      allUpDelete: allUpStats['delete'].toString(),
      allUpUpdate: allUpStats['update'].toString(),
    },
  });
  insights.trackMetric({
    name: 'QueryCacheConsistencyAdds',
    value: allUpStats['new'],
  });
  insights.trackMetric({
    name: 'QueryCacheConsistencyDeletes',
    value: allUpStats['delete'],
  });
  insights.trackMetric({
    name: 'QueryCacheConsistencyUpdates',
    value: allUpStats['update'],
  });
  return {
    successProperties: {
      adds: allUpStats['new'],
      deletes: allUpStats['delete'],
      updates: allUpStats['update'],
    },
  };
}

function updateConsistencyStats(
  stats: IConsistencyStats,
  outcomes: QueryCacheOperation | QueryCacheOperation[]
): void {
  let array: QueryCacheOperation[] = null;
  if (Array.isArray(outcomes)) {
    array = outcomes as QueryCacheOperation[];
  } else {
    array = [outcomes as QueryCacheOperation];
  }
  array.forEach((outcome) => {
    if (outcome === null) {
      return;
    }
    const stringKey = outcome as string;
    if (
      stats[stringKey] === undefined ||
      typeof stats[stringKey] !== 'number'
    ) {
      throw new Error(`invalid outcome ${stringKey}`);
    }
    ++stats[stringKey];
  });
}
