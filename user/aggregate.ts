//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Operations } from '../business';
import {
  SettleToStateValue,
  isPermissionBetterThan,
  ErrorHelper,
} from '../transitional';

import LinkManager from './linkManager';
import { Team, GitHubTeamRole } from '../business/team';
import {
  Organization,
  OrganizationMembershipRoleQuery,
  OrganizationMembershipRole,
} from '../business/organization';
import { Repository } from '../business/repository';
import QueryCache, {
  IQueryCacheTeamRepositoryPermission,
} from '../business/queryCache';
import { IPersonalizedUserAggregateRepositoryPermission } from '../business/graphManager';
import { TeamRepositoryPermission } from '../business/teamRepositoryPermission';
import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';

// PLANNING once consolidated
// TODO: expose a combined 'teams in' which is maint + member
// TODO: how to expose a -pending- org invite membership waiting?
// TODO: expose a 'in' org that is owner and member of
// TODO: object should expose each aspect as promises, allowing intra-middleware smart use & reuse in a request
// TODO: remove 'linkmanager' concept

export interface IAggregateUserOrganizations
  extends IKnownAggregateUserOrganizations {
  available: Organization[];
}

export interface IKnownAggregateUserOrganizations {
  member: Organization[];
  admin: Organization[];
  // in: TODO
  // pending: TODO
}

export interface IAggregateUserTeams {
  member: Team[];
  maintainer: Team[];
  // in: TODO
}

// TODO: implement these newer ones
export interface IAggregateUserRepositories {
  pull: Repository[];
  push: Repository[];
  admin: Repository[];
  // TODO: the newer types, also...
}

export interface IAggregateLegacyUserRespositories {
  byTeam: Repository[];
  // byCollaboration: was originally designed but was too slow to work
}

export interface IAggregateUserSummary {
  organizations: IAggregateUserOrganizations;
  teams: IAggregateUserTeams;
  // TODO: rename repos to repositories
  repos: IAggregateLegacyUserRespositories;
}

interface IRepositoryPermissionPair {
  repository: Repository;
  collaborationPermission: GitHubRepositoryPermission;
  teamPermissions: IQueryCacheTeamRepositoryPermission[];
}

// TODO: rename to UserAggregations class
export class UserContext {
  private _operations: Operations;
  private _queryCache: QueryCache;
  private _linkManager: any;
  private _organizations: IAggregateUserOrganizations;
  private _legacyRepositories: IAggregateLegacyUserRespositories;
  private _teams: IAggregateUserTeams;
  private _repositoryPermissions: IPersonalizedUserAggregateRepositoryPermission[];

  public id: number;

  constructor(
    operations: Operations,
    queryCache: QueryCache,
    id: string | number
  ) {
    this.id = typeof id === 'string' ? parseInt(id, 10) : id;

    this._operations = operations;
    this._queryCache = queryCache;
  }

  get linkManager() {
    if (!this._linkManager) {
      this._linkManager = new LinkManager(this._operations, this);
    }
    return this._linkManager;
  }

  async organizations(): Promise<IAggregateUserOrganizations> {
    if (this._organizations) {
      return this._organizations;
    }
    this._organizations = await this.aggregateOrganizations();
    return this._organizations;
  }

  async teams(): Promise<IAggregateUserTeams> {
    if (this._teams) {
      return this._teams;
    }
    this._teams = await this.aggregateTeams();
    return this._teams;
  }

  async repositories(): Promise<IAggregateLegacyUserRespositories> {
    if (this._legacyRepositories) {
      return this._legacyRepositories;
    }
    this._legacyRepositories = await this.aggregateLegacyRepositories();
    return this._legacyRepositories;
  }

  async repositoryPermissions(): Promise<
    IPersonalizedUserAggregateRepositoryPermission[]
  > {
    if (this._repositoryPermissions) {
      return this._repositoryPermissions;
    }
    this._repositoryPermissions = await this.aggregateRepositoryPermissions();
    return this._repositoryPermissions;
  }

  private async aggregateOrganizations(): Promise<IAggregateUserOrganizations> {
    let known: IAggregateUserOrganizations = null;
    if (this._queryCache && this._queryCache.supportsOrganizationMembership) {
      known = (await this.getQueryCacheOrganizations()) as IAggregateUserOrganizations;
    } else {
      known = (await this.getGraphManagerOrganizations()) as IAggregateUserOrganizations;
    }
    known.admin = known.admin.sort(insensitiveSortOrganizations);
    known.member = known.member.sort(insensitiveSortOrganizations);
    // Available organizations
    const all = new Set(this._operations.organizations.values());
    for (const o of known.member) {
      all.delete(o);
    }
    known.available = Array.from(all).sort(insensitiveSortOrganizations);
    return known;
  }

  private async aggregateTeams(): Promise<IAggregateUserTeams> {
    const known = await (this._queryCache &&
    this._queryCache.supportsTeamMembership
      ? this.getQueryCacheTeams()
      : this.getGraphManagerTeams());
    return known;
  }

  private async aggregateLegacyRepositories(): Promise<IAggregateLegacyUserRespositories> {
    if (
      this._queryCache &&
      this._queryCache.supportsRepositoryCollaborators &&
      this._queryCache.supportsTeamPermissions &&
      this._queryCache.supportsTeamMembership
    ) {
      return this.getQueryCacheRepositories();
    } else {
      return this.getGraphManagerRepos();
    }
  }

  private async aggregateRepositoryPermissions(): Promise<
    IPersonalizedUserAggregateRepositoryPermission[]
  > {
    if (
      this._queryCache &&
      this._queryCache.supportsTeamPermissions &&
      this._queryCache.supportsTeamMembership
    ) {
      return this.getQueryCacheRepositoryPermissions();
    } else {
      return this.getProxyRepositoryPermissions();
    }
  }

  async getAggregatedOverview(): Promise<IAggregateUserSummary> {
    let [organizations, teams, repositories] = await Promise.all([
      SettleToStateValue(this.aggregateOrganizations()),
      SettleToStateValue(this.aggregateTeams()),
      SettleToStateValue(this.aggregateLegacyRepositories()),
    ]);
    const results: IAggregateUserSummary = {
      organizations: organizations.value || {
        member: [],
        admin: [],
        available: [],
      },
      teams: teams.value || { maintainer: [], member: [] },
      repos: repositories.value || { byTeam: [] },
    };
    return results;
  }

  async getAggregatedOrganizationOverview(
    organization: Organization
  ): Promise<IAggregateUserSummary> {
    const results = await this.getAggregatedOverview();
    results.teams = this.reduceOrganizationTeams(organization, results.teams);
    // At this time it does not simplify or reduce repo lists or the general orgs list
    return results;
  }

  reduceOrganizationTeams(
    organization: Organization,
    teams: IAggregateUserTeams
  ): IAggregateUserTeams {
    const organizationName = organization.name.toLowerCase();
    return {
      member: teams.member.filter(
        (team) => team.organization.name.toLowerCase() === organizationName
      ),
      maintainer: teams.maintainer.filter(
        (team) => team.organization.name.toLowerCase() === organizationName
      ),
    };
  }

  async getRepoCollaborators(): Promise<any> {
    const operations = this._operations;
    const options = {};
    const repos = await operations.graphManager.getReposWithCollaborators(
      options
    );
    return repos;
  }

  // newer query cache optimized methods

  async getQueryCacheRepositoryPermissions(): Promise<
    IPersonalizedUserAggregateRepositoryPermission[]
  > {
    const userIdString = this.id.toString();
    const queryCache = this._queryCache;
    const personalized: IPersonalizedUserAggregateRepositoryPermission[] = [];
    const repositories = new Map<string, IRepositoryPermissionPair>();
    // Find all the repos that the user have permission to across all configured orgs
    function getOrCreatePair(
      repositoryId: string,
      repository: Repository
    ): IRepositoryPermissionPair {
      let pair = repositories.get(repositoryId);
      let newPair = !!pair;
      if (!pair) {
        pair = {
          repository: repository,
          collaborationPermission: null,
          teamPermissions: [],
        };
        repositories.set(repositoryId, pair);
      }
      if (newPair) {
        // console.log(`${newPair ? 'new' : 'exi'}: ${repositoryId} + ${pair.repository.name} + ${pair.repository.id}`);
      }
      return pair;
    }
    if (queryCache.supportsTeamMembership) {
      const theirTeams = await queryCache.userTeams(userIdString);
      const teamIds = theirTeams.map((team) => team.team.id.toString());
      const teamPermissions = await queryCache.teamsPermissions(teamIds);
      teamPermissions.map((tp) =>
        getOrCreatePair(
          tp.repository.id.toString(),
          tp.repository
        ).teamPermissions.push(tp)
      );
    }
    if (queryCache.supportsRepositoryCollaborators) {
      const theirCollaborationRepositories = await queryCache.userCollaboratorRepositories(
        userIdString
      );
      theirCollaborationRepositories.map(
        (tcr) =>
          (getOrCreatePair(
            tcr.repository.id.toString(),
            tcr.repository
          ).collaborationPermission = tcr.permission)
      );
    }
    // project into the new view
    for (const {
      repository,
      collaborationPermission,
      teamPermissions,
    } of repositories.values()) {
      let bestPermission = null;
      const perms = teamPermissions.map((tp) => {
        const team = tp.team;
        const permission = tp.permission;
        const entity = { ...team.toSimpleJsonObject() };
        entity['permission'] = permission;
        const teamRepositoryPermission = new TeamRepositoryPermission(
          team,
          entity,
          this._operations
        );
        if (isPermissionBetterThan(bestPermission, permission)) {
          bestPermission = permission;
        }
        return teamRepositoryPermission;
      });
      if (
        collaborationPermission &&
        isPermissionBetterThan(bestPermission, collaborationPermission)
      ) {
        bestPermission = collaborationPermission;
      }
      personalized.push({
        repository,
        bestComputedPermission: bestPermission,
        collaboratorPermission: collaborationPermission,
        teamPermissions: perms,
      });
    }
    return personalized;
  }

  async getQueryCacheOrganizations(): Promise<IKnownAggregateUserOrganizations> {
    const userIdString = this.id.toString();
    const membership = await this._queryCache.userOrganizations(userIdString);
    const state: IKnownAggregateUserOrganizations = {
      admin: [],
      member: [],
    };
    for (let { organization, role } of membership) {
      if (
        role !== OrganizationMembershipRole.Admin &&
        role !== OrganizationMembershipRole.Member
      ) {
        throw new Error(
          `Unrecognized or invalid organization ${organization.name} role=${role} for user ${this.id}`
        );
      }
      const bucket =
        OrganizationMembershipRole.Admin === role ? state.admin : state.member;
      bucket.push(organization);
    }
    return state;
  }

  async getQueryCacheTeams(): Promise<IAggregateUserTeams> {
    const maintainer: Team[] = [];
    const member: Team[] = [];
    const userIdString = this.id.toString();
    const teams = await this._queryCache.userTeams(userIdString);
    const awaits: Promise<any>[] = [];
    for (let { role, team } of teams) {
      try {
        if (
          role !== GitHubTeamRole.Maintainer &&
          role !== GitHubTeamRole.Member
        ) {
          throw new Error(
            `Unrecognized or invalid role ${role} for team ID ${team.id} in org ${team.organization.name}`
          );
        }
        const bucket = role === GitHubTeamRole.Maintainer ? maintainer : member;
        await team.getDetails();
        bucket.push(team);
      } catch (getTeamInfoError) {
        // Teams getting deleted is normal and OK.
        if (!ErrorHelper.IsNotFound(getTeamInfoError)) {
          console.log(`Unable to get team information: ${getTeamInfoError}`);
        }
      }
    }
    return { maintainer, member };
  }

  async getQueryCacheRepositories(): Promise<IAggregateLegacyUserRespositories> {
    // Still only returning the legacy version, however...
    const legacyResults: IAggregateLegacyUserRespositories = {
      byTeam: [],
    };

    const userIdString = this.id.toString();
    const teams = await this._queryCache.userTeams(userIdString);
    const teamIdsAsStrings = teams.map((t) => t.team.id.toString());

    const teamPermissions = await this._queryCache.teamsPermissions(
      teamIdsAsStrings
    );
    const reposMap = new Map<number, Repository>();
    teamPermissions.map((tp) => reposMap.set(tp.repository.id, tp.repository));
    legacyResults.byTeam.push(...Array.from(reposMap.values()));

    return legacyResults;
  }

  // legacy graph manager interop methods

  async getGraphManagerRepos(): Promise<IAggregateLegacyUserRespositories> {
    const repos = await this._operations.graphManager.getUserReposByTeamMemberships(
      this.id,
      {}
    );
    return { byTeam: repos.map((personalized) => personalized.repository) };
    // return { byTeam: repos.map(repo => this._operations.getRepositoryWithOrganization(repo.name, repo.organization.login, repo )) };
  }

  async getProxyRepositoryPermissions(): Promise<
    IPersonalizedUserAggregateRepositoryPermission[]
  > {
    return this._operations.graphManager.getUserReposByTeamMemberships(
      this.id,
      {}
    );
  }

  async getGraphManagerTeams(): Promise<IAggregateUserTeams> {
    const maintainer = await this._operations.graphManager.getTeamMemberships(
      this.id,
      GitHubTeamRole.Maintainer
    );
    const member = await this._operations.graphManager.getTeamMemberships(
      this.id
    );
    return {
      maintainer: maintainer.map((team) =>
        this._operations.getTeamByIdWithOrganization(
          team.id,
          team.organization.login,
          team
        )
      ),
      member: member.map((team) =>
        this._operations.getTeamByIdWithOrganization(
          team.id,
          team.organization.login,
          team
        )
      ),
    };
  }

  async getGraphManagerOrganizations(): Promise<IKnownAggregateUserOrganizations> {
    const admin = await this._operations.graphManager.getOrganizationStatusesByName(
      this.id,
      OrganizationMembershipRoleQuery.Admin
    );
    const member = await this._operations.graphManager.getOrganizationStatusesByName(
      this.id
    );
    const state: IKnownAggregateUserOrganizations = {
      admin: admin.map((name) => this._operations.getOrganization(name)),
      member: member.map((name) => this._operations.getOrganization(name)),
    };
    return state;
  }
}

function insensitiveSortOrganizations(a: Organization, b: Organization) {
  return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
}
