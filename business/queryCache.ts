//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Query cache is an optional alternate cache that is used for any non-operational
// purposes, namely customer-facing user interface components like the set of
// orgs the user is a member of, the teams they are a member of, or the respos
// they work with.

// The original implementation of this open source portal worked off of an exclusive
// proxy to the GitHub API, caching heavily in Redis, but after scaling past 10,000
// members and repos, it is no longer scaling appropriately, so this is an attempt
// to address key issues there by leveraging the entity provider setup and other
// jobs.

import {
  IProviders,
  MassagePermissionsToGitHubRepositoryPermission,
} from '../transitional';
import { Organization, OrganizationMembershipRole } from './organization';
import { OrganizationMemberCacheEntity } from '../entities/organizationMemberCache/organizationMemberCache';
import { Operations } from './operations';
import { Team, GitHubTeamRole } from './team';
import { TeamMemberCacheEntity } from '../entities/teamMemberCache/teamMemberCache';
import { Repository, GitHubCollaboratorType } from './repository';
import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';

import Debug from 'debug';
import { TeamCacheEntity } from '../entities/teamCache/teamCache';
import { RepositoryTeamCacheEntity } from '../entities/repositoryTeamCache/repositoryTeamCache';
import { RepositoryCacheEntity } from '../entities/repositoryCache/repositoryCache';
import { RepositoryCollaboratorCacheEntity } from '../entities/repositoryCollaboratorCache/repositoryCollaboratorCache';

const debug = Debug('querycache');

export enum QueryCacheOperation {
  New = 'new',
  Update = 'update',
  Delete = 'delete',
}

export interface IQueryCacheTeamMembership {
  team: Team;
  role: GitHubTeamRole;
  // debug aides:
  cacheEntity: TeamMemberCacheEntity;
  userId: string;
  login: string;
}

export interface IQueryCacheTeam {
  team: Team;
  cacheEntity: TeamCacheEntity;
}

export interface IQueryCacheRepository {
  repository: Repository;
  cacheEntity: RepositoryCacheEntity;
}

export interface IQueryCacheOrganizationMembership {
  organization: Organization;
  role: OrganizationMembershipRole;
  // debug aides:
  // cacheEntity: OrganizationMemberCacheEntity;
  userId: string;
}

export interface IQueryCacheRepositoryCollaborator {
  repository: Repository;
  affiliation: GitHubCollaboratorType;
  permission: GitHubRepositoryPermission;
  // debug aides:
  cacheEntity: RepositoryCollaboratorCacheEntity;
  userId: string;
}

export interface IQueryCacheTeamRepositoryPermission {
  repository: Repository;
  team: Team;
  permission: GitHubRepositoryPermission;
  cacheEntity: RepositoryTeamCacheEntity;
}

export default class QueryCache {
  private _providers: IProviders;

  constructor(providers: IProviders) {
    this._providers = providers;
  }

  get operations(): Operations {
    return this._providers.operations;
  }

  // -- Major removal function for when an organization is deleted or unmanaged

  async removeOrganizationById(organizationId: string): Promise<void> {
    try {
      if (this.supportsOrganizationMembership) {
        await this._providers.organizationMemberCacheProvider.deleteByOrganizationId(
          organizationId
        );
      }
      if (this.supportsRepositories) {
        await this._providers.repositoryCacheProvider.deleteByOrganizationId(
          organizationId
        );
      }
      if (this.supportsRepositoryCollaborators) {
        await this._providers.repositoryCollaboratorCacheProvider.deleteByOrganizationId(
          organizationId
        );
      }
      if (this.supportsTeamPermissions) {
        await this._providers.repositoryTeamCacheProvider.deleteByOrganizationId(
          organizationId
        );
      }
      if (this.supportsTeams) {
        await this._providers.teamCacheProvider.deleteByOrganizationId(
          organizationId
        );
      }
    } catch (groupError) {
      console.dir(groupError);
      throw groupError;
    }
    console.log('removed organization cache for ' + organizationId);
  }

  // -- Team Members

  get supportsTeamMembership(): boolean {
    const teamMemberCacheProvider = this._providers.teamMemberCacheProvider;
    return !!teamMemberCacheProvider;
  }

  async userTeams(githubId: string): Promise<IQueryCacheTeamMembership[]> {
    if (!this.supportsTeamMembership) {
      this.throwMethodNotSupported('userTeams', 'teamMemberCacheProvider');
    }
    const teamMemberCacheProvider = this._providers.teamMemberCacheProvider;
    const rawEntities = await teamMemberCacheProvider.queryTeamMembersByUserId(
      githubId
    );
    return rawEntities
      .map((cacheEntity) => this.hydrateTeamMember(cacheEntity))
      .filter((real) => real);
  }

  async teamMembers(teamId: string): Promise<IQueryCacheTeamMembership[]> {
    if (!this.supportsTeamMembership) {
      this.throwMethodNotSupported('teamMembers', 'teamMemberCacheProvider');
    }
    const teamMemberCacheProvider = this._providers.teamMemberCacheProvider;
    const rawEntities = await teamMemberCacheProvider.queryTeamMembersByTeamId(
      teamId
    );
    return rawEntities
      .map((cacheEntity) => this.hydrateTeamMember(cacheEntity))
      .filter((real) => real);
  }

  private hydrateTeamMember(
    entity: TeamMemberCacheEntity
  ): IQueryCacheTeamMembership {
    try {
      const organization = this.operations.getOrganizationById(
        Number(entity.organizationId)
      );
      const team = organization.team(Number(entity.teamId));
      return {
        team,
        cacheEntity: entity,
        role: entity.teamRole,
        userId: entity.userId,
        login: entity.login,
      };
    } catch (noConfiguredOrganization) {
      console.dir(noConfiguredOrganization);
      return;
    }
  }

  async removeOrganizationTeamMembershipsForUser(
    organizationId: string,
    userId: string
  ): Promise<QueryCacheOperation[]> {
    if (!this.supportsTeamMembership) {
      throw new Error('removeOrganizationTeamMembershipsForUser not supported');
    }
    const operations = [];
    const teamMemberCacheProvider = this._providers.teamMemberCacheProvider;
    const existingEntries = await teamMemberCacheProvider.queryTeamMembersByOrganizationIdAndUserId(
      organizationId,
      userId
    );
    debug(
      `removeOrganizationTeamMembershipsForUser: ${existingEntries.length} team memberships to remove in the organization id=${organizationId} and user id=${userId}`
    );
    for (const existing of existingEntries) {
      try {
        debug(
          `Removing team membership for organization id=${organizationId} user id=${userId}`
        );
        await teamMemberCacheProvider.deleteTeamMemberCache(existing);
        operations.push(QueryCacheOperation.Delete);
      } catch (ignored) {}
    }
    return operations;
  }

  async removeTeamMember(
    organizationId: string,
    teamId: string,
    userId: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsTeamMembership) {
      throw new Error('removeTeamMember not supported');
    }
    const teamMemberCacheProvider = this._providers.teamMemberCacheProvider;
    let cache: TeamMemberCacheEntity = null;
    try {
      cache = await teamMemberCacheProvider.getTeamMemberCacheByUserId(
        organizationId,
        teamId,
        userId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    let outcome: QueryCacheOperation = null;
    if (cache) {
      await teamMemberCacheProvider.deleteTeamMemberCache(cache);
      debug(
        `Removed team member user id=${userId} from team id=${teamId} in organization id=${organizationId}`
      );
      outcome = QueryCacheOperation.Delete;
    }
    return outcome;
  }

  async removeOrganizationTeamMembershipsForTeam(
    organizationId: string,
    teamId: string
  ): Promise<QueryCacheOperation[]> {
    if (!this.supportsTeamMembership) {
      throw new Error('removeOrganizationTeamMembershipsForTeam not supported');
    }
    const operations = [];
    const teamMemberCacheProvider = this._providers.teamMemberCacheProvider;
    const existingEntries = await teamMemberCacheProvider.queryTeamMembersByTeamId(
      teamId
    );
    debug(
      `removeOrganizationTeamMembershipsForTeam: ${existingEntries.length} team memberships to remove in the organization id=${organizationId} and team id=${teamId}`
    );
    for (const existing of existingEntries) {
      try {
        debug(
          `Removing team membership for organization id=${organizationId} team id=${teamId} user id=${existing.userId}`
        );
        await teamMemberCacheProvider.deleteTeamMemberCache(existing);
        operations.push(QueryCacheOperation.Delete);
      } catch (ignored) {}
    }
    return operations;
  }

  async addOrUpdateTeamMember(
    organizationId: string,
    teamId: string,
    userId: string,
    role: GitHubTeamRole,
    login: string,
    avatar: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsTeamMembership) {
      throw new Error('addOrUpdateTeamMember not supported');
    }
    const teamMemberCacheProvider = this._providers.teamMemberCacheProvider;
    let outcome: QueryCacheOperation = null;
    let memberCache: TeamMemberCacheEntity = null;
    try {
      memberCache = await teamMemberCacheProvider.getTeamMemberCacheByUserId(
        organizationId,
        teamId,
        userId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    if (memberCache) {
      const update =
        memberCache.teamRole !== role ||
        memberCache.avatar !== avatar ||
        memberCache.login !== login;
      if (update) {
        memberCache.cacheUpdated = new Date();
        memberCache.teamRole = role;
        memberCache.login = login;
        memberCache.avatar = avatar;
        await teamMemberCacheProvider.updateTeamMemberCache(memberCache);
        debug(
          `Updated team member id=${userId} login=${login} with role=${role} for team=${teamId}`
        );
        outcome = QueryCacheOperation.Update;
      }
    } else {
      memberCache = new TeamMemberCacheEntity();
      memberCache.uniqueId = TeamMemberCacheEntity.GenerateIdentifier(
        organizationId,
        teamId,
        userId
      );
      memberCache.organizationId = organizationId;
      memberCache.userId = userId;
      memberCache.teamId = teamId;
      memberCache.teamRole = role;
      memberCache.login = login;
      memberCache.avatar = avatar;
      await teamMemberCacheProvider.createTeamMemberCache(memberCache);
      debug(
        `Saved team member id=${userId} login=${login} with role=${role} for team=${teamId}`
      );
      outcome = QueryCacheOperation.New;
    }
    return outcome;
  }

  // -- Repositories

  get supportsRepositories(): boolean {
    const repositoryCacheProvider = this._providers.repositoryCacheProvider;
    return !!repositoryCacheProvider;
  }

  repositoryCacheOrganizationIds(): Promise<string[]> {
    if (!this.supportsRepositories) {
      this.throwMethodNotSupported(
        'repositoryCacheOrganizationIds',
        'repositoryCacheProvider'
      );
    }
    return this._providers.repositoryCacheProvider.queryAllOrganizationIds();
  }

  async addOrUpdateRepository(
    organizationId: string,
    repositoryId: string,
    repositoryDetails: any
  ): Promise<QueryCacheOperation> {
    if (!this.supportsRepositories) {
      throw new Error('addOrUpdateRepository not supported');
    }
    const repositoryFieldsToCache = [
      'name',
      'private',
      'description',
      'fork',
      'created_at',
      'updated_at',
      'pushed_at',
      'homepage',
      'size',
      'stargazers_count',
      'watchers_count',
      'language',
      'forks_count',
      'archived',
      'disabled',
      'open_issues_count',
      'license',
      'forks',
      'watchers',
      'network_count',
      'subscribers_count',
    ];
    let outcome: QueryCacheOperation = null;
    const repositoryCacheProvider = this._providers.repositoryCacheProvider;
    const clonedDetails = {};
    repositoryFieldsToCache.forEach(
      (key) => (clonedDetails[key] = repositoryDetails[key])
    );

    let repositoryCache: RepositoryCacheEntity = null;
    try {
      repositoryCache = await repositoryCacheProvider.getRepository(
        repositoryId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    if (repositoryCache) {
      const update =
        !repositoryCache.organizationId ||
        !repositoryCache.repositoryDetails ||
        !repositoryCache.repositoryDetails.updated_at ||
        repositoryCache.repositoryDetails.updated_at !==
          repositoryDetails.updated_at;
      if (update) {
        repositoryCache.cacheUpdated = new Date();
        repositoryCache.organizationId = organizationId;
        repositoryCache.repositoryName = repositoryDetails.name;
        repositoryCache.repositoryDetails = clonedDetails;
        await repositoryCacheProvider.updateRepositoryCache(repositoryCache);
        outcome = QueryCacheOperation.Update;
      }
    } else {
      repositoryCache = new RepositoryCacheEntity();
      repositoryCache.repositoryId = repositoryId;
      repositoryCache.organizationId = organizationId;
      repositoryCache.repositoryName = repositoryDetails.name;
      repositoryCache.repositoryDetails = clonedDetails;
      await repositoryCacheProvider.createRepositoryCache(repositoryCache);
      outcome = QueryCacheOperation.New;
    }
    return outcome;
  }

  async organizationRepositories(
    organizationId: string
  ): Promise<IQueryCacheRepository[]> {
    if (!this.supportsRepositories) {
      this.throwMethodNotSupported(
        'organizationRepositories',
        'repositoryCacheProvider'
      );
    }
    const repositoryCacheProvider = this._providers.repositoryCacheProvider;
    const entities = await repositoryCacheProvider.queryRepositoriesByOrganizationId(
      organizationId
    );
    return entities
      .map((cacheEntity) => this.hydrateRepository(cacheEntity))
      .filter((exists) => exists);
  }

  async allRepositories(): Promise<IQueryCacheRepository[]> {
    if (!this.supportsRepositories) {
      this.throwMethodNotSupported(
        'allRepositories',
        'repositoryCacheProvider'
      );
    }
    const repositoryCacheProvider = this._providers.repositoryCacheProvider;
    const entities = await repositoryCacheProvider.queryAllRepositories();
    return entities
      .map((cacheEntity) => this.hydrateRepository(cacheEntity))
      .filter((exists) => exists);
  }

  private hydrateRepository(
    cacheEntity: RepositoryCacheEntity
  ): IQueryCacheRepository {
    try {
      const operations = this.operations;
      const repository = cacheEntity.hydrateToInstance(operations);
      return {
        repository,
        cacheEntity,
      };
    } catch (noConfiguredOrganization) {
      console.dir(noConfiguredOrganization);
      return;
    }
  }

  async removeRepository(
    organizationId: string,
    repositoryId: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsRepositories) {
      throw new Error('removeRepository not supported');
    }
    const repositoryCacheProvider = this._providers.repositoryCacheProvider;
    let cache: RepositoryCacheEntity = null;
    try {
      cache = await repositoryCacheProvider.getRepository(repositoryId);
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    let outcome: QueryCacheOperation = null;
    if (cache) {
      await repositoryCacheProvider.deleteRepositoryCache(cache);
      debug(
        `Removed organization id=${organizationId} repository id=${repositoryId}`
      );
      outcome = QueryCacheOperation.Delete;
    }
    // Remove all team memberships for this repository ID as well
    if (this.supportsTeamPermissions) {
      await this.removeAllTeamPermissionsForRepository(
        organizationId,
        repositoryId
      );
    }
    // Remove all collaborators from this repository ID
    if (this.supportsRepositoryCollaborators) {
      await this.removeAllCollaboratorsForRepository(repositoryId);
    }
    return outcome;
  }

  // -- Teams

  get supportsTeams(): boolean {
    const teamCacheProvider = this._providers.teamCacheProvider;
    return !!teamCacheProvider;
  }

  async organizationTeams(organizationId: string): Promise<IQueryCacheTeam[]> {
    if (!this.supportsTeams) {
      this.throwMethodNotSupported('organizationTeams', 'teamCacheProvider');
    }
    const teamCacheProvider = this._providers.teamCacheProvider;
    const entities = await teamCacheProvider.queryTeamsByOrganizationId(
      organizationId
    );
    return entities
      .map((cacheEntity) => this.hydrateTeam(cacheEntity))
      .filter((exists) => exists);
  }

  private hydrateTeam(entity: TeamCacheEntity): IQueryCacheTeam {
    try {
      const organization = this.operations.getOrganizationById(
        Number(entity.organizationId)
      );
      const entityBasics = { ...entity.teamDetails };
      entityBasics.id = Number(entity.teamId);
      entityBasics.slug = entity.teamSlug;
      entityBasics.name = entity.teamName;
      entityBasics.description = entity.teamDescription;
      const team = organization.team(Number(entity.teamId), entityBasics);
      return {
        team,
        cacheEntity: entity,
      };
    } catch (noConfiguredOrganization) {
      console.dir(noConfiguredOrganization);
      return;
    }
  }

  async addOrUpdateTeam(
    organizationId: string,
    teamId: string,
    teamDetails: any
  ): Promise<QueryCacheOperation> {
    if (!this.supportsTeams) {
      throw new Error('addTeam not supported');
    }
    let outcome: QueryCacheOperation = null;
    const teamCacheProvider = this._providers.teamCacheProvider;
    const clonedDetails = {
      privacy: teamDetails.privacy,
      created_at: teamDetails.created_at,
      updated_at: teamDetails.updated_at,
      repos_count: teamDetails.repos_count,
      members_count: teamDetails.members_count,
    };

    let teamCache: TeamCacheEntity = null;
    try {
      teamCache = await teamCacheProvider.getTeam(teamId);
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    if (teamCache) {
      const update =
        teamCache.teamName !== teamDetails.name ||
        teamCache.teamSlug !== teamDetails.slug ||
        teamCache.teamDescription !== teamDetails.description ||
        teamCache.teamDetails.updated_at !== teamDetails.updated_at ||
        teamCache.teamDetails.privacy !== teamDetails.privacy ||
        teamCache.teamDetails.created_at !== teamDetails.created_at ||
        teamCache.teamDetails.repos_count !== teamDetails.repos_count ||
        teamCache.teamDetails.members_count !== teamDetails.members_count;
      if (update) {
        teamCache.cacheUpdated = new Date();
        teamCache.teamName = teamDetails.name;
        teamCache.teamSlug = teamDetails.slug;
        teamCache.teamDescription = teamDetails.description;
        teamCache.teamDetails = clonedDetails;
        await teamCacheProvider.updateTeamCache(teamCache);
        debug(`team: updated cache for ${teamCache.teamSlug}`);
        outcome = QueryCacheOperation.Update;
      }
    } else {
      teamCache = new TeamCacheEntity();
      teamCache.teamId = teamId;
      teamCache.organizationId = organizationId;
      teamCache.teamDescription = teamDetails.description;
      teamCache.teamName = teamDetails.name;
      teamCache.teamSlug = teamDetails.slug;
      teamCache.teamDetails = clonedDetails;
      await teamCacheProvider.createTeamCache(teamCache);
      debug(`team: new cache for ${teamCache.teamSlug}`);
      outcome = QueryCacheOperation.New;
    }
    return outcome;
  }

  async removeOrganizationTeam(
    organizationId: string,
    teamId: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsTeams) {
      throw new Error('removeOrganizationTeam not supported');
    }
    const teamCacheProvider = this._providers.teamCacheProvider;
    let cache: TeamCacheEntity = null;
    try {
      cache = await teamCacheProvider.getTeam(teamId);
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    let outcome: QueryCacheOperation = null;
    if (cache) {
      await teamCacheProvider.deleteTeamCache(cache);
      debug(`Removed organization id=${organizationId} team id=${teamId}`);
      outcome = QueryCacheOperation.Delete;
    }
    // Remove all team memberships for this team as well
    if (this.supportsTeamMembership) {
      await this.removeOrganizationTeamMembershipsForTeam(
        organizationId,
        teamId
      );
    }
    if (this.supportsTeamPermissions) {
      await this.removeAllRepositoryPermissionsForTeam(organizationId, teamId);
    }
    return outcome;
  }

  // -- Team permissions

  get supportsTeamPermissions(): boolean {
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    return !!repositoryTeamCacheProvider;
  }

  teamOrganizationIds(): Promise<string[]> {
    if (!this.supportsTeams) {
      this.throwMethodNotSupported('teamOrganizationIds', 'teamCacheProvider');
    }
    return this._providers.teamCacheProvider.queryAllOrganizationIds();
  }

  async addOrUpdateTeamsPermission(
    organizationId: string,
    repositoryId: string,
    repositoryPrivate: boolean,
    repositoryName: string,
    teamId: string,
    permission: GitHubRepositoryPermission
  ): Promise<QueryCacheOperation> {
    if (!this.supportsTeamPermissions) {
      throw new Error('addOrUpdateTeamsPermission not supported');
    }
    let outcome: QueryCacheOperation = null;
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    let cache: RepositoryTeamCacheEntity = null;
    try {
      cache = await repositoryTeamCacheProvider.getRepositoryTeamCacheByTeamId(
        organizationId,
        repositoryId,
        teamId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    if (cache) {
      const update =
        cache.permission !== permission ||
        cache.repositoryName !== repositoryName ||
        cache.repositoryPrivate !== repositoryPrivate;
      if (update) {
        cache.cacheUpdated = new Date();
        cache.repositoryName = repositoryName;
        cache.permission = permission;
        cache.repositoryPrivate = repositoryPrivate ? true : false;
        await repositoryTeamCacheProvider.updateRepositoryTeamCache(cache);
        console.log(
          `Updated repo ${repositoryName} to permission=${permission} team=${teamId}`
        );
        outcome = QueryCacheOperation.Update;
      }
    } else {
      cache = new RepositoryTeamCacheEntity();
      cache.uniqueId = RepositoryTeamCacheEntity.GenerateIdentifier(
        organizationId,
        repositoryId,
        teamId
      );
      cache.organizationId = organizationId;
      cache.repositoryId = repositoryId;
      cache.repositoryName = repositoryName;
      cache.teamId = teamId;
      cache.permission = permission;
      cache.repositoryPrivate = repositoryPrivate ? true : false;
      await repositoryTeamCacheProvider.createRepositoryTeamCache(cache);
      console.log(
        `Saved repo ${repositoryName} permission ${permission} to team ${teamId}`
      );
      outcome = QueryCacheOperation.New;
    }
    return outcome;
  }

  async teamsPermissions(
    teamIds: string[]
  ): Promise<IQueryCacheTeamRepositoryPermission[]> {
    if (!this.supportsTeamPermissions) {
      this.throwMethodNotSupported(
        'teamsPermissions',
        'repositoryTeamCacheProvider'
      );
    }
    if (teamIds.length === 0) {
      return [];
    }
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    const rawEntities = await repositoryTeamCacheProvider.queryByTeamIds(
      teamIds
    );
    return rawEntities
      .map((cacheEntity) => this.hydrateTeamPermission(cacheEntity))
      .filter((exists) => exists);
  }

  async removeRepositoryTeam(
    organizationId: string,
    repositoryId: string,
    teamId: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsTeamPermissions) {
      throw new Error('removeRepositoryTeam not supported');
    }
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    let cache: RepositoryTeamCacheEntity = null;
    try {
      cache = await repositoryTeamCacheProvider.getRepositoryTeamCacheByTeamId(
        organizationId,
        repositoryId,
        teamId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    let outcome: QueryCacheOperation = null;
    if (cache) {
      await repositoryTeamCacheProvider.deleteRepositoryTeamCache(cache);
      debug(
        `Removed permission for repo id=${repositoryId} from team id=${teamId}`
      );
      outcome = QueryCacheOperation.Delete;
    }
    return outcome;
  }

  repositoryTeamOrganizationIds(): Promise<string[]> {
    if (!this.supportsTeamPermissions) {
      this.throwMethodNotSupported(
        'repositoryTeamOrganizationIds',
        'repositoryTeamCacheProvider'
      );
    }
    return this._providers.repositoryTeamCacheProvider.queryAllOrganizationIds();
  }

  async repositoryTeamPermissions(
    repositoryId: string
  ): Promise<IQueryCacheTeamRepositoryPermission[]> {
    if (!this.supportsTeamPermissions) {
      this.throwMethodNotSupported(
        'repositoryTeamPermissions',
        'repositoryTeamCacheProvider'
      );
    }
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    const rawEntities = await repositoryTeamCacheProvider.queryByRepositoryId(
      repositoryId
    );
    return rawEntities
      .map((cacheEntity) => this.hydrateTeamPermission(cacheEntity))
      .filter((exists) => exists);
  }

  async allRepositoryTeamPermissions(): Promise<
    IQueryCacheTeamRepositoryPermission[]
  > {
    if (!this.supportsTeamPermissions) {
      this.throwMethodNotSupported(
        'allRepositoryTeamPermissions',
        'repositoryTeamCacheProvider'
      );
    }
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    const entities = await repositoryTeamCacheProvider.queryAllTeams();
    return entities
      .map((cacheEntity) => this.hydrateTeamPermission(cacheEntity))
      .filter((exists) => exists);
  }

  private hydrateTeamPermission(
    cacheEntity: RepositoryTeamCacheEntity
  ): IQueryCacheTeamRepositoryPermission {
    try {
      const organization = this.operations.getOrganizationById(
        Number(cacheEntity.organizationId)
      );
      const team = organization.team(Number(cacheEntity.teamId));
      const iid = cacheEntity.repositoryId;
      const repository = organization.repository(cacheEntity.repositoryName, {
        id: cacheEntity.repositoryId, // a string version of repositoryId FYI
        private: cacheEntity.repositoryPrivate,
      });
      return {
        repository,
        team,
        permission: cacheEntity.permission,
        cacheEntity,
      };
    } catch (noConfiguredOrganization) {
      console.dir(noConfiguredOrganization);
      return;
    }
  }

  async removeAllRepositoryPermissionsForTeam(
    organizationId: string,
    teamId: string
  ): Promise<QueryCacheOperation[]> {
    if (!this.supportsTeamPermissions) {
      throw new Error('removeRepositoryPermissionsForTeam not supported');
    }
    const operations = [];
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    const existingEntries = await repositoryTeamCacheProvider.queryByTeamId(
      teamId
    );
    debug(
      `removeRepositoryPermissionsForTeam: ${existingEntries.length} repository permissions to remove in the organization id=${organizationId} for team id=${teamId}`
    );
    for (const existing of existingEntries) {
      try {
        debug(
          `Removing repository=${existing.repositoryName} permissions for organization id=${organizationId} team id=${teamId}`
        );
        await repositoryTeamCacheProvider.deleteRepositoryTeamCache(existing);
        operations.push(QueryCacheOperation.Delete);
      } catch (ignored) {}
    }
    return operations;
  }

  async removeAllTeamPermissionsForRepository(
    organizationId: string,
    repositoryId: string
  ): Promise<QueryCacheOperation[]> {
    if (!this.supportsTeamPermissions) {
      throw new Error('removeRepositoryPermissionsForRepository not supported');
    }
    const operations = [];
    const repositoryTeamCacheProvider = this._providers
      .repositoryTeamCacheProvider;
    const existingEntries = await repositoryTeamCacheProvider.queryByRepositoryId(
      repositoryId
    );
    debug(
      `removeRepositoryPermissionsForRepository: ${existingEntries.length} repository permissions to remove in the organization id=${organizationId} for repository id=${repositoryId}`
    );
    for (const existing of existingEntries) {
      try {
        debug(
          `Removing team permissions for organization id=${organizationId} repository id=${repositoryId}`
        );
        await repositoryTeamCacheProvider.deleteRepositoryTeamCache(existing);
        operations.push(QueryCacheOperation.Delete);
      } catch (ignored) {}
    }
    return operations;
  }

  // -- Repo collaboration

  get supportsRepositoryCollaborators(): boolean {
    const repositoryCollaboratorCacheProvider = this._providers
      .repositoryCollaboratorCacheProvider;
    return !!repositoryCollaboratorCacheProvider;
  }

  async allRepositoryCollaborators(): Promise<
    IQueryCacheRepositoryCollaborator[]
  > {
    if (!this.supportsRepositoryCollaborators) {
      throw new Error('allRepositoryCollaborators not supported');
    }
    const repositoryCollaboratorCacheProvider = this._providers
      .repositoryCollaboratorCacheProvider;
    const entities = await repositoryCollaboratorCacheProvider.queryAllCollaborators();
    return entities
      .map((cacheEntity) => this.hydrateRepositoryCollaborator(cacheEntity))
      .filter((real) => real);
  }

  async addOrUpdateCollaborator(
    organizationId: string,
    repositoryId: string,
    repository: Repository,
    repositoryName: string,
    userId: string,
    userLogin: string,
    userAvatar: string,
    permission: GitHubRepositoryPermission,
    collaboratorType: GitHubCollaboratorType
  ): Promise<QueryCacheOperation> {
    if (!this.supportsRepositoryCollaborators) {
      throw new Error('addOrUpdateCollaborator not supported');
    }
    const repositoryCollaboratorCacheProvider = this._providers
      .repositoryCollaboratorCacheProvider;
    let outcome: QueryCacheOperation = null;
    let collaboratorCache: RepositoryCollaboratorCacheEntity = null;
    try {
      collaboratorCache = await repositoryCollaboratorCacheProvider.getRepositoryCollaboratorCacheByUserId(
        organizationId,
        repositoryId,
        userId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    if (collaboratorCache) {
      const update =
        collaboratorCache.avatar !== userAvatar ||
        collaboratorCache.collaboratorType !== collaboratorType ||
        collaboratorCache.login !== userLogin ||
        collaboratorCache.repositoryPrivate !== repository.private ||
        collaboratorCache.repositoryName !== repositoryName ||
        collaboratorCache.permission !== permission;
      if (update) {
        collaboratorCache.cacheUpdated = new Date();
        collaboratorCache.avatar = userAvatar;
        collaboratorCache.repositoryName = repositoryName;
        collaboratorCache.repositoryPrivate = repository.private ? true : false;
        collaboratorCache.collaboratorType = collaboratorType;
        collaboratorCache.login = userLogin;
        collaboratorCache.permission = permission;
        await repositoryCollaboratorCacheProvider.updateRepositoryCollaboratorCache(
          collaboratorCache
        );
        console.log(
          `Updated collaborator login=${userLogin} id=${userId} with type=${collaboratorType}`
        );
      }
    } else {
      collaboratorCache = new RepositoryCollaboratorCacheEntity();
      collaboratorCache.repositoryId = repositoryId;
      collaboratorCache.organizationId = organizationId;
      collaboratorCache.userId = userId;
      collaboratorCache.uniqueId = RepositoryCollaboratorCacheEntity.GenerateIdentifier(
        organizationId,
        repositoryId,
        userId
      );
      collaboratorCache.repositoryName = repositoryName;
      collaboratorCache.repositoryPrivate = repository.private ? true : false;
      collaboratorCache.avatar = userAvatar;
      collaboratorCache.collaboratorType = collaboratorType;
      collaboratorCache.login = userLogin;
      collaboratorCache.permission = permission;
      await repositoryCollaboratorCacheProvider.createRepositoryCollaboratorCache(
        collaboratorCache
      );
      console.log(
        `Saved collaborator login=${userLogin} id=${userId} with type=${collaboratorType}`
      );
    }
    return outcome;
  }

  async userCollaboratorRepositories(
    githubId: string
  ): Promise<IQueryCacheRepositoryCollaborator[]> {
    if (!this.supportsRepositoryCollaborators) {
      this.throwMethodNotSupported(
        'userCollaboratorRepositories',
        'repositoryCollaboratorCacheProvider'
      );
    }
    const repositoryCollaboratorCacheProvider = this._providers
      .repositoryCollaboratorCacheProvider;
    const rawEntities = await repositoryCollaboratorCacheProvider.queryCollaboratorsByUserId(
      githubId
    );
    return rawEntities
      .map((cacheEntity) => this.hydrateRepositoryCollaborator(cacheEntity))
      .filter((real) => real);
  }

  async repositoryCollaborators(
    repositoryId: string
  ): Promise<IQueryCacheRepositoryCollaborator[]> {
    if (!this.supportsRepositoryCollaborators) {
      this.throwMethodNotSupported(
        'repositoryCollaborators',
        'repositoryCollaboratorCacheProvider'
      );
    }
    const repositoryCollaboratorCacheProvider = this._providers
      .repositoryCollaboratorCacheProvider;
    const rawEntities = await repositoryCollaboratorCacheProvider.queryCollaboratorsByRepositoryId(
      repositoryId
    );
    return rawEntities
      .map((cacheEntity) => this.hydrateRepositoryCollaborator(cacheEntity))
      .filter((real) => real);
  }

  repositoryCollaboratorCacheOrganizationIds(): Promise<string[]> {
    if (!this.supportsRepositoryCollaborators) {
      this.throwMethodNotSupported(
        'supportsRepositoryCollaborators',
        'repositoryCollaboratorCacheProvider'
      );
    }
    return this._providers.repositoryCollaboratorCacheProvider.queryAllOrganizationIds();
  }

  async removeAllCollaboratorsForRepository(
    repositoryId: string
  ): Promise<QueryCacheOperation[]> {
    if (!this.supportsRepositoryCollaborators) {
      throw new Error('removeAllCollaboratorsForRepository not supported');
    }
    const operations = [];
    const repositoryCollaboratorCacheProvider = this._providers
      .repositoryCollaboratorCacheProvider;
    await repositoryCollaboratorCacheProvider.deleteByRepositoryId(
      repositoryId
    );
    return operations;
  }

  async removeRepositoryCollaborator(
    organizationId: string,
    repositoryId: string,
    userId: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsRepositoryCollaborators) {
      throw new Error('removeRepositoryCollaborator not supported');
    }
    const repositoryCollaboratorCacheProvider = this._providers
      .repositoryCollaboratorCacheProvider;
    let cache: RepositoryCollaboratorCacheEntity = null;
    try {
      cache = await repositoryCollaboratorCacheProvider.getRepositoryCollaboratorCacheByUserId(
        organizationId,
        repositoryId,
        userId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    let outcome: QueryCacheOperation = null;
    if (cache) {
      await repositoryCollaboratorCacheProvider.deleteRepositoryCollaboratorCache(
        cache
      );
      debug(
        `Removed collaborator from repo id=${repositoryId} collaborator id=${userId}`
      );
      outcome = QueryCacheOperation.Delete;
    }
    return outcome;
  }

  private hydrateRepositoryCollaborator(
    cacheEntity: RepositoryCollaboratorCacheEntity
  ): IQueryCacheRepositoryCollaborator {
    const organization = this.operations.getOrganizationById(
      Number(cacheEntity.organizationId)
    );
    const iid = cacheEntity.repositoryId;
    const repository = organization.repository(cacheEntity.repositoryName, {
      id: cacheEntity.repositoryId,
      private: cacheEntity.repositoryPrivate,
    }); // a string version of repositoryId FYI
    return {
      repository,
      affiliation: cacheEntity.collaboratorType,
      cacheEntity,
      userId: cacheEntity.userId,
      permission: MassagePermissionsToGitHubRepositoryPermission(
        cacheEntity.permission
      ),
    };
  }

  // -- Organization membership

  async addOrUpdateOrganizationMember(
    organizationId: string,
    role: OrganizationMembershipRole,
    userId: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsOrganizationMembership) {
      throw new Error('addOrganizationMember not supported');
    }
    let outcome: QueryCacheOperation = null;
    const organizationMemberCacheProvider = this._providers
      .organizationMemberCacheProvider;
    let cache: OrganizationMemberCacheEntity = null;
    try {
      cache = await organizationMemberCacheProvider.getOrganizationMemberCacheByUserId(
        organizationId,
        userId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    if (cache) {
      const update = cache.role !== role;
      if (update) {
        cache.role = role;
        cache.cacheUpdated = new Date();
        await organizationMemberCacheProvider.updateOrganizationMemberCache(
          cache
        );
        debug(
          `Updated organization id=${organizationId} member id=${userId} to role=${role}`
        );
        outcome = QueryCacheOperation.Update;
      }
    } else {
      cache = new OrganizationMemberCacheEntity();
      cache.organizationId = organizationId;
      cache.userId = userId;
      cache.uniqueId = OrganizationMemberCacheEntity.GenerateIdentifier(
        organizationId,
        userId
      );
      cache.role = role;
      await organizationMemberCacheProvider.createOrganizationMemberCache(
        cache
      );
      debug(
        `Saved organization id=${organizationId} member id=${userId} with role=${role}`
      );
      outcome = QueryCacheOperation.New;
    }
    return outcome;
  }

  async removeOrganizationMember(
    organizationId: string,
    userId: string
  ): Promise<QueryCacheOperation> {
    if (!this.supportsOrganizationMembership) {
      throw new Error('removeOrganizationMember not supported');
    }
    const organizationMemberCacheProvider = this._providers
      .organizationMemberCacheProvider;
    let cache: OrganizationMemberCacheEntity = null;
    try {
      cache = await organizationMemberCacheProvider.getOrganizationMemberCacheByUserId(
        organizationId,
        userId
      );
    } catch (error) {
      if (!error.status || error.status !== 404) {
        throw error;
      }
    }
    let outcome: QueryCacheOperation = null;
    if (cache) {
      await organizationMemberCacheProvider.deleteOrganizationMemberCache(
        cache
      );
      debug(`Removed organization id=${organizationId} member id=${userId}`);
      outcome = QueryCacheOperation.Delete;
    }
    // Repository collaborators become outside collaborators, nothing to cleanup
    // Cleanup any team memberships
    if (this.supportsTeamMembership) {
      await this.removeOrganizationTeamMembershipsForUser(
        organizationId,
        userId
      );
    }
    return outcome;
  }

  get supportsOrganizationMembership(): boolean {
    const organizationMemberCacheProvider = this._providers
      .organizationMemberCacheProvider;
    return !!organizationMemberCacheProvider;
  }

  async organizationMembers(
    organizationId: string
  ): Promise<IQueryCacheOrganizationMembership[]> {
    if (!this.supportsOrganizationMembership) {
      this.throwMethodNotSupported(
        'organizationMembers',
        'organizationMemberCacheProvider'
      );
    }
    const organizationMemberCacheProvider = this._providers
      .organizationMemberCacheProvider;
    const rawEntities = await organizationMemberCacheProvider.queryOrganizationMembersByOrganizationId(
      organizationId
    );
    return this.hydrateOrganizationMembers(rawEntities);
  }

  private hydrateOrganizationMembers(
    rawEntities: OrganizationMemberCacheEntity[]
  ): IQueryCacheOrganizationMembership[] {
    return rawEntities
      .map((cacheEntity) => {
        try {
          return {
            organization: this.operations.getOrganizationById(
              Number(cacheEntity.organizationId)
            ),
            role: cacheEntity.role,
            cacheEntity,
            userId: cacheEntity.userId,
          };
        } catch (noConfiguredOrganization) {
          return;
        }
      })
      .filter((exists) => exists);
  }

  organizationMemberCacheOrganizationIds(): Promise<string[]> {
    if (!this.supportsOrganizationMembership) {
      this.throwMethodNotSupported(
        'organizationMemberCacheOrganizationIds',
        'organizationMemberCacheProvider'
      );
    }
    return this._providers.organizationMemberCacheProvider.queryAllOrganizationIds();
  }

  async userOrganizations(
    githubId: string
  ): Promise<IQueryCacheOrganizationMembership[]> {
    if (!this.supportsOrganizationMembership) {
      this.throwMethodNotSupported(
        'userOrganizations',
        'organizationMemberCacheProvider'
      );
    }
    const organizationMemberCacheProvider = this._providers
      .organizationMemberCacheProvider;
    const rawEntities = await organizationMemberCacheProvider.queryOrganizationMembersByUserId(
      githubId
    );
    return this.hydrateOrganizationMembers(rawEntities);
  }

  private throwMethodNotSupported(name: string, providerName: string) {
    throw new Error(
      `${name} is not supported by QueryCache, the ${providerName} is not configured`
    );
  }
}
