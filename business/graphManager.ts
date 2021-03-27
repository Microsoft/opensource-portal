//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import moment from 'moment';

import { ICorporateLink } from './corporateLink';
import {
  ICacheOptions,
  ILocalCacheOptions,
  IPagedCrossOrganizationCacheOptions,
  isPermissionBetterThan,
} from '../transitional';
import { Operations } from './operations';
import {
  IGetOrganizationMembersOptions,
  OrganizationMembershipRoleQuery,
} from './organization';
import { ICrossOrganizationTeamMembership, GitHubTeamRole } from './team';
import { Repository } from './repository';
import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';
import { TeamRepositoryPermission } from './teamRepositoryPermission';

interface ILocalLinksCache {
  updated: moment.Moment;
  map: Map<string, ICorporateLink>;
}

export interface IPersonalizedUserAggregateRepositoryPermission {
  repository: Repository;
  bestComputedPermission: GitHubRepositoryPermission;

  collaboratorPermission: GitHubRepositoryPermission;
  teamPermissions: TeamRepositoryPermission[];
}

// TODO: rename to CachedCrossOrganizationProxy
export class GraphManager {
  private _operations: Operations;
  private _linksCache: ILocalLinksCache;

  constructor(operations) {
    this._operations = operations;

    return this;
  }

  async getAllOrganizationMember(
    githubId: string | number,
    options?: IGetOrganizationMembersOptions
  ): Promise<any> {
    const allMembers = await this._operations.getMembers(options);
    githubId = typeof githubId === 'string' ? parseInt(githubId, 10) : githubId;
    const member = raiseCrossOrganizationSingleResult(allMembers.get(githubId));
    return member;
  }

  async getOrganizationStatusesByName(
    id: number,
    optionalRole?: OrganizationMembershipRoleQuery
  ): Promise<any> {
    const options: IGetOrganizationMembersOptions = {};
    // options['role'] is not typed, need to validate down the call chain to be clean
    if (optionalRole) {
      options.role = optionalRole;
    }
    const member = await this.getAllOrganizationMember(id, options);
    const value = member && member.orgs ? member.orgs : [];
    return value;
  }

  async getTeamMemberships(
    id: number,
    optionalRole?: GitHubTeamRole
  ): Promise<any> {
    const options: ICrossOrganizationTeamMembership = {};
    if (optionalRole) {
      options.role = optionalRole;
    }
    options.maxAgeSeconds = 60 * 20; // 20m
    options.backgroundRefresh = true;
    const teams = await this.getUserTeams(id, options);
    return teams;
  }

  async getUserTeams(
    githubId: string | number,
    options: ICrossOrganizationTeamMembership
  ): Promise<any[]> {
    const everything = await this.getTeamsWithMembers(options);
    githubId = typeof githubId === 'string' ? parseInt(githubId, 10) : githubId;
    const teams = [];
    for (let i = 0; i < everything.length; i++) {
      const oneTeam = everything[i];
      if (oneTeam && oneTeam.members) {
        for (let j = 0; j < oneTeam.members.length; j++) {
          if (githubId === oneTeam.members[j].id) {
            const teamClone = Object.assign({}, oneTeam);
            oneTeam.organization = {
              login: oneTeam.organization.login,
            };
            delete teamClone.members;
            teams.push(teamClone);
            break;
          }
        }
      }
    }
    return teams;
  }

  getTeamsWithMembers(
    options: ICrossOrganizationTeamMembership
  ): Promise<any[]> {
    options = options || {};
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = 24 * 60 * 60; // One day
    }
    if (options.backgroundRefresh === undefined) {
      options.backgroundRefresh = true;
    }
    options.individualMaxAgeSeconds = 12 * 60 * 60; // Half day
    return this._operations.getTeamsWithMembers(options);
  }

  async getUserReposByTeamMemberships(
    githubId: string | number,
    options: ICacheOptions
  ): Promise<IPersonalizedUserAggregateRepositoryPermission[]> {
    // This is an expensive Redis and refresh user if the app is not deployed
    // with the QueryCache: it reads _all_ the teams for the user by the
    // membership APIs. Do not recommend. Did not scale past 5,000 repos well.
    // This version never takes into account repository direct collaborator
    // permissions, as that is too expensive to compute and coordinate in the
    // traditional proxy cache.
    const everything = await this.getUserTeams(githubId, {}); // TODO:CONFIRM: should this pass options down or not?
    const userTeams = new Set();
    for (let i = 0; i < everything.length; i++) {
      userTeams.add(everything[i].id);
    }
    const allRepos = await this.getReposWithTeams(options);
    const personalizedResults: IPersonalizedUserAggregateRepositoryPermission[] = [];
    for (let i = 0; i < allRepos.length; i++) {
      try {
        const repo = allRepos[i];
        const organizationName = repo.organization.login;
        const organization = this._operations.getOrganization(organizationName);
        if (repo && repo.teams) {
          const userTeamPermissions: TeamRepositoryPermission[] = [];
          let bestPermission = null;
          for (let j = 0; j < repo.teams.length; j++) {
            const t = repo.teams[j]; // technicall a GET /repos/:owner/:repo/teams team permission response
            if (userTeams.has(t.id)) {
              if (repo.private === false && t.permission === 'pull') {
                // Public repos, ignore teams with pull access
              } else {
                const team = organization.team(t.id, t);
                const teamPermission = new TeamRepositoryPermission(
                  team,
                  t,
                  this._operations
                );
                userTeamPermissions.push(teamPermission);
                if (isPermissionBetterThan(bestPermission, t.permission)) {
                  bestPermission = t.permission;
                }
              }
            }
          }
          if (userTeamPermissions.length > 0) {
            const repository = organization.repository(repo.name, repo);
            const personalizedRepositoryPermission: IPersonalizedUserAggregateRepositoryPermission = {
              repository,
              collaboratorPermission: null,
              bestComputedPermission: bestPermission as GitHubRepositoryPermission,
              teamPermissions: userTeamPermissions,
            };
            personalizedResults.push(personalizedRepositoryPermission);
          }
        }
      } catch (individualRepoError) {
        // organization may not be configured for this environment
        console.dir(individualRepoError);
      }
    }
    return personalizedResults;
  }

  getReposWithTeams(
    options?: IPagedCrossOrganizationCacheOptions
  ): Promise<any> {
    options = options || {};
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = 60 * 20 /* 20m per-org collabs list OK */;
    }
    options.individualMaxAgeSeconds = 12 * 60 * 60; // Half day
    if (options.backgroundRefresh === undefined) {
      options.backgroundRefresh = true;
    }
    return this._operations.getRepoTeams(options);
  }

  getReposWithCollaborators(
    options: IPagedCrossOrganizationCacheOptions
  ): Promise<any> {
    options = options || {};
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = 60 * 20 /* 20m per-org collabs list OK */;
    }
    options.individualMaxAgeSeconds = 12 * 60 * 60; // Half day
    if (options.backgroundRefresh === undefined) {
      options.backgroundRefresh = true;
    }
    return this._operations.getRepoCollaborators(options);
  }

  private async getCachedLinksMap(
    maxAgeSecondsLocal: number,
    maxAgeSecondsRemote: number,
    backgroundRefresh: boolean
  ): Promise<Map<string, ICorporateLink>> {
    const operations = this._operations;
    if (!this._linksCache) {
      this._linksCache = {
        updated: null,
        map: new Map(),
      };
    }
    const linksCache = this._linksCache;
    const now = moment();
    const beforeNow = moment().subtract(maxAgeSecondsLocal, 'seconds');
    let isCacheValid =
      linksCache.map &&
      linksCache.updated &&
      beforeNow.isAfter(linksCache.updated);
    if (isCacheValid) {
      return linksCache.map;
    }
    const remoteOptions = {
      backgroundRefresh: backgroundRefresh,
      maxAgeSeconds: maxAgeSecondsRemote,
      // Include all available information
      includeNames: true,
      includeId: true,
      includeServiceAccounts: true,
    };
    const links = await operations.getLinks(remoteOptions);
    const map = new Map();
    for (let i = 0; i < links.length; i++) {
      const link = links[i] as ICorporateLink;
      let id: string | number = link.thirdPartyId;
      if (id) {
        id = parseInt(id, 10);
        map.set(id, links[i]);
      }
    }
    if (
      linksCache.map &&
      linksCache.updated &&
      linksCache.updated.isAfter(now)
    ) {
      // Abandon this update, a newer update has already returned
    } else {
      linksCache.updated = now;
      linksCache.map = map;
    }
    return linksCache.map;
  }
}

function raiseCrossOrganizationSingleResult(result, keyProperty?: string) {
  keyProperty = keyProperty || 'id';
  if (!result || !result[keyProperty] || !result.orgs) {
    return;
  }
  const parentValue = result[keyProperty];
  const clone = Object.assign({}, result);
  clone.orgs = [];
  let copiedFirst = false;
  for (const orgName of Object.getOwnPropertyNames(result.orgs)) {
    const orgResult = result.orgs[orgName];
    if (!orgResult[keyProperty]) {
      throw new Error(
        `The result for the "${orgName}" org does not have a key property, "${keyProperty}".`
      );
    }
    if (orgResult[keyProperty] !== parentValue) {
      throw new Error(
        `The result for the "${orgName}" org key property, "${keyProperty}" does not match the parent key value.`
      );
    }
    if (orgResult.orgs) {
      throw new Error(
        `The result for the "${orgName}" org has a nested 'orgs' property, which is not allowed.`
      );
    }
    if (!copiedFirst) {
      Object.assign(clone, orgResult);
      copiedFirst = true;
    }
    clone.orgs.push(orgName);
  }
  return clone;
}
