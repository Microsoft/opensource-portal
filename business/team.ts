//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import util from 'util';
import _ from 'lodash';

import * as common from './common';

import { wrapError } from '../utils';
import { Organization, OrganizationMembershipState } from './organization';
import {
  ICacheOptions,
  ICacheOptionsPageLimiter,
  IPagedCacheOptions,
  IGetAuthorizationHeader,
  IPurposefulGetAuthorizationHeader,
  IPagedCrossOrganizationCacheOptions,
  ErrorHelper,
  IOperationsInstance,
  throwIfNotCapable,
  IOperationsUrls,
  CoreCapability,
  throwIfNotGitHubCapable,
  validateGitHubLogin,
  IOperationsProviders,
} from '../transitional';
import { TeamMember } from './teamMember';
import { TeamRepositoryPermission } from './teamRepositoryPermission';
import { IApprovalProvider } from '../entities/teamJoinApproval/approvalProvider';
import { TeamJoinApprovalEntity } from '../entities/teamJoinApproval/teamJoinApproval';
import { AppPurpose } from '../github';
import { CacheDefault, getMaxAgeSeconds, getPageSize } from '.';

const teamPrimaryProperties = [
  'id',
  'name',
  'slug',
  'description',
  'members_count',
  'repos_count',
  'created_at',
  'updated_at',
];
const teamSecondaryProperties = [
  'privacy',
  'permission',
  'organization',
  'url',
  'members_url',
  'repositories_url',
];

export interface IGitHubTeamBasics {
  id: number;
  name: string;
  slug: string;
}

export enum GitHubRepositoryType {
  Sources = 'sources',
}

export enum TeamJsonFormat {
  Simple, // basics
  Detailed, // full entity
  Augmented, // entity + corporate configuration layer
}

export interface ICheckRepositoryPermissionOptions extends ICacheOptions {
  organizationName?: string;
}

export interface IGetTeamRepositoriesOptions extends ICacheOptionsPageLimiter {
  type?: GitHubRepositoryType;
}

export interface ITeamMembershipRoleState {
  role?: GitHubTeamRole;
  state?: OrganizationMembershipState;
}

export interface IIsMemberOptions extends ICacheOptions {
  role?: GitHubTeamRole;
}

export interface IGetMembersOptions extends ICacheOptionsPageLimiter {
  role?: GitHubTeamRole;
}

export enum GitHubTeamRole {
  Member = 'member',
  Maintainer = 'maintainer',
}

export interface ICrossOrganizationTeamMembership
  extends IPagedCrossOrganizationCacheOptions {
  role?: GitHubTeamRole;
}

export interface ITeamMembershipOptions {
  role?: GitHubTeamRole;
}

export interface IUpdateTeamMembershipOptions extends ICacheOptions {
  role?: GitHubTeamRole;
}

interface IGetMembersParameters {
  team_slug: string;
  org: string;
  per_page: number;
  role?: string;
  pageLimit?: any;
}

interface IGetRepositoriesParameters {
  org: string;
  team_slug: string;
  per_page: number;
  pageLimit?: any;
}

// TODO: cleanup intentional memory leak
// MEMORY_LEAK: INTENTIONAL: keep a cache going from ID to slug
const memoryIdToSlugStore = new Map<number, string>();

export class Team {
  public static PrimaryProperties = teamPrimaryProperties;

  private _organization: Organization;
  private _operations: IOperationsInstance;
  private _getAuthorizationHeader: IPurposefulGetAuthorizationHeader;

  private _id: number;

  private _slug?: string;
  private _name?: string;

  private _created_at?: any;
  private _updated_at?: any;

  private _description: string;

  private _repos_count: any;
  private _members_count: any;

  private _detailsEntity?: any;
  private _ctorEntity?: any; // temp

  get id(): number {
    return this._id;
  }

  get name(): string {
    return this._name;
  }

  get slug(): string {
    return this._slug;
  }

  get description(): string {
    return this._description;
  }

  get repos_count(): any {
    return this._repos_count;
  }

  get members_count(): any {
    return this._members_count;
  }

  get created_at(): any {
    return this._created_at;
  }

  get updated_at(): any {
    return this._updated_at;
  }

  get organization(): Organization {
    return this._organization;
  }

  constructor(
    organization: Organization,
    entity,
    getAuthorizationHeader: IPurposefulGetAuthorizationHeader,
    operations: IOperationsInstance
  ) {
    if (!entity || !entity.id) {
      throw new Error(
        'Team instantiation requires an incoming entity, or minimum-set entity containing an id property.'
      );
    }
    if (typeof entity.id !== 'number') {
      throw new Error('Team constructor entity.id must be a Number');
    }
    this._organization = organization;
    // TODO: remove assignKnownFieldsPrefixed concept, use newer field definitions instead?
    common.assignKnownFieldsPrefixed(
      this,
      entity,
      'team',
      teamPrimaryProperties,
      teamSecondaryProperties
    );
    this._getAuthorizationHeader = getAuthorizationHeader;
    this._operations = operations;
    this._ctorEntity = entity;
  }

  [util.inspect.custom](depth, options) {
    return `GitHub Team: slug=${this._slug} id=${this.id} org=${this._organization?.name}`;
  }

  asJson(format?: TeamJsonFormat) {
    if (
      format === TeamJsonFormat.Detailed ||
      format === TeamJsonFormat.Augmented
    ) {
      const clone = { ...this._ctorEntity, ...this._detailsEntity };
      // technically will also include `.parent`
      clone.organization = {
        login: clone.organization?.login || this.organization.name,
        id: clone.organization?.id || this.organization.id,
      };
      delete clone.members_url;
      delete clone.repositories_url;
      delete clone.cost;
      delete clone.headers;
      if (format === TeamJsonFormat.Detailed) {
        return clone;
      }
      // Augment with corporate information
      clone.corporateMetadata = {
        isSystemTeam: this.isSystemTeam,
        isBroadAccessTeam: this.isBroadAccessTeam,
      };
      return clone;
    }
    return {
      id: this.id,
      slug: this.slug,
      name: this.name,
      description: this.description,
    };
  }

  get baseUrl() {
    const operations = throwIfNotCapable<IOperationsUrls>(
      this._operations,
      CoreCapability.Urls
    );
    if (this._organization && (this._slug || this._name)) {
      return (
        this._organization.baseUrl + 'teams/' + (this._slug || this._name) + '/'
      );
    }
    return operations.baseUrl + 'teams?q=' + this._id;
  }

  get absoluteBaseUrl(): string {
    return `${this._organization.absoluteBaseUrl}teams/${
      this._slug || this._name
    }/`;
  }

  get nativeUrl() {
    if (this._organization && this._slug) {
      return this._organization.nativeManagementUrl + `teams/${this._slug}/`;
    }
    // Less ideal fallback
    return this._organization.nativeManagementUrl + `teams/`;
  }

  async ensureName(): Promise<void> {
    if (this._name && this._slug) {
      return;
    }
    return await this.getDetails();
  }

  async isDeleted(options?: ICacheOptions): Promise<boolean> {
    try {
      await this.getDetails(options);
    } catch (maybeDeletedError) {
      if (
        maybeDeletedError &&
        maybeDeletedError.status &&
        maybeDeletedError.status === 404
      ) {
        return true;
      }
    }
    return false;
  }

  async getDetails(options?: ICacheOptions): Promise<any> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const cacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgTeamDetailsStaleSeconds,
        options,
        60
      ),
      backgroundRefresh: false,
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    const id = this._id;
    if (!id) {
      throw new Error('team.id required to retrieve team details');
    }
    // If the details already have been loaded, move along without refreshing
    // CONSIDER: Either a time-based cache or ability to override the local cached behavior
    if (this._detailsEntity) {
      return this._detailsEntity;
    }
    const parameters = {
      org_id: this.organization.id,
      team_id: id,
    };
    try {
      const entity = await operations.github.request(
        this.authorize(AppPurpose.Data),
        'GET /organizations/:org_id/team/:team_id',
        parameters,
        cacheOptions
      );
      this._detailsEntity = entity;
      // TODO: move beyond setting with this approach
      common.assignKnownFieldsPrefixed(
        this,
        entity,
        'team',
        teamPrimaryProperties,
        teamSecondaryProperties
      );
      return entity;
    } catch (error) {
      if (error?.status === 403) {
        error = new Error(`Error retrieving team details: ${error}`);
        error.status = 403;
        throw error;
      }
      if (error.status && error.status === 404) {
        error = new Error(`The GitHub team ID ${id} could not be found`);
        error.status = 404;
        throw error;
      }
      throw wrapError(
        error,
        `Could not get details about team ID ${this._id} in the GitHub organization ${this.organization.name}: ${error.message}`
      );
    }
  }

  async getChildTeams(options?: IPagedCacheOptions): Promise<Team[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    if (!this.slug) {
      await this.getDetails();
    }
    const parameters = {
      org: this.organization.name,
      per_page: getPageSize(operations),
      team_slug: this.slug,
    };
    const caching: IPagedCacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgTeamsStaleSeconds,
        options
      ),
      backgroundRefresh: true,
      pageRequestDelay: options.pageRequestDelay || null,
    };
    caching.backgroundRefresh = options.backgroundRefresh;
    const getAuthorizationHeader = this._getAuthorizationHeader.bind(
      this,
      AppPurpose.Data
    ) as IGetAuthorizationHeader;
    const teamEntities = await github.collections.getTeamChildTeams(
      getAuthorizationHeader,
      parameters,
      caching
    );
    const teams = common.createInstances<Team>(
      this,
      this.organization.teamFromEntity,
      teamEntities
    );
    return teams;
  }

  get isBroadAccessTeam(): boolean {
    const teams = this._organization.broadAccessTeams;
    // TODO: validating typing here - number or int?
    if (typeof this._id !== 'number') {
      throw new Error('Team.id must be a number');
    }
    const res = teams.indexOf(this._id);
    return res >= 0;
  }

  get isSystemTeam(): boolean {
    const systemTeams = this._organization.systemTeamIds;
    const res = systemTeams.indexOf(this._id);
    return res >= 0;
  }

  delete(): Promise<void> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      org_id: this.organization.id,
      team_id: this._id,
    };
    // alternate of teams.deleteInOrg
    return github.requestAsPost(
      this.authorize(AppPurpose.Operations),
      'DELETE /organizations/:org_id/team/:team_id',
      parameters
    );
  }

  edit(patch: unknown): Promise<void> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      org_id: this.organization.id,
      team_id: this._id,
    };
    Object.assign({}, patch, parameters);
    // alternate of teams.editInOrg
    return github.requestAsPost(
      this.authorize(AppPurpose.Operations),
      'PATCH /organizations/:org_id/team/:team_id',
      parameters
    );
  }

  removeMembership(username: string): Promise<void> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      org_id: this.organization.id,
      team_id: this._id,
      username: validateGitHubLogin(username),
    };
    return github.requestAsPost(
      this.authorize(AppPurpose.Operations),
      'DELETE /organizations/:org_id/team/:team_id/memberships/:username',
      parameters
    );
  }

  async addMembership(
    username: string,
    options?: IUpdateTeamMembershipOptions
  ): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    options = options || {};
    const role = options.role || GitHubTeamRole.Member;
    if (!this.slug) {
      await this.getDetails();
    }
    const parameters = {
      org: this.organization.name,
      team_slug: this.slug,
      username: validateGitHubLogin(username),
      role,
    };
    const ok = await github.post(
      this.authorize(AppPurpose.CustomerFacing),
      'teams.addOrUpdateMembershipForUserInOrg',
      parameters
    );
    return ok;
  }

  addMaintainer(username: string): Promise<void> {
    return this.addMembership(username, { role: GitHubTeamRole.Maintainer });
  }

  async getMembership(
    username: string,
    options: ICacheOptions
  ): Promise<ITeamMembershipRoleState | boolean> {
    const operations = throwIfNotGitHubCapable(this._operations);
    options = options || {};
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgMembershipDirectStaleSeconds
      );
    }
    // If a background refresh setting is not present, perform a live
    // lookup with this call. This is the opposite of most of the library's
    // general behavior.
    if (options.backgroundRefresh === undefined) {
      options.backgroundRefresh = false;
    }
    const parameters = {
      org_id: this.organization.id,
      team_id: this._id,
      username: validateGitHubLogin(username),
    };
    try {
      const result = await operations.github.request(
        this.authorize(AppPurpose.CustomerFacing),
        'GET /organizations/:org_id/team/:team_id/memberships/:username',
        parameters,
        options
      );
      return result;
    } catch (error) {
      if (error.status == /* loose */ 404) {
        return false;
      }
      let reason = error.message;
      if (error.status) {
        reason += ' ' + error.status;
      }
      const wrappedError = wrapError(
        error,
        `Trouble retrieving the membership for ${username} in team ${this._id}.`
      );
      if (error.status) {
        wrappedError['status'] = error.status;
      }
      throw wrappedError;
    }
  }

  async getMembershipEfficiently(
    username: string,
    options?: IIsMemberOptions
  ): Promise<ITeamMembershipRoleState | boolean> {
    // Hybrid calls are used to check for membership. Since there is
    // often a relatively fresh cache available of all of the members
    // of a team, that data source is used first to avoid a unique
    // GitHub API call.
    const operations = throwIfNotGitHubCapable(this._operations);
    // A background cache is used that is slightly more aggressive
    // than the standard org members list to at least frontload a
    // refresh of the data.
    options = options || {};
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgMembershipStaleSeconds,
        null,
        60
      );
    }
    const isMaintainer = await this.isMaintainer(username, options);
    if (isMaintainer) {
      return {
        role: GitHubTeamRole.Maintainer,
        state: OrganizationMembershipState.Active,
      };
    }
    const isMember = await this.isMember(username);
    if (isMember) {
      return {
        role: GitHubTeamRole.Member,
        state: OrganizationMembershipState.Active,
      };
    }
    // Fallback to the standard membership lookup
    const membershipOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgMembershipDirectStaleSeconds
      ),
    };
    const result = await this.getMembership(username, membershipOptions);
    if (result === false || (result as ITeamMembershipRoleState).role) {
      return false;
    }
    return result;
  }

  async isMaintainer(
    username: string,
    options?: ICacheOptions
  ): Promise<boolean> {
    const isOptions: IIsMemberOptions = Object.assign({}, options);
    isOptions.role = GitHubTeamRole.Maintainer;
    const maintainer = (await this.isMember(
      username,
      isOptions
    )) as GitHubTeamRole;
    return maintainer === GitHubTeamRole.Maintainer ? true : false;
  }

  async isMember(
    username: string,
    options?: IIsMemberOptions
  ): Promise<GitHubTeamRole | boolean> {
    const operations = throwIfNotGitHubCapable(this._operations);
    options = options || {};
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgMembershipStaleSeconds
      );
    }
    const getMembersOptions: IGetMembersOptions = Object.assign({}, options);
    if (!options.role) {
      getMembersOptions.role = GitHubTeamRole.Member;
    }
    const members = await this.getMembers(getMembersOptions);
    const expected = username.toLowerCase();
    for (let i = 0; i < members.length; i++) {
      const member = members[i];
      if (member.login.toLowerCase() === expected) {
        return getMembersOptions.role;
      }
    }
    return false;
  }

  getMaintainers(options?: ICacheOptionsPageLimiter): Promise<TeamMember[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.teamMaintainersStaleSeconds
      );
    }
    const getMemberOptions: IGetMembersOptions = Object.assign(
      {},
      options || {}
    );
    getMemberOptions.role = GitHubTeamRole.Maintainer;
    return this.getMembers(getMemberOptions);
  }

  async getMembers(options?: IGetMembersOptions): Promise<TeamMember[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    if (!this.slug) {
      const cachedSlug = memoryIdToSlugStore.get(Number(this.id));
      if (cachedSlug) {
        this._slug = cachedSlug;
      } else {
        console.log(
          'WARN: team.getMembers had to slowly retrieve a slug to perform the call'
        );
        await this.getDetails(); // octokit rest v17 requires slug or custom endpoint requests
        if (this._slug) {
          memoryIdToSlugStore.set(Number(this.id), this._slug);
        }
      }
    }
    const parameters: IGetMembersParameters = {
      team_slug: this.slug,
      org: this.organization.name,
      per_page: getPageSize(operations),
    };
    const caching: IPagedCacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgMembersStaleSeconds,
        options
      ),
      backgroundRefresh: true,
    };
    if (options && options.backgroundRefresh === false) {
      caching.backgroundRefresh = false;
    }
    if (options.role) {
      parameters.role = options.role;
    }
    if (options.pageLimit) {
      parameters.pageLimit = options.pageLimit;
    }
    // CONSIDER: Check the error object, if present, for error.status == /* loose */ 404 to alert/store telemetry on deleted teams
    try {
      const teamMembersEntities = await github.collections.getTeamMembers(
        this.authorize(AppPurpose.Data),
        parameters,
        caching
      );
      const teamMembers = common.createInstances<TeamMember>(
        this,
        this.memberFromEntity,
        teamMembersEntities
      );
      return teamMembers;
    } catch (error) {
      if (ErrorHelper.IsNotFound(error)) {
        // If a previously cached slug is no longer good, remove from the leaky store
        memoryIdToSlugStore.delete(Number(this.id));
      }
      throw error;
    }
  }

  async getRepositories(
    options?: IGetTeamRepositoriesOptions
  ): Promise<TeamRepositoryPermission[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    // GitHub does not have a concept of filtering this out so we add it
    const customTypeFilteringParameter = options.type;
    if (
      customTypeFilteringParameter &&
      customTypeFilteringParameter !== GitHubRepositoryType.Sources
    ) {
      throw new Error(
        `Custom \'type\' parameter is specified, but at this time only \'sources\' is a valid enum value. Value: ${customTypeFilteringParameter}`
      );
    }
    if (!this.slug) {
      console.log('WARN: had to request team.slug slowly');
      await this.getDetails();
    }
    const parameters: IGetRepositoriesParameters = {
      org: this.organization.name,
      team_slug: this.slug,
      per_page: getPageSize(operations),
    };
    const caching: IPagedCacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgMembersStaleSeconds,
        options
      ),
      backgroundRefresh: true,
    };
    if (options && options.backgroundRefresh === false) {
      caching.backgroundRefresh = false;
    }
    if (options.pageLimit) {
      parameters.pageLimit = options.pageLimit;
    }
    const entities = await github.collections.getTeamRepos(
      this.authorize(AppPurpose.Data),
      parameters,
      caching
    );
    if (customTypeFilteringParameter === 'sources') {
      // Remove forks (non-sources)
      _.remove(entities, (repo: any) => {
        return repo.fork;
      });
    }
    return common.createInstances<TeamRepositoryPermission>(
      this,
      teamRepositoryPermissionsFromEntity,
      entities
    );
  }

  async getOfficialMaintainers(): Promise<TeamMember[]> {
    await this.getDetails();
    const maintainers = await this.getMaintainers();
    if (maintainers.length > 0) {
      return resolveDirectLinks(maintainers);
    }
    const members = await this.organization.sudoersTeam.getMembers();
    return resolveDirectLinks(members);
  }

  member(id, optionalEntity?) {
    let entity = optionalEntity || {};
    if (!optionalEntity) {
      entity.id = id;
    }
    const member = new TeamMember(this, entity, this._operations);
    // CONSIDER: Cache any members in the local instance
    return member;
  }

  memberFromEntity(entity) {
    return this.member(entity.id, entity);
  }

  async getApprovals(): Promise<TeamJoinApprovalEntity[]> {
    const operations = throwIfNotCapable<IOperationsProviders>(
      this._operations,
      CoreCapability.Providers
    );
    const approvalProvider = operations.providers
      .approvalProvider as IApprovalProvider;
    if (!approvalProvider) {
      throw new Error('No approval provider instance available');
    }
    let pendingApprovals: TeamJoinApprovalEntity[] = null;
    try {
      pendingApprovals = await approvalProvider.queryPendingApprovalsForTeam(
        this.id.toString()
      );
    } catch (error) {
      throw wrapError(
        error,
        'We were unable to retrieve the pending approvals list for this team. There may be a data store problem or temporary outage.'
      );
    }
    return pendingApprovals;
  }

  toSimpleJsonObject() {
    return {
      id: typeof this.id === 'number' ? this.id : parseInt(this.id, 10),
      name: this.name,
      slug: this.slug,
      description: this.description,
      repos_count: this.repos_count,
      members_count: this.members_count,
      created_at: this.created_at,
      updated_at: this.updated_at,
    };
  }

  private authorize(purpose: AppPurpose): IGetAuthorizationHeader | string {
    const getAuthorizationHeader = this._getAuthorizationHeader.bind(
      this,
      purpose
    ) as IGetAuthorizationHeader;
    return getAuthorizationHeader;
  }
}

async function resolveDirectLinks(people: TeamMember[]): Promise<TeamMember[]> {
  for (let i = 0; i < people.length; i++) {
    const member = people[i];
    await member.getMailAddress();
  }
  return people;
}

function teamRepositoryPermissionsFromEntity(entity) {
  // private, remapped "this"
  const instance = new TeamRepositoryPermission(this, entity, this._operations);
  return instance;
}
