//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "log"] }] */

import _ from 'lodash';

import {
  IReposError,
  ICacheOptions,
  IPagedCacheOptions,
  IGetAuthorizationHeader,
  IPurposefulGetAuthorizationHeader,
  IReposRestRedisCacheCost,
  IAuthorizationHeaderValue,
  ErrorHelper,
  CreateError,
  NoCacheNoBackground,
  IOperationsUrls,
  IOperationsInstance,
  CoreCapability,
  operationsIsCapable,
  operationsWithCapability,
  throwIfNotCapable,
  throwIfNotGitHubCapable,
  IOperationsLockdownFeatureFlags,
  IOperationsProviders,
  IOperationsLegalEntities,
  IOperationsServiceAccounts,
  IOperationsLinks,
  IOperationsCentralOperationsToken,
  IOperationsTemplates,
} from '../transitional';
import * as common from './common';
import { OrganizationMember } from './organizationMember';
import { Team } from './team';
import { Repository } from './repository';

import { wrapError } from '../utils';
import { StripGitHubEntity } from '../lib/github/restApi';
import { GitHubResponseType } from '../lib/github/endpointEntities';
import { AppPurpose } from '../github';
import {
  OrganizationSetting,
  SpecialTeam,
} from '../entities/organizationSettings/organizationSetting';
import { ICorporateLink } from './corporateLink';
import { createOrganizationSudoInstance, IOrganizationSudo } from '../features';
import {
  CacheDefault,
  DefaultPageSize,
  getMaxAgeSeconds,
  getPageSize,
  OperationsCore,
} from './operations/core';

export interface IAccountBasics {
  id: number;
  login: string;
  avatar_url: string;
  created_at: any;
  updated_at: any;
}

export interface ICreateRepositoryResult {
  response: any;
  repository: Repository;
}

export enum OrganizationMembershipState {
  Active = 'active',
  Pending = 'pending',
}

export enum OrganizationMembershipRole {
  Member = 'member',
  Admin = 'admin',
}

export enum OrganizationMembershipRoleQuery {
  Member = 'member',
  Admin = 'admin',
  All = 'all',
}

export enum OrganizationMembershipTwoFactorFilter {
  AllMembers = 'all',
  TwoFactorOff = '2fa_disabled',
}

export enum GitHubAuditLogInclude {
  Web = 'web',
  Git = 'git',
  All = 'all',
}

export interface GitHubAuditLogEntry {
  '@timestamp': number;
  action: string;
  actor: string;
  created_at: number;
  event: string;
  head_branch?: string;
  head_sha?: string;
  name?: string;
  org: string;
  repo: string;
  started_at: string;
  trigger_id?: string;
  workflow_id?: number;
  workflow_run_id?: number;
  user?: string;
}

export interface GitHubAuditLogFormattedEntryMvp {
  pretty: string;
  raw: GitHubAuditLogEntry;
}

export enum GitHubAuditLogOrder {
  Ascending = 'asc',
  Descending = 'desc',
}

export interface IGetOrganizationAuditLogOptions extends ICacheOptions {
  phrase?: string;
  include?: GitHubAuditLogInclude;
  after?: string;
  before?: string;
  order?: GitHubAuditLogOrder;
  per_page?: number;
}

export interface IGetOrganizationMembersOptions extends IPagedCacheOptions {
  filter?: OrganizationMembershipTwoFactorFilter;
  role?: OrganizationMembershipRoleQuery;
}

export interface IAddOrganizationMembershipOptions extends ICacheOptions {
  role?: OrganizationMembershipRole;
}

export interface IOrganizationMemberPair {
  member?: OrganizationMember;
  link?: ICorporateLink;
}

export interface IOrganizationMembership {
  state: OrganizationMembershipState;
  role: OrganizationMembershipRole;
  organization: any;
  user: any;
}

interface IGetMembersParameters {
  org: string;
  per_page: number;
  filter?: string;
  role?: string;
}

interface ICheckPublicMembershipParameters {
  username: string;
  org: string;
  allowEmptyResponse?: boolean;
}

interface IRedirectError extends IReposError {
  status?: number;
  slug?: string;
  team?: Team;
}

export interface IAdministratorBasics {
  id: number;
  login: string;
  sudo: boolean;
  owner: boolean;
}

export interface IGitHubOrganizationPlanResponse {
  filled_seats: number;
  name: string;
  private_repos: number;
  seats: number;
  space: number;
}

export interface IGitHubOrganizationResponse {
  avatar_url?: string;
  billing_email?: string;
  blog?: string;
  collaborators: number;
  company?: string;
  cost?: IReposRestRedisCacheCost;
  created_at: string;
  default_repository_permission: string;
  description: string;
  disk_usage: number;
  email: string;
  followers: number;
  following: number;
  has_organization_projects: boolean;
  has_repository_projects: boolean;
  headers?: unknown;
  html_url: string;
  id: number;
  is_verified: boolean;
  location: string;
  login: string;
  members_can_create_repositories: boolean;
  name: string;
  node_id: string;
  owned_private_repos: number;
  plan: IGitHubOrganizationPlanResponse;
  private_gists: number;
  public_gists: number;
  total_private_repos: number;
  two_factor_requirement_enabled: boolean;
  type: string;
  updated_at: string;
  url: string;
}

export class Organization {
  private _name: string;
  private _baseUrl: string;
  private _nativeUrl: string;
  private _nativeManagementUrl: string;

  private _operations: IOperationsInstance;
  private _getAuthorizationHeader: IPurposefulGetAuthorizationHeader;
  private _getSpecificAuthorizationHeader: IPurposefulGetAuthorizationHeader;
  private _usesGitHubApp: boolean;
  private _settings: OrganizationSetting;

  private _entity: IGitHubOrganizationResponse;

  private _organizationSudo: IOrganizationSudo;

  id: number;
  uncontrolled: boolean;

  constructor(
    operations: IOperationsInstance,
    name: string,
    settings: OrganizationSetting,
    getAuthorizationHeader: IPurposefulGetAuthorizationHeader,
    getSpecificAuthorizationHeader: IPurposefulGetAuthorizationHeader,
    public hasDynamicSettings: boolean
  ) {
    this._name = settings.organizationName || name;
    this._operations = operations;
    this._settings = settings;
    this._usesGitHubApp = hasDynamicSettings;
    this._getAuthorizationHeader = getAuthorizationHeader;
    this._getSpecificAuthorizationHeader = getSpecificAuthorizationHeader;
    if (settings && settings.organizationId) {
      this.id = Number(settings.organizationId);
    }
    if (operationsIsCapable<IOperationsUrls>(operations, CoreCapability.Urls)) {
      this._baseUrl = `${operations.baseUrl}${this.name}/`;
      this._nativeUrl = `https://github.com/${this.name}/`;
      this._nativeManagementUrl = `https://github.com/orgs/${this.name}/`;
    }
    const withProviders = operations as OperationsCore;
    if (withProviders?.providers) {
      this._organizationSudo = createOrganizationSudoInstance(
        withProviders.providers,
        this
      );
    }
  }

  get baseUrl(): string {
    return this._baseUrl;
  }

  get nativeUrl(): string {
    return this._nativeUrl;
  }

  get nativeManagementUrl(): string {
    return this._nativeManagementUrl;
  }

  get absoluteBaseUrl(): string {
    const operations = throwIfNotCapable<IOperationsUrls>(
      this._operations,
      CoreCapability.Urls
    );
    return operations.absoluteBaseUrl + this.name + '/';
  }

  get name(): string {
    return this._name;
  }

  get active(): boolean {
    return this._settings ? this._settings.active : false;
  }

  get usesApp(): boolean {
    return this._usesGitHubApp;
  }

  asClientJson() {
    // TEMP: TEMP: TEMP: not long-term as currently designed
    return {
      active: this.active,
      createRepositoriesOnGitHub: this.createRepositoriesOnGitHub,
      description: this.description,
      externalMembersPermitted: this.externalMembersPermitted,
      id: this.id,
      locked: this.locked,
      hidden: this.hidden,
      appOnly: this.isAppOnly,
      name: this.name,
      priority: this.priority,
      privateEngineering: this.privateEngineering,
      management: this.getManagementApproach(),
    };
  }

  getManagementApproach() {
    // TEMP: not long-term as designed
    let management = null;
    if (this.hasDynamicSettings) {
      const val = this.getDynamicSettings().getProperty('management') as string;
      management = val;
    }
    return management;
  }

  getEntity() {
    return this._entity;
  }

  async supportsUpdatesApp() {
    try {
      await this._getSpecificAuthorizationHeader(AppPurpose.Updates);
      return true;
    } catch (errror) {
      return false;
    }
  }

  async requireUpdatesApp(functionName: string) {
    const supports = await this.supportsUpdatesApp();
    if (!supports) {
      throw new Error(
        `The ${this.name} organization is not configured to support the necessary Updates app to complete this operation: ${functionName}`
      );
    }
  }

  repository(name: string, optionalEntity?) {
    const entity = Object.assign({}, optionalEntity || {}, {
      name,
    });
    const repository = new Repository(
      this,
      entity,
      this._getAuthorizationHeader,
      this._getSpecificAuthorizationHeader,
      this._operations
    );
    // CONSIDER: Cache any repositories in the local instance
    return repository;
  }

  async getRepositoryById(
    id: number,
    options?: ICacheOptions
  ): Promise<Repository> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    if (!id) {
      throw new Error(
        'Must provide a repository ID to retrieve the repository.'
      );
    }
    const parameters = {
      id,
    };
    const cacheOptions: ICacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.accountDetailStaleSeconds,
        options
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    try {
      const entity = await operations.github.request(
        this.authorize(AppPurpose.Data),
        'GET /repositories/:id',
        parameters,
        cacheOptions
      );
      if (entity.owner.id !== this.id) {
        throw CreateError.NotFound(
          `Repository ID ${parameters.id} has a different owner of ${entity.owner.login} instead of ${this.name}. It has been relocated and will be treated as a 404.`
        );
      }
      return this.repositoryFromEntity(entity);
    } catch (error) {
      if (error.status && error.status === 404) {
        error = new Error(`The GitHub repository ID ${id} could not be found`);
        error.status = 404;
        throw error;
      }
      throw wrapError(
        error,
        `Could not get details about repository ID ${id}: ${error.message}`
      );
    }
  }

  async getRepositories(options?: IPagedCacheOptions): Promise<Repository[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      org: this.name,
      type: 'all',
      per_page: getPageSize(operations),
    };
    const caching = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgReposStaleSeconds,
        options
      ),
      backgroundRefresh: true,
      pageRequestDelay: options.pageRequestDelay || null,
    };
    if (options && options.backgroundRefresh === false) {
      caching.backgroundRefresh = false;
    }
    const repoEntities = await github.collections.getOrgRepos(
      this.authorize(AppPurpose.Data),
      parameters,
      caching
    );
    const repositories = common.createInstances<Repository>(
      this,
      this.repositoryFromEntity,
      repoEntities
    );
    return repositories;
  }

  get priority(): string {
    return this._settings.properties['priority'] || 'secondary';
  }

  get isAppOnly(): boolean {
    return this._settings.hasFeature('appOnly') || false;
  }

  get locked(): boolean {
    return this._settings.hasFeature('locked') || false;
  }

  get hidden(): boolean {
    return this._settings.hasFeature('hidden') || false;
  }

  get pilot_program() {
    return this._settings.properties['1es'];
  }

  get createRepositoriesOnGitHub(): boolean {
    return this._settings.hasFeature('createReposDirect') || false;
  }

  get configuredOrganizationRepositoryTypes(): string {
    return this._settings.properties['type'] || 'public';
  }

  get privateEngineering(): boolean {
    return this._settings.hasFeature('privateEngineering') || false;
  }

  get externalMembersPermitted(): boolean {
    return this._settings.hasFeature('externalMembersPermitted') || false;
  }

  get preventLargeTeamPermissions(): boolean {
    return this._settings.hasFeature('preventLargeTeamPermissions') || false;
  }

  get description(): string {
    return this._settings.portalDescription;
  }

  get webhookSharedSecrets(): string[] {
    const operations = throwIfNotCapable<IOperationsProviders>(
      this._operations,
      CoreCapability.Providers
    );
    const orgSettings = this._settings;
    // Multiple shared can be specified at the organization level to allow for rotation
    // NOTE: hook secrets are no longer moved over...
    let orgSpecificSecrets = orgSettings.properties['hookSecrets'] || [];
    const systemwideConfig = operations.providers.config;
    let systemwideSecrets =
      systemwideConfig.github &&
      systemwideConfig.github.webhooks &&
      systemwideConfig.github.webhooks.sharedSecret
        ? systemwideConfig.github.webhooks.sharedSecret
        : null;
    return _.concat([], orgSpecificSecrets, systemwideSecrets);
  }

  get broadAccessTeams(): number[] {
    return this.getSpecialTeam(SpecialTeam.Everyone, 'everyone membership');
  }

  get invitationTeam(): Team {
    const teams = this.broadAccessTeams;
    if (teams.length > 1) {
      throw new Error('Multiple invitation teams are not supported.');
    }
    return teams.length === 1 ? this.team(teams[0]) : null;
  }

  get systemSudoersTeam(): Team {
    const teams = this.getSpecialTeam(SpecialTeam.GlobalSudo, 'system sudoers');
    if (teams.length > 1) {
      throw new Error('Multiple system sudoer teams are not supported.');
    }
    return teams.length === 1 ? this.team(teams[0]) : null;
  }

  get privateRepositoriesSupported(): boolean {
    return this.getSupportedRepositoryTypesByPriority().includes('private');
  }

  get sudoersTeam(): Team {
    const teams = this.getSpecialTeam(SpecialTeam.Sudo, 'organization sudoers');
    if (teams.length > 1) {
      throw new Error('Multiple sudoer teams are not supported.');
    }
    return teams.length === 1 ? this.team(teams[0]) : null;
  }

  getDynamicSettings(): OrganizationSetting {
    if (!this.hasDynamicSettings) {
      throw new Error(
        'This organization is not configured for dynamic settings'
      );
    }
    return this._settings;
  }

  get specialRepositoryPermissionTeams() {
    return {
      read: this.getSpecialTeam(SpecialTeam.SystemRead, 'read everything'),
      write: this.getSpecialTeam(SpecialTeam.SystemWrite, 'write everything'),
      admin: this.getSpecialTeam(
        SpecialTeam.SystemAdmin,
        'administer everything'
      ),
    };
  }

  getAuthorizationHeader(): IPurposefulGetAuthorizationHeader {
    return this._getAuthorizationHeader;
  }

  async getOrganizationAdministrators(): Promise<IAdministratorBasics[]> {
    // returns an array containing an ID and properties 'owner' and 'sudo' for each
    const administrators = new Map<number, IAdministratorBasics>();
    function getAdministratorEntry(id: number, login: string) {
      let administrator = administrators.get(id);
      if (!administrator) {
        administrator = {
          id,
          login,
          sudo: false,
          owner: false,
        };
        administrators.set(id, administrator);
      }
      return administrator;
    }
    const owners = await this.getOwners();
    for (let i = 0; i < owners.length; i++) {
      const id = owners[i].id;
      const login = owners[i].login;
      getAdministratorEntry(id, login).owner = true;
    }
    const sudoTeam = this.sudoersTeam;
    if (!sudoTeam) {
      return Array.from(administrators.values());
    }
    try {
      const members = await sudoTeam.getMembers();
      for (let i = 0; i < members.length; i++) {
        const id = members[i].id;
        const login = members[i].login;
        getAdministratorEntry(id, login).sudo = true;
      }
      return Array.from(administrators.values());
    } catch (error) {
      if (error && error.status === 404) {
        // The sudo team no longer exists, but we should still have administrator information
        return Array.from(administrators.values());
      }
      throw error;
    }
  }

  get systemTeamIds(): number[] {
    const teamIds = [];
    const sudoTeamInstance = this.sudoersTeam;
    if (sudoTeamInstance) {
      teamIds.push(sudoTeamInstance.id);
    }
    const broadAccessTeams = this.broadAccessTeams;
    if (broadAccessTeams) {
      for (let i = 0; i < broadAccessTeams.length; i++) {
        teamIds.push(broadAccessTeams[i]); // is the actual ID, not the team object
      }
    }
    const specialTeams = this.specialRepositoryPermissionTeams;
    const keys = Object.getOwnPropertyNames(specialTeams);
    keys.forEach((type) => {
      const values = specialTeams[type];
      if (Array.isArray(values)) {
        Array.prototype.push.apply(teamIds, values);
      }
    });
    return teamIds;
  }

  get legalEntities(): string[] {
    const operations = throwIfNotCapable<IOperationsLegalEntities>(
      this._operations,
      CoreCapability.LegalEntities
    );
    const settings = this._settings;
    if (
      settings.legalEntities &&
      Array.isArray(settings.legalEntities) &&
      settings.legalEntities.length > 0
    ) {
      return settings.legalEntities;
    }
    const centralLegalEntities = operations.getDefaultLegalEntities();
    if (centralLegalEntities.length > 0) {
      return centralLegalEntities;
    }
    throw new Error(
      'No legal entities available or defined for the organization, or all organizations through the default value'
    );
  }

  async getRepositoryCreateGitHubToken(): Promise<IAuthorizationHeaderValue> {
    // This method leaks/releases the owner token. In the future a more crisp
    // way of accomplishing this without exposing the token should be created.
    // The function name is specific to the intended use instead of a general-
    // purpose token name.
    const token = await (this.authorize(
      AppPurpose.Operations
    ) as IGetAuthorizationHeader)();
    token.source = 'repository create token';
    return token;
  }

  async createRepository(
    repositoryName: string,
    options
  ): Promise<ICreateRepositoryResult> {
    // TODO: create repository options interface
    const operations = throwIfNotGitHubCapable(this._operations);
    const orgName = this.name;
    delete options.name;
    delete options.org;
    const parameters = Object.assign(
      {
        org: orgName,
        name: repositoryName,
      },
      options
    );
    try {
      const details = await operations.github.post(
        this.authorize(AppPurpose.Operations),
        'repos.createInOrg',
        parameters
      );
      const newRepository = this.repositoryFromEntity(details);
      let response = details;
      try {
        response = StripGitHubEntity(
          GitHubResponseType.Repository,
          details,
          'repos.createInOrg'
        );
      } catch (parseError) {}
      const result: ICreateRepositoryResult = {
        repository: newRepository,
        response,
      };
      return result;
    } catch (error) {
      let contextualError = '';
      if (error.errors && Array.isArray(error.errors)) {
        contextualError =
          error.errors.map((errorEntry) => errorEntry.message).join(', ') +
          '. ';
      }
      const friendlyErrorMessage = `${contextualError}Could not create the repository ${orgName}/${repositoryName}`;
      throw wrapError(error, friendlyErrorMessage);
    }
  }

  async getDetails(): Promise<IGitHubOrganizationResponse> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      org: this.name,
    };
    try {
      const entity = await operations.github.call(
        this.authorize(AppPurpose.Data),
        'orgs.get',
        parameters
      );
      if (entity && entity.id) {
        this.id = entity.id;
      }
      this._entity = entity;
      return entity as IGitHubOrganizationResponse;
    } catch (error) {
      throw wrapError(
        error,
        `Could not get details about the ${this.name} organization: ${error.message}`
      );
    }
  }

  getRepositoryCreateMetadata(options?: any) {
    const operations = throwIfNotCapable<IOperationsProviders>(
      this._operations,
      CoreCapability.Providers
    );
    const settings = this._settings;
    const config = operations.providers.config;
    const metadata = {
      approval: {
        fields: config.github.approvalTypes
          ? config.github.approvalTypes.fields
          : undefined,
      },
      legalEntities: this.legalEntities,
      gitIgnore: {
        default:
          settings.properties['defaultGitIgnoreLanguage'] ||
          config.github.gitignore.default,
        languages: config.github.gitignore.languages,
      },
      templates: this.sanitizedRepositoryCreateTemplates(options || {}),
      visibilities: this.getSupportedRepositoryTypesByPriority(),
    };
    return metadata;
  }

  async getTeamFromSlug(slug: string, options?: ICacheOptions): Promise<Team> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const cacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgTeamDetailsStaleSeconds,
        options
      ),
      backgroundRefresh: false,
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    const parameters = {
      org: this.name,
      team_slug: slug,
    };
    try {
      const entity = await operations.github.call(
        this.authorize(AppPurpose.Data),
        'teams.getByName',
        parameters,
        cacheOptions
      );
      return this.teamFromEntity(entity);
    } catch (error) {
      if (error.status && error.status === 404) {
        error = new Error(
          `The GitHub team with the slug ${slug} could not be found`
        );
        error.status = 404;
        throw error;
      }
      throw error;
    }
  }

  async getTeamFromName(
    nameOrSlug: string,
    options?: ICacheOptions
  ): Promise<Team> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    // Slightly more aggressive attempt to look for the latest team
    // information to help prevent downtime when a new team is created
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgTeamsSlugLookupStaleSeconds
      );
    }
    // Try a direct slug lookup first, for better performance
    try {
      const team = await this.getTeamFromSlug(nameOrSlug);
      if (team) {
        return team;
      }
    } catch (teamSlugLookupError) {
      if (ErrorHelper.IsNotFound(teamSlugLookupError)) {
        // that's OK...
      } else {
        console.log('teamSlugLookupError:');
        console.warn(teamSlugLookupError);
      }
    }
    const expected = nameOrSlug.toLowerCase();
    const teams = await this.getTeams(options);
    let alternativeCandidateById: Team = null;
    for (let i = 0; i < teams.length; i++) {
      const team = teams[i];
      const name = team.name.toLowerCase();
      const slug = team.slug.toLowerCase();
      // Considered a light error condition, this will callback with
      // both a suggestion to redirect to the slug-based name and
      // a legitimate link to the team in the error;
      // TODO: hook up this new change
      if (expected === name && name !== slug) {
        const redirectError: IRedirectError = new Error(
          `The team is also available by slug: ${slug}.`
        );
        redirectError.status = 301;
        redirectError.slug = slug;
        redirectError.team = team;
        throw redirectError;
      }
      if (team.id.toString() == /* loose */ expected) {
        alternativeCandidateById = team;
      }
      if (expected === slug) {
        return team;
      }
    }
    if (alternativeCandidateById) {
      const redirectError: IRedirectError = new Error(
        `The team is also available by slug: ${alternativeCandidateById.slug}.`
      );
      redirectError.status = 301;
      redirectError.slug = alternativeCandidateById.slug;
      redirectError.team = alternativeCandidateById;
      throw alternativeCandidateById;
    }
    const teamNotFoundError: IReposError = new Error(
      'No team was found within the organization matching the provided name'
    );
    teamNotFoundError.status = 404;
    teamNotFoundError.skipLog = true;
    throw teamNotFoundError;
  }

  async getAuthorizedOperationsAccount(): Promise<IAccountBasics> {
    const operations = throwIfNotGitHubCapable(this._operations);
    // LEARN: what happens if this is a bot account?
    try {
      const entity = await operations.github.post(
        this.authorize(AppPurpose.Operations),
        'users.getAuthenticated',
        {}
      );
      return entity as IAccountBasics;
    } catch (error) {
      throw wrapError(
        error,
        'Could not get details about the authenticated account'
      );
    }
  }

  team(id: number, optionalEntity?): Team {
    let entity = optionalEntity || {};
    if (!optionalEntity) {
      entity.id = id;
    }
    const team = new Team(
      this,
      entity,
      this._getAuthorizationHeader.bind(this),
      this._operations
    );
    return team;
  }

  member(id: number, optionalEntity?): OrganizationMember {
    let entity = optionalEntity || {};
    if (!optionalEntity) {
      entity.id = id;
    }
    const member = new OrganizationMember(this, entity, this._operations);
    return member;
  }

  getOwners(
    options?: IPagedCacheOptions
  ): Promise<OrganizationMember[] /* TODO: validate return type */> {
    const memberOptions = Object.assign(
      {},
      options
    ) as IGetOrganizationMembersOptions;
    memberOptions.role = OrganizationMembershipRoleQuery.Admin;
    return this.getMembers(memberOptions);
  }

  async getAuditLog(
    options?: IGetOrganizationAuditLogOptions
  ): Promise<GitHubAuditLogEntry[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const name = this.name;
    const parameters = {
      org: name,
      phrase: options.phrase,
      include: options.include,
      after: options.after,
      before: options.before,
      order: options.order,
      per_page: getPageSize(operations, options),
    };
    const cacheOptions: ICacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.accountDetailStaleSeconds,
        options
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    try {
      const entities = (await operations.github.request(
        this.authorize(AppPurpose.Data),
        'GET /orgs/:org/audit-log',
        parameters,
        cacheOptions
      )) as GitHubAuditLogEntry[];
      // common.assignKnownFieldsPrefixed(this, entity, 'account', primaryAccountProperties, secondaryAccountProperties);
      return entities;
    } catch (error) {
      if (error.status && error.status === 404) {
        error = new Error(
          `The GitHub audit log endpoint is not available for the ${this.name} organization`
        );
        error.status = 404;
        throw error;
      }
      throw wrapError(
        error,
        `Could not get the audit log for the ${this.name} org: ${error.message}`
      );
    }
  }

  async isSudoer(username: string, link: ICorporateLink): Promise<boolean> {
    return await this._organizationSudo.isSudoer(username, link);
  }

  async acceptOrganizationInvitation(
    userToken: string
  ): Promise<IOrganizationMembership> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      org: this.name,
      state: 'active',
    };
    try {
      const response = await operations.github.post(
        `token ${userToken}`,
        'orgs.updateMembershipForAuthenticatedUser',
        parameters
      );
      return response;
    } catch (error) {
      const wrappedError = wrapError(
        error,
        `Could not accept your invitation for the ${this.name} organization on GitHub`
      );
      if (
        error.status === 403 &&
        error.headers &&
        error.headers['x-github-sso']
      ) {
        const xGitHubSso = error.headers['x-github-sso'] as string;
        const i = xGitHubSso.indexOf('url=');
        if (i >= 0) {
          const remainder = xGitHubSso.substr(i + 4);
          console.log(`remaining SSO URL: ${remainder}`);
          wrappedError['x-github-sso-url'] = remainder;
        }
      }
      throw wrappedError;
    }
  }

  async getMembership(
    username: string,
    options?: ICacheOptions
  ): Promise<IOrganizationMembership> {
    options = options || {};
    const orgName = this.name;
    const parameters = {
      username: username,
      org: orgName,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    try {
      const result = await operations.github.call(
        this.authorize(AppPurpose.Operations),
        'orgs.getMembershipForUser',
        parameters
      );
      return result;
    } catch (error) {
      if (error.status == /* loose */ 404) {
        return null;
      }
      let reason = error.message;
      if (error.status) {
        reason += ' ' + error.status;
      }
      const wrappedError = wrapError(
        error,
        `Trouble retrieving the membership for "${username}" in the ${orgName} organization.`
      );
      if (error.status) {
        wrapError['status'] = error.status;
      }
      throw wrappedError;
    }
  }

  async getOperationalMembership(
    username: string
  ): Promise<IOrganizationMembership> {
    if (!username) {
      throw new Error('username must be provided');
    }
    // This is a specific version of the getMembership function that takes
    // no options and never allows for caching [outside of the standard
    // e-tag validation with the real-time GitHub API]
    return await this.getMembership(username, NoCacheNoBackground);
  }

  async addMembership(
    username: string,
    options?: IAddOrganizationMembershipOptions
  ): Promise<IOrganizationMembership> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    options = options || {};
    const role = options.role || 'member';
    const parameters = {
      org: this.name,
      username: username,
      role: role,
    };
    const ok = await github.post(
      this.authorize(AppPurpose.Operations),
      'orgs.setMembershipForUser',
      parameters
    );
    return ok as IOrganizationMembership; // state: pending or acive, role: admin or member
  }

  async checkPublicMembership(
    username: string,
    options?: ICacheOptions
  ): Promise<boolean> {
    // NOTE: This method is unable to be cached by the underlying
    // system since there is no etag returned for status code-only
    // results.
    options = options || {};
    const parameters: ICheckPublicMembershipParameters = {
      username: username,
      org: this.name,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    parameters.allowEmptyResponse = true;
    try {
      await operations.github.post(
        this.authorize(AppPurpose.CustomerFacing),
        'orgs.checkPublicMembershipForUser',
        parameters
      );
      return true;
    } catch (error) {
      // The user either is not a member of the organization, or their membership is concealed
      if (error && error.status == /* loose */ 404) {
        return false;
      }
      throw wrapError(
        error,
        `Trouble retrieving the public membership status for ${username} in the ${this.name} organization: ${error.message}`
      );
    }
  }

  async concealMembership(login: string, userToken: string): Promise<void> {
    // This call required a provider user token with the expanded write:org scope
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      org: this.name,
      username: login,
    };
    try {
      const ok = await operations.github.post(
        `token ${userToken}`,
        'orgs.removePublicMembershipForAuthenticatedUser',
        parameters
      );
    } catch (error) {
      throw wrapError(
        error,
        `Could not conceal the ${this.name} organization membership for  ${login}: ${error.message}`
      );
    }
  }

  async publicizeMembership(login: string, userToken: string): Promise<void> {
    // This call required a provider user token with the expanded write:org scope
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      org: this.name,
      username: login,
    };
    try {
      await operations.github.post(
        `token ${userToken}`,
        'orgs.setPublicMembershipForAuthenticatedUser',
        parameters
      );
    } catch (error) {
      throw wrapError(
        error,
        `Could not publicize the ${this.name} organization membership for  ${login}: ${error.message}`
      );
    }
  }

  async getMembers(
    options?: IGetOrganizationMembersOptions
  ): Promise<OrganizationMember[] /*todo: validate*/> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const getAuthorizationHeader = this._getAuthorizationHeader.bind(
      this,
      AppPurpose.Data
    ) as IGetAuthorizationHeader;
    const github = operations.github;
    const parameters: IGetMembersParameters = {
      org: this.name,
      per_page: getPageSize(operations),
    };
    if (options.filter) {
      parameters.filter = options.filter;
    }
    if (options.role) {
      parameters.role = options.role;
    }
    const caching = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgMembersStaleSeconds,
        options
      ),
      backgroundRefresh: true,
      pageRequestDelay: options.pageRequestDelay,
    };
    if (options && options.backgroundRefresh === false) {
      caching.backgroundRefresh = false;
    }
    const memberEntities = await github.collections.getOrgMembers(
      getAuthorizationHeader,
      parameters,
      caching
    );
    const members = common.createInstances<OrganizationMember>(
      this,
      this.memberFromEntity,
      memberEntities
    );
    return members;
  }

  getMembersWithoutTwoFactor(options?: IPagedCacheOptions): Promise<any> {
    const clonedOptions: IGetOrganizationMembersOptions = Object.assign(
      {},
      options || {}
    );
    clonedOptions.filter = OrganizationMembershipTwoFactorFilter.TwoFactorOff;
    return this.getMembers(clonedOptions);
  }

  async isMemberSingleFactor(
    username: string,
    options?: IPagedCacheOptions
  ): Promise<boolean> {
    const membersWithoutTwoFactor = await this.getMembersWithoutTwoFactor(
      options
    );
    const lowerCase = username.toLowerCase();
    for (let i = 0; i < membersWithoutTwoFactor.length; i++) {
      const lc = membersWithoutTwoFactor[i].login.toLowerCase();
      if (lowerCase === lc) {
        return true;
      }
    }
    return false;
  }

  private async getMemberPairs(
    options?: IGetOrganizationMembersOptions
  ): Promise<IOrganizationMemberPair[]> {
    const members = await this.getMembers(options);
    const operations = throwIfNotCapable<IOperationsLinks>(
      this._operations,
      CoreCapability.Links
    );
    const linksArray = await operations.getLinks();
    const links = new Map<string, ICorporateLink>();
    for (const link of linksArray) {
      links.set(link.thirdPartyUsername.toLowerCase(), link);
    }
    return members.map((member) => {
      return {
        member,
        link: links.get(member.login.toLowerCase()),
      };
    });
  }

  async getServiceAccounts(
    excludeSystemAccounts: boolean,
    options?: IGetOrganizationMembersOptions
  ): Promise<IOrganizationMemberPair[]> {
    const operations = operationsWithCapability<IOperationsServiceAccounts>(
      this._operations,
      CoreCapability.ServiceAccounts
    );
    if (operations) {
      const pairs = await this.getMemberPairs(options);
      let accounts = pairs.filter(
        (pair) => pair.link && pair.link.isServiceAccount
      );
      if (excludeSystemAccounts) {
        accounts = accounts.filter(
          (pair) => !operations.isSystemAccountByUsername(pair.member.login)
        );
      }
      return accounts;
    }
    return [];
  }

  async getLinkedMembers(
    options?: IGetOrganizationMembersOptions
  ): Promise<IOrganizationMemberPair[]> {
    const pairs = await this.getMemberPairs(options);
    return pairs.filter((pair) => pair.link);
  }

  async getUnlinkedMembers(
    options?: IGetOrganizationMembersOptions
  ): Promise<OrganizationMember[]> {
    const pairs = await this.getMemberPairs(options);
    return pairs.filter((pair) => !pair.link).map((entry) => entry.member);
  }

  async getTeams(options?: IPagedCacheOptions): Promise<Team[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      org: this.name,
      per_page: getPageSize(operations),
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
    const teamEntities = await github.collections.getOrgTeams(
      getAuthorizationHeader,
      parameters,
      caching
    );
    const teams = common.createInstances<Team>(
      this,
      this.teamFromEntity,
      teamEntities
    );
    return teams;
  }

  async removeMember(login: string, optionalId?: string): Promise<void> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const opsWithProviders = operationsWithCapability<IOperationsProviders>(
      operations,
      CoreCapability.Providers
    );
    const queryCache = opsWithProviders?.providers.queryCache;
    const parameters = {
      org: this.name,
      username: login,
    };
    try {
      await operations.github.post(
        this.authorize(AppPurpose.Operations),
        'orgs.removeMembershipForUser',
        parameters
      );
      if (queryCache?.supportsOrganizationMembership) {
        try {
          if (!optionalId) {
            const centralOps = operationsWithCapability<IOperationsCentralOperationsToken>(
              operations,
              CoreCapability.GitHubCentralOperations
            );
            if (centralOps) {
              const account = await centralOps.getAccountByUsername(login);
              optionalId = account.id.toString();
            }
          }
          await queryCache.removeOrganizationMember(
            this.id.toString(),
            optionalId
          );
        } catch (ignored) {}
      }
    } catch (error) {
      throw wrapError(
        error,
        `Could not remove the organization member ${login}`
      );
    }
  }

  async getMembershipInvitations(): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      org: this.name,
    };
    try {
      const invitations = await operations.github.call(
        this.authorize(AppPurpose.Operations),
        'orgs.listPendingInvitations',
        parameters
      );
      return invitations;
    } catch (error) {
      if (error.status == /* loose */ 404) {
        return null;
      }
      throw error;
    }
  }

  private authorize(purpose: AppPurpose): IGetAuthorizationHeader {
    const getAuthorizationHeader = this._getAuthorizationHeader.bind(
      this,
      purpose
    ) as IGetAuthorizationHeader;
    return getAuthorizationHeader;
  }

  private sanitizedRepositoryCreateTemplates(options) {
    return this.repositoryCreateTemplates(options).map((template) => {
      return {
        id: template.id,
        spdx: template.spdx,
        name: template.name,
        environments: template.environments,
        legalEntities: template.legalEntities,
      };
    });
  }

  private repositoryCreateTemplates(options) {
    const operations = throwIfNotCapable<IOperationsProviders>(
      this._operations,
      CoreCapability.Providers
    );
    const opsTemplates = throwIfNotCapable<IOperationsTemplates>(
      operations,
      CoreCapability.Templates
    );
    options = options || {};
    const projectType = options.projectType;
    // projectType option:
    // if any only if present in the request AND there is a 'forceForReleaseType'
    // value set on at least one template, return only the set of 'forced'
    // templates. the scenario enabled here is to allow sample code to always
    // force one of the official sample code templates and not fallback to
    // standard templates.
    const config = operations.providers.config;
    const templates = [];
    const configuredTemplateRoot = config.github.templates || {};
    const configuredTemplateDefinitions =
      configuredTemplateRoot && configuredTemplateRoot.definitions
        ? configuredTemplateRoot.definitions
        : {};
    const templateDefinitions = configuredTemplateDefinitions || {};
    const allTemplateNames = Object.getOwnPropertyNames(templateDefinitions);
    const fallbackTemplates =
      opsTemplates.getDefaultRepositoryTemplateNames() || allTemplateNames;
    const ts =
      this._settings.templates && this._settings.templates.length > 0
        ? this._settings.templates
        : fallbackTemplates;
    const legalEntities = this.legalEntities;
    const limitedTypeTemplates = [];
    ts.forEach((templateId) => {
      const td = templateDefinitions[templateId];
      const candidateTemplate = Object.assign({ id: templateId }, td);
      let template = null;
      if (candidateTemplate.legalEntity) {
        for (let i = 0; i < legalEntities.length && !template; i++) {
          if (
            legalEntities[i].toLowerCase() ===
            candidateTemplate.legalEntity.toLowerCase()
          ) {
            template = candidateTemplate;
            template.legalEntities = [template.legalEntity];
            delete template.legalEntity;
          }
        }
      } else {
        candidateTemplate.legalEntities = legalEntities;
        template = candidateTemplate;
      }
      if (template && template.name) {
        templates.push(template);
        if (
          projectType &&
          template.forceForReleaseType &&
          template.forceForReleaseType == projectType
        ) {
          limitedTypeTemplates.push(template);
        }
      }
    });
    if (projectType && limitedTypeTemplates.length) {
      return limitedTypeTemplates;
    }
    return templates;
  }

  // Specialized features, opt-in only

  isNewRepositoryLockdownSystemEnabled() {
    const operations = operationsWithCapability<IOperationsLockdownFeatureFlags>(
      this._operations,
      CoreCapability.LockdownFeatureFlags
    );
    if (operations) {
      return (
        operations.allowUnauthorizedNewRepositoryLockdownSystemFeature() &&
        this._settings.hasFeature('new-repository-lockdown-system')
      );
    }
    return false;
  }

  isForkLockdownSystemEnabled() {
    const operations = operationsWithCapability<IOperationsLockdownFeatureFlags>(
      this._operations,
      CoreCapability.LockdownFeatureFlags
    );
    return (
      operations.allowUnauthorizedForkLockdownSystemFeature() &&
      this._settings.hasFeature('lock-new-forks')
    );
  }

  isTransferLockdownSystemEnabled() {
    const operations = operationsWithCapability<IOperationsLockdownFeatureFlags>(
      this._operations,
      CoreCapability.LockdownFeatureFlags
    );
    if (operations) {
      return (
        operations.allowTransferLockdownSystemFeature() &&
        this._settings.hasFeature('lock-transfers')
      );
    }
    return false;
  }

  // Helper functions

  memberFromEntity(entity): OrganizationMember {
    return this.member(entity.id, entity);
  }

  teamFromEntity(entity): Team {
    return this.team(entity.id, entity);
  }

  repositoryFromEntity(entity): Repository {
    return this.repository(entity.name, entity);
  }

  getLegacySystemObjects() {
    const settings = this._settings;
    const operations = this._operations;
    return { settings, operations };
  }

  private getSpecialTeam(
    specialTeam: SpecialTeam,
    friendlyName: string,
    throwIfMissing?: boolean
  ): number[] {
    let teamId: number = null;
    for (const entry of this._settings.specialTeams) {
      if (entry.specialTeam === specialTeam) {
        teamId = entry.teamId;
        break;
      }
    }
    if (throwIfMissing) {
      throw new Error(
        `Missing configured organization "${this.name}" special team ${specialTeam} - ${friendlyName}`
      );
    }
    const teams: number[] = [];
    if (teamId) {
      teams.push(teamId);
    }
    return teams;
  }

  private getSupportedRepositoryTypesByPriority() {
    // Returns the types of repositories supported by the configuration for the organization.
    // The returned array position 0 represents the recommended default choice for new repos.
    // Note that while the configuration may say 'private', the organization may not have
    // a billing relationship, so repo create APIs would fail asking you to upgrade to a paid
    // plan.
    const settings = this._settings;
    const type = settings.properties['type'] || 'public';
    let types = ['public'];
    switch (type) {
      case 'public':
        break;
      case 'publicprivate':
        types.push('private');
        break;
      case 'private':
        types.splice(0, 1, 'private');
        break;
      case 'privatepublic':
        types.splice(0, 0, 'private');
        break;
      default:
        throw new Error(
          `Unsupported configuration for repository types in the organization: ${type}`
        );
    }
    return types;
  }
}
