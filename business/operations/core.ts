//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { getCentralOperationsAuthorizationHeader } from '..';

import { OrganizationSetting } from '../../entities/organizationSettings/organizationSetting';
import { GitHubAppAuthenticationType, AppPurpose } from '../../github';
import { GitHubTokenManager } from '../../github/tokenManager';
import { RestLibrary } from '../../lib/github';
import {
  CoreCapability,
  CreateError,
  IAuthorizationHeaderValue,
  ICacheDefaultTimes,
  ICacheOptions,
  IOperationsCentralOperationsToken,
  IOperationsDefaultCacheTimes,
  IOperationsGitHubRestLibrary,
  IOperationsInstance,
  IOperationsProviders,
  IOperationsUrls,
  IProviders,
  throwIfNotCapable,
  throwIfNotGitHubCapable,
} from '../../transitional';
import { wrapError } from '../../utils';
import { Account } from '../account';
import GitHubApplication from '../application';

export interface IOperationsCoreOptions {
  github: RestLibrary;
  providers: IProviders;
}

export enum CacheDefault {
  orgReposStaleSeconds = 'orgReposStaleSeconds',
  orgRepoTeamsStaleSeconds = 'orgRepoTeamsStaleSeconds',
  orgRepoCollaboratorsStaleSeconds = 'orgRepoCollaboratorsStaleSeconds',
  orgRepoCollaboratorStaleSeconds = 'orgRepoCollaboratorStaleSeconds',
  orgRepoDetailsStaleSeconds = 'orgRepoDetailsStaleSeconds',
  orgTeamsStaleSeconds = 'orgTeamsStaleSeconds',
  orgTeamDetailsStaleSeconds = 'orgTeamDetailsStaleSeconds',
  orgTeamsSlugLookupStaleSeconds = 'orgTeamsSlugLookupStaleSeconds',
  orgMembersStaleSeconds = 'orgMembersStaleSeconds',
  teamMaintainersStaleSeconds = 'teamMaintainersStaleSeconds',
  orgMembershipStaleSeconds = 'orgMembershipStaleSeconds',
  orgMembershipDirectStaleSeconds = 'orgMembershipDirectStaleSeconds',
  crossOrgsReposStaleSecondsPerOrg = 'crossOrgsReposStaleSecondsPerOrg',
  crossOrgsReposParallelCalls = 'crossOrgsReposParallelCalls',
  crossOrgsMembersStaleSecondsPerOrg = 'crossOrgsMembersStaleSecondsPerOrg',
  crossOrgsMembersParallelCalls = 'crossOrgsMembersParallelCalls',
  corporateLinksStaleSeconds = 'corporateLinksStaleSeconds',
  repoBranchesStaleSeconds = 'repoBranchesStaleSeconds',
  repoPullsStaleSeconds = 'repoPullsStaleSeconds',
  accountDetailStaleSeconds = 'accountDetailStaleSeconds',
  teamDetailStaleSeconds = 'teamDetailStaleSeconds',
  orgRepoWebhooksStaleSeconds = 'orgRepoWebhooksStaleSeconds',
  teamRepositoryPermissionStaleSeconds = 'teamRepositoryPermissionStaleSeconds',
}

// defaults could move to configuration alternatively
const defaults: ICacheDefaultTimes = {
  [CacheDefault.orgReposStaleSeconds]: 60 * 15 /* 15m */,
  [CacheDefault.orgRepoTeamsStaleSeconds]: 60 * 3 /* 3m */,
  [CacheDefault.orgRepoCollaboratorsStaleSeconds]: 60 * 30 /* 30m */,
  [CacheDefault.orgRepoCollaboratorStaleSeconds]: 30 /* half minute */,
  [CacheDefault.orgRepoDetailsStaleSeconds]: 60 * 5 /* 5m */,
  [CacheDefault.orgTeamsStaleSeconds]: 60 * 5 /* 5m */,
  [CacheDefault.orgTeamDetailsStaleSeconds]: 60 * 30 /* 30m */,
  [CacheDefault.orgTeamsSlugLookupStaleSeconds]: 30 /* half a minute */,
  [CacheDefault.orgMembersStaleSeconds]: 60 * 30 /* 30m */,
  [CacheDefault.teamMaintainersStaleSeconds]: 60 * 2 /* 2m */,
  [CacheDefault.orgMembershipStaleSeconds]: 60 * 5 /* 5m */,
  [CacheDefault.orgMembershipDirectStaleSeconds]: 30 /* 30s */,
  [CacheDefault.crossOrgsReposStaleSecondsPerOrg]:
    60 * 60 * 2 /* 2 hours per org */,
  [CacheDefault.crossOrgsReposParallelCalls]: 3,
  [CacheDefault.crossOrgsMembersStaleSecondsPerOrg]:
    60 * 60 * 2 /* 2 hours per org */,
  [CacheDefault.crossOrgsMembersParallelCalls]: 5,
  [CacheDefault.corporateLinksStaleSeconds]: 30 /* 30s (used to be 5m) */,
  [CacheDefault.repoBranchesStaleSeconds]: 60 * 5 /* 5m */,
  [CacheDefault.repoPullsStaleSeconds]: 60 * 60 /* 60m */,
  [CacheDefault.accountDetailStaleSeconds]: 60 * 60 * 24 /* 24h */,
  [CacheDefault.teamDetailStaleSeconds]: 60 * 60 * 2 /* 2h */,
  [CacheDefault.orgRepoWebhooksStaleSeconds]: 60 * 60 * 8 /* 8h */,
  [CacheDefault.teamRepositoryPermissionStaleSeconds]: 0 /* 0m */,
};

export const DefaultPageSize = 100;

export interface IOptionWithPageSize {
  per_page?: number;
}

export function getPageSize(
  operations: IOperationsInstance,
  options?: IOptionWithPageSize
) {
  if (options?.per_page) {
    return options.per_page;
  }
  if (operations && (operations as any).defaultPageSize) {
    return (operations as any).defaultPageSize as number;
  }
  return DefaultPageSize;
}

export function getMaxAgeSeconds(
  operations: IOperationsInstance,
  cacheDefault: CacheDefault,
  options?: ICacheOptions,
  fallback?: number
) {
  if (options && options.maxAgeSeconds !== undefined) {
    return options.maxAgeSeconds as number;
  }
  if (operations.hasCapability(CoreCapability.DefaultCacheTimes)) {
    const ops = (operations as any) as IOperationsDefaultCacheTimes;
    if (ops.defaults && ops.defaults[cacheDefault] !== undefined) {
      return ops.defaults[cacheDefault] as number;
    }
  }
  return fallback || undefined;
}

export abstract class OperationsCore
  implements
    IOperationsGitHubRestLibrary,
    IOperationsUrls,
    IOperationsDefaultCacheTimes,
    IOperationsProviders,
    IOperationsInstance {
  private _github: RestLibrary;
  private _defaults: ICacheDefaultTimes;
  private _applicationIds: Map<number, GitHubApplication>;
  private _initialized: Date;
  protected _baseUrl: string;
  private _providers: IProviders;

  private _capabilities: Set<CoreCapability>;

  constructor(options: IOperationsCoreOptions) {
    this._defaults = Object.assign({}, defaults);
    const providers = options.providers;
    this._github = options.github || providers.github;
    this._applicationIds = new Map();
    this._baseUrl = '/ae';
    this._providers = providers;
    this._capabilities = new Set();

    this.addCapability(CoreCapability.GitHubRestApi);
    this.addCapability(CoreCapability.DefaultCacheTimes);
    this.addCapability(CoreCapability.Urls);
    this.addCapability(CoreCapability.Providers);
  }

  protected addCapability(capability: CoreCapability) {
    if (capability) {
      this._capabilities.add(capability);
    }
  }

  hasCapability(capability: CoreCapability): boolean {
    return this._capabilities.has(capability);
  }

  throwIfNotCompatible(capability: CoreCapability) {
    if (!this.hasCapability(capability)) {
      throw new Error(
        `The operations implementation is not capable of supporting ${capability}`
      );
    }
  }

  protected abstract get tokenManager(): GitHubTokenManager;

  async getAccountByUsername(
    username: string,
    options?: ICacheOptions
  ): Promise<Account> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this);
    const centralOperations = throwIfNotCapable<IOperationsCentralOperationsToken>(
      this,
      CoreCapability.GitHubCentralOperations
    );
    if (!username) {
      throw CreateError.ParameterRequired('username');
    }
    const parameters = {
      username: username,
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
      const getHeaderFunction = centralOperations.getCentralOperationsToken();
      // const getHeaderFunction = getCentralOperationsAuthorizationHeader(operations);
      const authorizationHeader = await getHeaderFunction(AppPurpose.Data);
      const entity = await operations.github.call(
        authorizationHeader,
        'users.getByUsername',
        parameters,
        cacheOptions
      );
      const account = new Account(
        entity,
        this,
        getHeaderFunction.bind(null, AppPurpose.Data)
      );
      return account;
    } catch (error) {
      if (error.status && error.status == /* loose */ 404) {
        error = new Error(
          `The GitHub username ${username} could not be found (or has been deleted)`
        );
        error.status = 404;
        throw error;
      } else if (error) {
        throw wrapError(
          error,
          `Could not get details about account ${username}: ${error.message}`
        );
      }
    }
  }

  get providers(): IProviders {
    return this._providers;
  }
  get config(): any {
    return this.providers.config;
  }

  get insights(): any {
    return this.providers.insights;
  }

  get baseUrl(): string {
    return this._baseUrl;
  }

  get absoluteBaseUrl(): string {
    let baseUrl =
      this.config && this.config.webServer && this.config.webServer.baseUrl
        ? this.config.webServer.baseUrl
        : null;
    if (baseUrl) {
      return baseUrl.endsWith('/') ? baseUrl : `${baseUrl}/`;
    }
    return '/';
  }

  get initialized(): Date {
    return this._initialized;
  }

  get defaults(): ICacheDefaultTimes {
    return this._defaults;
  }

  get github(): RestLibrary {
    return this._github;
  }

  async initialize() {
    const tokenManager = this.tokenManager;
    await tokenManager.initialize();
    tokenManager.getAppIds().map((appId) => {
      const { friendlyName } = tokenManager.getAppById(appId);
      const slug = tokenManager.getSlugById(appId);
      const app = new GitHubApplication(
        this,
        appId,
        slug,
        friendlyName,
        this.getAppAuthorizationHeader.bind(this, tokenManager, appId)
      );
      this._applicationIds.set(appId, app);
    });
    this._initialized = new Date();

    return this;
  }

  getApplicationById(appId: number): GitHubApplication {
    return this._applicationIds.get(appId);
  }

  getApplications(): GitHubApplication[] {
    return Array.from(this._applicationIds.values());
  }

  protected async getAppAuthorizationHeader(
    tokenManager: GitHubTokenManager,
    appId: number
  ): Promise<string> {
    const jwt = await tokenManager
      .getAppById(appId)
      .getAppAuthenticationToken();
    const value = `bearer ${jwt}`;
    return value;
  }

  protected async getAuthorizationHeader(
    organizationName: string,
    organizationSettings: OrganizationSetting,
    legacyOwnerToken: string,
    centralOperationsFallbackToken: string,
    appAuthenticationType: GitHubAppAuthenticationType,
    purpose: AppPurpose
  ): Promise<IAuthorizationHeaderValue> {
    if (
      !this.tokenManager.organizationSupportsAnyPurpose(
        organizationName,
        organizationSettings
      )
    ) {
      const legacyTokenValue =
        legacyOwnerToken || centralOperationsFallbackToken;
      if (!legacyTokenValue) {
        throw new Error(
          `Organization ${organizationName} is not configured with a GitHub app, Personal Access Token ownerToken configuration value, or a fallback central operations token`
        );
      }
      return {
        value: `token ${legacyTokenValue}`,
        purpose: null,
        source: legacyOwnerToken
          ? 'legacyOwnerToken'
          : 'centralOperationsFallbackToken',
      };
    }
    if (!purpose) {
      purpose = AppPurpose.Data;
      console.log(
        `TODO: consider investigating the callback here as to why the getAuthorizationHeader call was not provided a purpose for the ${organizationName} org. falling back to: purpose=${purpose}`
      );
    }
    return this.tokenManager.getOrganizationAuthorizationHeader(
      organizationName,
      purpose,
      organizationSettings,
      appAuthenticationType
    );
  }
}
