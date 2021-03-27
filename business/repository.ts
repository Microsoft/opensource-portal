//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import moment from 'moment';

import { wrapError } from '../utils';
import { IAccountBasics, Organization } from './organization';
import {
  ICacheOptions,
  IPagedCacheOptions,
  IGetAuthorizationHeader,
  IPurposefulGetAuthorizationHeader,
  NoCacheNoBackground,
  IOperationsInstance,
  throwIfNotGitHubCapable,
  throwIfNotCapable,
  IOperationsProviders,
  CoreCapability,
  operationsWithCapability,
  IOperationsCentralOperationsToken,
  IOperationsServiceAccounts,
  IRepositoryGetIssuesOptions,
} from '../transitional';
import * as common from './common';
import { RepositoryPermission } from './repositoryPermission';
import { Collaborator } from './collaborator';
import { TeamPermission } from './teamPermission';
import {
  RepositoryMetadataEntity,
  GitHubRepositoryPermission,
} from '../entities/repositoryMetadata/repositoryMetadata';
import { AppPurpose } from '../github';
import {
  IListPullsParameters,
  GitHubPullRequestState,
  GitHubPullRequestSort,
  GitHubSortDirection,
} from '../lib/github/collections';
import { RepositoryIssue } from './repositoryIssue';
import { IGitHubTeamBasics } from './team';
import {
  CacheDefault,
  DefaultPageSize,
  getMaxAgeSeconds,
  getPageSize,
} from '.';

export interface IGitHubCollaboratorInvitation {
  id: string;
  permissions: GitHubRepositoryPermission;
  created_at: string; // Date
  url: string; // API url
  html_url: string; // user-facing URL
}

export interface IAlternateTokenRequiredOptions extends ICacheOptions {
  alternateToken: string;
}

export interface IGetBranchesOptions extends ICacheOptions {
  protected?: boolean;
}

export interface IGetContentOptions extends ICacheOptions {
  branch?: string;
  tag?: string;
  ref?: string;
}

export enum GitHubCollaboratorAffiliationQuery {
  All = 'all',
  Outside = 'outside',
  Direct = 'direct',
}

export enum GitHubCollaboratorType {
  Outside = 'outside',
  Direct = 'direct',
}

export interface IListContributorsOptions extends IPagedCacheOptions {
  anon?: boolean;
}

export interface IGetCollaboratorsOptions extends IPagedCacheOptions {
  affiliation?: GitHubCollaboratorAffiliationQuery;
}

export interface IGitHubProtectedBranchConfiguration {
  id: string;
  pattern: string;
}

export interface IGetPullsOptions extends ICacheOptions {
  state?: GitHubPullRequestState;
  head?: string;
  base?: string;
  sort?: GitHubPullRequestSort;
  direction?: GitHubSortDirection;
}

export interface ICreateWebhookOptions {
  name?: string;
  active?: boolean;
  config?: {
    url?: string;
    content_type?: string;
    secret?: string;
    insecure_ssl?: string;
  };
  url?: string;
  events?: string[];
}

export enum SecretScanningState {
  Resolved = 'resolved',
  Open = 'open',
}

export enum SecretScanningResolution {
  FalsePositive = 'false_positive',
  WontFix = 'wont_fix',
  Revoked = 'revoked',
  UsedInTests = 'used_in_tests',
}

export interface IGitHubSecretScanningAlert {
  number: number;
  created_at: string;
  url: string;
  html_url: string;
  state: SecretScanningState;
  resolution?: SecretScanningResolution;
  resolved_at?: string;
  resolved_by?: any;
  secret_type: string;
  secret: string;
}

interface IGitHubGetFileParameters {
  owner: string;
  repo: string;
  path: string;
  ref?: string;

  alternateToken?: string;
}

interface IGitHubFileContents {
  type: string;
  encoding: 'base64';
  size: number;
  name: string;
  path: string;
  content: string;
  sha: string;
  url: string;
  git_url: string;
  html_url: string;
  download_url: string;
  _links: {
    git: string;
    self: string;
    html: string;
  };
}

interface ICreateFileParameters {
  owner: string;
  repo: string;
  path: string;
  message: string;
  content: string;
  branch?: string;
  committer?: any;
  sha?: string;

  alternateToken?: string;
}

interface IGitHubGetFileOptions {
  ref?: string;
}

interface ICreateFileOptions {
  branch?: string;
  committer?: any;
  alternateToken?: string;
  sha?: string;
}

interface IGetBranchesParameters {
  owner: string;
  repo: string;
  per_page: number;
  protected?: boolean;
}

export interface IGitHubBranch {
  name: string;
  commit: {
    sha: string;
    url: string;
  };
  protected: boolean;
}

export interface IGitHubBranchDetailed {
  name: string;
  commit: {
    sha: string;
    node_id: string;
    commit: {
      author: {
        name: string;
        date: string; // iso8601
        email: string;
      };
      url: string;
      message: string;
      tree: {
        sha: string;
        url: string;
      };
      committer: {
        name: string;
        date: string;
        email: string;
      };
      verification: {
        verified: boolean;
        reason: string; // 'unsigned', ...
        signature: unknown;
        payload: unknown;
      };
      comment_count: number;
    };
    author: unknown; // basic user, avatar, id, etc.
    parents: unknown[];
    url: string;
    committer: unknown; // basic user
    protected: boolean;
    protection: {
      enabled: boolean;
      required_status_checks: {
        enforcement_level: 'non_admins' | 'admins';
        contexts: string[];
      };
    };
    protection_url: string;
  };
}

export interface IRepositoryBranchAccessProtections {
  allow_deletions: {
    enabled: boolean;
  };
  allow_force_pushes: {
    enabled: boolean;
  };
  enforce_admins: {
    enabled: boolean;
    url: string;
  };
  required_linear_history: {
    enabled: boolean;
  };
  restrictions: {
    users: IAccountBasics[];
    teams: IGitHubTeamBasics[];
    apps: unknown[];
  };
  url: string;
}

interface IRepositoryMoments {
  created?: moment.Moment;
  updated?: moment.Moment;
  pushed?: moment.Moment;
}

interface IRepositoryMomentsAgo {
  created?: string;
  updated?: string;
  pushed?: string;
}

export interface ITemporaryCommandOutput {
  error?: Error;
  message?: string;
}

interface IProtectedBranchRule {
  pattern: string;
}

const safeEntityFieldsForJsonSend = [
  'fork',
  'name',
  'size',
  'forks',
  'license',
  'private',
  'archived',
  'disabled',
  'homepage',
  'language',
  'watchers',
  'pushed_at',
  'created_at',
  'updated_at',
  'description',
  'forks_count',
  'watchers_count',
  'stargazers_count',
  'open_issues_count',
  'id',
];

export class Repository {
  private _entity: any;
  private _baseUrl: string;

  private _awesomeness: number;

  private _getAuthorizationHeader: IPurposefulGetAuthorizationHeader;
  private _getSpecificAuthorizationHeader: IPurposefulGetAuthorizationHeader;
  private _operations: IOperationsInstance;

  private _organization: Organization;

  private _name: string;

  private _moments: IRepositoryMoments;

  getEntity(): any {
    return this._entity;
  }

  asJson() {
    const organizationSubset = {
      organization: {
        login: this.organization.name,
        id: this.organization.id,
      },
    };
    const entity = this.getEntity();
    const safeClone = {};
    for (let i = 0; i < safeEntityFieldsForJsonSend.length; i++) {
      const key = safeEntityFieldsForJsonSend[i];
      if (entity[key] !== undefined) {
        safeClone[key] = entity[key];
      }
    }
    return Object.assign(organizationSubset, safeClone);
  }

  get id(): number {
    return this._entity ? this._entity.id : null;
  }
  get name(): string {
    return this._entity ? this._entity.name : this._name;
  }
  get full_name(): string {
    return this._entity ? this._entity.full_name : null;
  }
  get private(): boolean {
    return this._entity ? this._entity.private : false;
  }
  get html_url(): string {
    return this._entity ? this._entity.html_url : null;
  }
  get description(): string {
    return this._entity ? this._entity.description : null;
  }
  get fork(): boolean {
    return this._entity ? this._entity.fork : null;
  }
  get url(): string {
    return this._entity ? this._entity.url : null;
  }
  get archived(): boolean {
    return this._entity ? this._entity.archived : false;
  }
  get created_at(): Date {
    return this._entity ? this._entity.created_at : null;
  }
  get updated_at(): Date {
    return this._entity ? this._entity.updated_at : null;
  }
  get pushed_at(): Date {
    return this._entity ? this._entity.pushed_at : null;
  }
  get git_url(): string {
    return this._entity ? this._entity.git_url : null;
  }
  get homepage(): string {
    return this._entity ? this._entity.homepage : null;
  }
  get size(): any {
    return this._entity ? this._entity.size : null;
  }
  get stargazers_count(): any {
    return this._entity ? this._entity.stargazers_count : null;
  }
  get watchers_count(): any {
    return this._entity ? this._entity.watchers_count : null;
  }
  get language(): string {
    return this._entity ? this._entity.language : null;
  }
  get has_issues(): boolean {
    return this._entity ? this._entity.has_issues : null;
  }
  get has_wiki(): boolean {
    return this._entity ? this._entity.has_wiki : null;
  }
  get has_pages(): boolean {
    return this._entity ? this._entity.has_pages : null;
  }
  get forks_count(): any {
    return this._entity ? this._entity.forks_count : null;
  }
  get open_issues_count(): any {
    return this._entity ? this._entity.open_issues_count : null;
  }
  get forks(): any {
    return this._entity ? this._entity.forks : null;
  }
  get open_issues(): any {
    return this._entity ? this._entity.open_issues : null;
  }
  get watchers(): any {
    return this._entity ? this._entity.watchers : null;
  }
  get license(): any {
    return this._entity ? this._entity.license : null;
  }
  get default_branch(): any {
    return this._entity ? this._entity.default_branch : null;
  }
  get clone_url(): any {
    return this._entity ? this._entity.clone_url : null;
  }
  get ssh_url(): any {
    return this._entity ? this._entity.ssh_url : null;
  }
  get parent(): any {
    return this._entity ? this._entity.parent : null;
  }

  get organization(): Organization {
    return this._organization;
  }

  get baseUrl(): string {
    return this._baseUrl;
  }

  get absoluteBaseUrl(): string {
    return this.organization.absoluteBaseUrl + 'repos/' + this.name + '/';
  }

  constructor(
    organization: Organization,
    entity: any,
    getAuthorizationHeader: IPurposefulGetAuthorizationHeader,
    getSpecificAuthorizationHeader: IPurposefulGetAuthorizationHeader,
    operations: IOperationsInstance
  ) {
    this._organization = organization;
    this._entity = entity;
    this._baseUrl = organization.baseUrl + 'repos/' + this.name + '/';
    this._getAuthorizationHeader = getAuthorizationHeader;
    this._getSpecificAuthorizationHeader = getSpecificAuthorizationHeader;
    this._operations = operations;
  }

  get moment(): IRepositoryMoments {
    if (!this._moments) {
      this._moments = {
        updated: this.updated_at ? moment(this.updated_at) : undefined,
        pushed: this.pushed_at ? moment(this.pushed_at) : undefined,
        created: this.created_at ? moment(this.created_at) : undefined,
      };
    }
    return this._moments;
  }

  get momentDisplay(): IRepositoryMomentsAgo {
    const moments = this.moment;
    return {
      updated: moments.updated ? moments.updated.fromNow() : undefined,
      created: moments.created ? moments.created.fromNow() : undefined,
      pushed: moments.pushed ? moments.pushed.fromNow() : undefined,
    };
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
    if (this.id && !this.name) {
      try {
        const lookupById = await this.organization.getRepositoryById(this.id);
        this._entity = lookupById.getEntity();
        this._name = this._entity.name;
      } catch (getByIdError) {
        throw getByIdError;
      }
    }
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
    };
    const cacheOptions: ICacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoDetailsStaleSeconds,
        options
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    try {
      const entity = await operations.github.call(
        this.authorize(AppPurpose.Data),
        'repos.get',
        parameters,
        cacheOptions
      );
      this._entity = entity;
      return entity;
    } catch (error) {
      const notFound = error.status && error.status == /* loose */ 404;
      error = wrapError(
        error,
        notFound
          ? 'The repo could not be found.'
          : 'Could not get details about the repo.',
        notFound
      );
      if (notFound) {
        error.status = 404;
      }
      throw error;
    }
  }

  async getRepositoryMetadata(): Promise<RepositoryMetadataEntity> {
    const operations = throwIfNotCapable<IOperationsProviders>(
      this._operations,
      CoreCapability.Providers
    );
    const repositoryMetadataProvider =
      operations.providers.repositoryMetadataProvider;
    try {
      return await repositoryMetadataProvider.getRepositoryMetadata(
        this.id.toString()
      );
    } catch (getMetadataError) {
      return null;
    }
  }

  async getBranches(
    cacheOptions: IGetBranchesOptions
  ): Promise<IGitHubBranch[]> {
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters: IGetBranchesParameters = {
      owner: this.organization.name,
      repo: this.name,
      per_page: getPageSize(operations),
    };
    if (cacheOptions.protected !== undefined) {
      parameters.protected = cacheOptions.protected;
    }
    delete cacheOptions.protected;
    if (!cacheOptions.maxAgeSeconds) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.repoBranchesStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    return github.collections.getRepoBranches(
      this.authorize(AppPurpose.Data),
      parameters,
      cacheOptions
    );
  }

  async getPulls(options?: IGetPullsOptions): Promise<any> {
    await this.organization.requireUpdatesApp('getPulls');
    // CONSIDER: might really need to probe for the app and pick which has pull request access
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const cacheOptions: ICacheOptions = {};
    const parameters: IListPullsParameters = Object.assign({}, options || {}, {
      owner: this.organization.name,
      repo: this.name,
      per_page: getPageSize(operations),
    });
    if (options && options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
      delete parameters['backgroundRefresh'];
    }
    if (options && options.maxAgeSeconds !== undefined) {
      cacheOptions.maxAgeSeconds = options.maxAgeSeconds;
      delete parameters['maxAgeSeconds'];
    }
    if (cacheOptions.maxAgeSeconds === undefined) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.repoPullsStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    return github.collections.getRepoPullRequests(
      this.authorize(AppPurpose.Updates),
      parameters,
      cacheOptions
    );
  }

  async getContent(path: string, options?: IGetContentOptions): Promise<any> {
    options = options || {};
    const ref = options.branch || options.tag || options.ref || 'master';
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      path: path,
      ref: ref,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    return operations.github.call(
      this.authorize(AppPurpose.Data),
      'repos.getContent',
      parameters
    );
  }

  async getLastCommitToBranch(branchName: string): Promise<string> {
    await this.organization.requireUpdatesApp('getLastCommitToBranch');
    const options = {
      owner: this.organization.name,
      repo: this.name,
      ref: `heads/${branchName}`,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    const data = await operations.github.requestAsPost(
      this.authorize(AppPurpose.Updates),
      'GET /repos/:owner/:repo/git/ref/:ref',
      options
    );
    return data.object.sha;
  }

  async renameDefaultBranch(
    newBranchName?: string
  ): Promise<ITemporaryCommandOutput[]> {
    newBranchName = newBranchName || 'main';
    await this.organization.requireUpdatesApp('renameDefaultBranch');
    const output: ITemporaryCommandOutput[] = [];
    try {
      await this.getDetails(NoCacheNoBackground);
      if (this.default_branch === newBranchName) {
        return [
          {
            message: `The default branch is already '${newBranchName}' for the repo ${this.full_name}. No further action required.`,
          },
        ];
      }
      const currentBranchName = this.default_branch;
      const sha = await this.getLastCommitToBranch(currentBranchName);
      // TODO: what if the branch already exists? Should let this keep running to update more PRs until done.
      await this.createNewBranch(sha, newBranchName);
      output.push({
        message: `Created a new branch '${newBranchName}' from '${currentBranchName}' which points to SHA ${sha}.`,
      });
      const branchProtectionRules = await this.listBranchProtectionRules();
      // there can only be one protection per pattern
      const branchProtection = branchProtectionRules.find(
        (rule: IProtectedBranchRule) => rule.pattern === currentBranchName
      );
      if (branchProtectionRules.length > 0) {
        const branchMessage = branchProtection
          ? `. The default branch is protected and the protection will be shifted to target '${newBranchName}'.`
          : ', but no action is required as the default branch is not protected.';
        output.push({
          message: `There are ${branchProtectionRules.length} protected branches${branchMessage}`,
        });
      }
      if (branchProtection) {
        const { id } = branchProtection;
        await this.updateBranchProtectionRule(id, newBranchName);
        output.push({
          message: `Branch protection rule shifted from the old branch '${currentBranchName}' to '${newBranchName}'.`,
        });
      }
      const pulls = await this.getPulls({
        state: GitHubPullRequestState.Open,
        base: currentBranchName,
      });
      if (pulls.length === 0) {
        output.push({
          message: `No open pull requests targeting '${currentBranchName}' to update.`,
        });
      } else {
        output.push({
          message: `There are ${pulls.length} open pull requests targeting '${currentBranchName}' that will be updated to '${newBranchName}'.`,
        });
      }
      for (const pull of pulls) {
        try {
          await this.patchPullRequestBranch(pull.number, newBranchName);
          output.push({
            message: `Pull request #${pull.number} has been updated to target '${newBranchName}' (URL: https://github.com/${this.full_name}/pull/${pull.number}, title: '${pull.title}').`,
          });
        } catch (pullError) {
          // To keep the operation going, failed pulls do not short-circuit the process
          output.push({
            message: `Pull request #${pull.number} could not be updated to target '${newBranchName}. Please inspect https://github.com/${this.full_name}/pull/${pull.number}.`,
          });
          output.push({ error: pullError });
        }
      }
      output.push({
        message: `Setting the default branch of the repo ${this.full_name} to '${newBranchName}'.`,
      });
      await this.setDefaultBranch(newBranchName);
      output.push({ message: `Deleting the branch '${currentBranchName}'.` });
      await this.deleteBranch(currentBranchName);
      output.push({
        message: `The repo's default branch is now '${newBranchName}'. Thank you. You may inspect the repo at https://github.com/${this.full_name}/.`,
      });
    } catch (error) {
      output.push({
        message: `The branch rename to '${newBranchName}' was not completely successful. Please review the error and inspect the repo.`,
      });
      output.push({ error });
    }
    return output;
  }

  async patchPullRequestBranch(
    number: string,
    targetBranch: string
  ): Promise<void> {
    await this.organization.requireUpdatesApp('patchPullRequestBranch');
    const operations = throwIfNotGitHubCapable(this._operations);
    const options = {
      owner: this.organization.name,
      repo: this.name,
      pull_number: number,
      base: targetBranch,
    };
    await operations.github.requestAsPost(
      this.authorize(AppPurpose.Updates),
      'PATCH /repos/:owner/:repo/pulls/:pull_number',
      options
    );
  }

  async createNewBranch(sha: string, newBranchName: string): Promise<void> {
    await this.organization.requireUpdatesApp('createNewBranch');
    const options = {
      owner: this.organization.name,
      repo: this.name,
      ref: `refs/heads/${newBranchName}`,
      sha,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    await operations.github.requestAsPost(
      this.authorize(AppPurpose.Updates),
      'POST /repos/:owner/:repo/git/refs',
      options
    );
  }

  async updateBranchProtectionRule(
    id: string,
    newPattern: string
  ): Promise<void> {
    await this.organization.requireUpdatesApp('updateBranchProtectionRule');
    const operations = throwIfNotGitHubCapable(this._operations);
    const mutation = `mutation($branchProtectionRuleId:ID!,$pattern:String!) {
      updateBranchProtectionRule (input:{branchProtectionRuleId:$branchProtectionRuleId,pattern:$pattern}) {
        branchProtectionRule {
          id,
          pattern
        }
      }
    }`;
    try {
      await operations.github.graphql(
        this.authorize(AppPurpose.Updates),
        mutation,
        {
          branchProtectionRuleId: id,
          pattern: newPattern,
        }
      );
    } catch (error) {
      throw error;
    }
  }

  async listBranchProtectionRules(): Promise<
    IGitHubProtectedBranchConfiguration[]
  > {
    await this.organization.requireUpdatesApp('listBranchProtectionRules');
    const query = `query($owner: String!, $repo: String!) {
      repository(owner:$owner,name:$repo) {
        branchProtectionRules(first:100) {
          nodes {
            id
            pattern
          }
        }
      }
    }`;
    const operations = throwIfNotGitHubCapable(this._operations);
    try {
      const {
        repository: {
          branchProtectionRules: { nodes: branchProtectionRules },
        },
      } = await operations.github.graphql(
        this.authorize(AppPurpose.Updates),
        query,
        {
          owner: this.organization.name,
          repo: this.name,
        }
      );
      return branchProtectionRules as IGitHubProtectedBranchConfiguration[];
    } catch (error) {
      throw error;
    }
  }

  async getProtectedBranchAccessRestrictions(
    branchName: string,
    cacheOptions?: ICacheOptions
  ): Promise<IRepositoryBranchAccessProtections> {
    // NOTE: GitHub has a "100-item limit" currently. This is an object response and not
    // technically paginated.
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      branch: branchName,
    };
    if (!cacheOptions.maxAgeSeconds) {
      //cacheOptions.maxAgeSeconds = operations.defaults.orgRepoCollaboratorStaleSeconds;
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      //cacheOptions.backgroundRefresh = true;
    }
    Object.assign(parameters, cacheOptions);
    // GET /repos/{owner}/{repo}/branches/{branch}/protection/restrictions
    const protections = await github.call(
      this.authorize(AppPurpose.Data),
      'repos.getBranchProtection',
      parameters
    );
    return protections as IRepositoryBranchAccessProtections;
  }

  async setDefaultBranch(defaultBranchName: string): Promise<void> {
    await this.organization.requireUpdatesApp('setDefaultBranch');
    const options = {
      owner: this.organization.name,
      repo: this.name,
      name: this.name,
      default_branch: defaultBranchName,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    await operations.github.requestAsPost(
      this.authorize(AppPurpose.Updates),
      'PATCH /repos/:owner/:repo',
      options
    );
  }

  async deleteBranch(branchName: string): Promise<void> {
    await this.organization.requireUpdatesApp('deleteBranch');
    const options = {
      owner: this.organization.name,
      repo: this.name,
      ref: `heads/${branchName}`,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    await operations.github.requestAsPost(
      this.authorize(AppPurpose.Updates),
      'DELETE /repos/:owner/:repo/git/refs/:ref',
      options
    );
  }

  async getPages(options?: ICacheOptions): Promise<any> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
    };
    const cacheOptions: ICacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoDetailsStaleSeconds,
        options
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    try {
      // CONSIDER: need a fallback authentication approach: try and app for a specific capability (the installation knows) and fallback to central ops
      // const centralOps = operationsWithCapability<IOperationsCentralOperationsToken>(operations, CoreCapability.GitHubCentralOperations);
      const tokenSource = this._getSpecificAuthorizationHeader(AppPurpose.Data); // centralOps ? centralOps.getCentralOperationsToken()(AppPurpose.Data) :
      const token = await tokenSource;
      return await operations.github.call(
        token,
        'repos.getPages',
        parameters,
        cacheOptions
      );
    } catch (error) {
      const notFound = error.status && error.status == /* loose */ 404;
      error = wrapError(
        error,
        notFound
          ? 'The repo is not configured for pages.'
          : 'Could not get details about the repo pages configuration.',
        notFound
      );
      if (notFound) {
        error.status = 404;
      }
      throw error;
    }
  }

  async checkCollaborator(
    username: string,
    cacheOptions?: ICacheOptions
  ): Promise<boolean> {
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      username,

      allowEmptyResponse: true,
    };
    if (!cacheOptions.maxAgeSeconds) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoTeamsStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    try {
      const ok = await operations.github.post(
        this.authorize(AppPurpose.Data),
        'repos.checkCollaborator',
        parameters
      );
      return true;
    } catch (error) {
      if (error && error.status == /* loose */ 404) {
        return false;
      }
      throw wrapError(
        error,
        `Could not verify the collaborator level for user ${username} in the repo ${this.organization.name}/${this.name}`
      );
    }
  }

  async isUserPermissionFromTeam(login: string, permission: any) {
    const teamPermissions = await this.getTeamPermissions();
    for (let i = 0; i < teamPermissions.length; i++) {
      try {
        const teamPermission = teamPermissions[i];
        if (teamPermission.permission === permission) {
          if (await teamPermission.team.isMember(login)) {
            return true;
          }
        }
      } catch (ignore) {
        /* ignore */
      }
    }
    return false;
  }

  async getCollaborator(
    username: string,
    cacheOptions?: ICacheOptions
  ): Promise<RepositoryPermission> {
    // This call is used in customer-facing sites by permissions middleware
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      username: username,
    };
    if (!cacheOptions.maxAgeSeconds) {
      //cacheOptions.maxAgeSeconds = operations.defaults.orgRepoCollaboratorStaleSeconds;
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      //cacheOptions.backgroundRefresh = true;
    }
    Object.assign(parameters, cacheOptions);
    const userPermissionLevel = await github.call(
      this.authorize(AppPurpose.CustomerFacing),
      'repos.getCollaboratorPermissionLevel',
      parameters
    );
    return new RepositoryPermission(userPermissionLevel);
  }

  async listContributors(
    cacheOptions?: IListContributorsOptions
  ): Promise<any[]> {
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      per_page: getPageSize(operations),
      anon: cacheOptions.anon || false,
    };
    delete cacheOptions.anon;
    if (!cacheOptions.maxAgeSeconds) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoCollaboratorsStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    const contributors = await github.collections.getRepoContributors(
      this.authorize(AppPurpose.Data),
      parameters,
      cacheOptions
    );
    // const contributors = common.createInstances<Collaborator>(this, collaboratorPermissionFromEntity, contributorsEntities);
    return contributors;
  }

  async getCollaborators(
    cacheOptions?: IGetCollaboratorsOptions
  ): Promise<Collaborator[]> {
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      per_page: getPageSize(operations),
      affiliation:
        cacheOptions.affiliation || GitHubCollaboratorAffiliationQuery.All,
    };
    delete cacheOptions.affiliation;
    if (!cacheOptions.maxAgeSeconds) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoCollaboratorsStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    const collaboratorEntities = await github.collections.getRepoCollaborators(
      this.authorize(AppPurpose.Data),
      parameters,
      cacheOptions
    );
    const collaborators = common.createInstances<Collaborator>(
      this,
      collaboratorPermissionFromEntity,
      collaboratorEntities
    );
    return collaborators;
  }

  async addCollaborator(
    username: string,
    permission: GitHubRepositoryPermission
  ): Promise<IGitHubCollaboratorInvitation> {
    // BREAKING CHANGE in the GitHub API: as of August 2017, this is "inviteCollaborator', it does not automatically add
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      username: username,
      permission: permission,
    };
    // CONSIDER: If status code 404 on return, the username does not exist on GitHub as entered
    const response = await github.post(
      this.authorize(AppPurpose.Operations),
      'repos.addCollaborator',
      parameters
    );
    return response as IGitHubCollaboratorInvitation;
  }

  async acceptCollaborationInvite(
    invitationId: string,
    options: IAlternateTokenRequiredOptions
  ): Promise<any> {
    // This could go in Account _or_ here in Repository
    if (!options || !options.alternateToken) {
      throw new Error(
        'acceptCollaborationInvite requires options.alternateToken'
      );
    }
    const alternateTokenHeader = `token ${options.alternateToken}`;
    const parameters = {
      invitation_id: invitationId,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    return operations.github.post(
      alternateTokenHeader,
      'repos.acceptInvitation',
      parameters
    );
  }

  removeCollaborator(username: string): Promise<any> {
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      username: username,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.removeCollaborator',
      parameters
    );
  }

  delete(): Promise<void> {
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
    };
    const operations = throwIfNotGitHubCapable(this._operations);
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.delete',
      parameters
    );
  }

  createFile(
    path: string,
    base64Content: string,
    commitMessage: string,
    options?: ICreateFileOptions
  ): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters: ICreateFileParameters = Object.assign(
      {
        owner: this.organization.name,
        repo: this.name,
        path,
        message: commitMessage,
        content: base64Content,
      },
      options
    );
    if (options && options.sha) {
      parameters.sha = options.sha;
    }
    if (options && options.branch) {
      parameters.branch = options.branch;
    }
    if (options && options.committer) {
      parameters.committer = options.committer;
    }
    const alternateHeader = options.alternateToken
      ? `token ${options.alternateToken}`
      : null;
    return operations.github.post(
      alternateHeader || this.authorize(AppPurpose.Operations),
      'repos.createOrUpdateFileContents',
      parameters
    );
  }

  getFile(
    path: string,
    options?: IGitHubGetFileOptions
  ): Promise<IGitHubFileContents> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters: IGitHubGetFileParameters = Object.assign(
      {
        owner: this.organization.name,
        repo: this.name,
        path,
      },
      options
    );
    if (options && options.ref) {
      parameters.ref = options.ref;
    }
    // const alternateHeader = options.alternateToken ? `token ${options.alternateToken}` : null;
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.getContent',
      parameters
    );
  }

  getFiles(
    path: string,
    options?: IGitHubGetFileOptions
  ): Promise<IGitHubFileContents[]> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters: IGitHubGetFileParameters = Object.assign(
      {
        owner: this.organization.name,
        repo: this.name,
        path,
      },
      options
    );
    if (options.ref) {
      parameters.ref = options.ref;
    }
    // const alternateHeader = options.alternateToken ? `token ${options.alternateToken}` : null;
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.getContent',
      parameters
    );
  }

  async setTeamPermission(
    teamId: number,
    newPermission: GitHubRepositoryPermission
  ): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const team = this.organization.team(teamId);
    // CONSIDER: note the performance penalty on the slug resolution; the alternate path has not been working for GitHub Apps
    await team.getDetails();
    const options = {
      org: this.organization.name,
      team_slug: team.slug,
      owner: this.organization.name,
      repo: this.name,
      permission: newPermission,
    };
    // alternate version of: 'teams.octokit.teams.addOrUpdateRepoPermissionsInOrg': 'PUT /organizations/:org_id/team/:team_id/repos/:owner/:repo'
    const result = await operations.github.post(
      this.authorize(AppPurpose.Operations),
      'teams.addOrUpdateRepoPermissionsInOrg',
      options
    );
    return result;
  }

  removeTeamPermission(teamId: number): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const options = {
      org_id: this.organization.id.toString(),
      team_id: teamId,
      owner: this.organization.name,
      repo: this.name,
    };
    // alternate version of: 'teams.removeRepoInOrg'
    return operations.github.requestAsPost(
      this.authorize(AppPurpose.Operations),
      'DELETE /organizations/:org_id/team/:team_id/repos/:owner/:repo',
      options
    );
  }

  async getWebhooks(options?: ICacheOptions): Promise<any> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
    };
    const cacheOptions: ICacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoWebhooksStaleSeconds,
        options
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    return operations.github.call(
      this.authorize(AppPurpose.Data),
      'repos.listWebhooks',
      parameters,
      cacheOptions
    );
  }

  deleteWebhook(webhookId: string): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      id: webhookId,
    };
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.deleteWebhook',
      parameters
    );
  }

  createWebhook(options: ICreateWebhookOptions): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    delete options['owner'];
    delete options['repo'];
    const parameters = Object.assign(
      {
        owner: this.organization.name,
        repo: this.name,
      },
      options
    );
    // Smart defaults: create an active JSON web hook to the 'url' option
    if (!options.name) {
      parameters.name = 'web';
    }
    if (options.active === undefined) {
      parameters.active = true;
    }
    if (options.url && !options.config) {
      delete parameters.url;
      parameters.config = {
        url: options.url,
        content_type: 'json',
      };
    }
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.createWebhook',
      parameters
    );
  }

  async editPublicPrivate(options): Promise<void> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    if (options.private !== true && options.private !== false) {
      throw new Error(
        'editPublicPrivate.options requires private to be set to true or false'
      );
    }
    const parameters = Object.assign(
      {
        owner: this.organization.name,
        repo: this.name,
      },
      {
        private: options.private,
      }
    );
    // BUG: GitHub Apps do not work with locking down no repository permissions as documented here: https://github.community/t5/GitHub-API-Development-and/GitHub-App-cannot-patch-repo-visibility-in-org-with-repo/m-p/33448#M3150
    // const token = this._operations.getCentralOperationsToken();
    // return this._operations.github.post(token, 'repos.update', parameters);
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.update',
      parameters
    );
  }

  async archive(): Promise<void> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = Object.assign(
      {
        owner: this.organization.name,
        repo: this.name,
      },
      {
        archived: true,
      }
    );
    return operations.github.post(
      this.authorize(AppPurpose.Operations),
      'repos.update',
      parameters
    );
  }

  async getTeamPermissions(
    cacheOptions?: IPagedCacheOptions
  ): Promise<TeamPermission[]> {
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      per_page: getPageSize(operations),
    };
    if (!cacheOptions.maxAgeSeconds) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoTeamsStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    const permissionEntities = await github.collections.getRepoTeams(
      this.authorize(AppPurpose.Data),
      parameters,
      cacheOptions
    );
    const teamPermissions = common.createInstances<TeamPermission>(
      this,
      teamPermissionFromEntity,
      permissionEntities
    );
    return teamPermissions;
  }

  async checkTeamManages(
    teamId: string,
    cacheOptions?: ICacheOptions
  ): Promise<boolean> {
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      owner: this.organization.name,
      org_id: this.organization.id.toString(),
      repo: this.name,
      team_id: teamId,

      allowEmptyResponse: true,
    };
    if (!cacheOptions.maxAgeSeconds) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoTeamsStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    try {
      // this is the alternate form of 'teams.checkPermissionsForRepoInOrg'
      await operations.github.requestAsPost(
        this.authorize(AppPurpose.Data),
        'GET /organizations/:org_id/team/:team_id/repos/:owner/:repo',
        parameters
      );
      return true;
    } catch (error) {
      if (error && error.status == /* loose */ 404) {
        return false;
      }
      throw wrapError(
        error,
        `Could not verify the team management permissions of the repo ${this.organization.name}/${this.name} for team ${teamId}`
      );
    }
  }

  async enableSecretScanning(): Promise<boolean> {
    // NOTE: this is an experimental API as part of the program public beta, and likely not available
    // to most users. Expect this call to fail.
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      repo_id: this.id.toString(),
    };
    try {
      await operations.github.requestAsPost(
        this.authorize(AppPurpose.Operations),
        'PUT /repositories/:repo_id/secret-scanning',
        parameters
      );
      return true;
    } catch (error) {
      if (
        error &&
        error.status == /* loose */ 404 &&
        error.message === 'Secret scanning is disabled'
      ) {
        return false;
      }
      throw error;
    }
  }

  async getSecretScanningAlerts(
    cacheOptions?: ICacheOptions
  ): Promise<IGitHubSecretScanningAlert[]> {
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      // repo_id: this.id.toString(),
      owner: this.organization.name,
      repo: this.name,
      per_page: 100,
      // state: 'open' | 'resolved'
    };
    // NOTE: not paginating for now
    // if (!cacheOptions.maxAgeSeconds) {
    //   cacheOptions.maxAgeSeconds = operations.defaults.orgRepoTeamsStaleSeconds;
    // }
    // if (cacheOptions.backgroundRefresh === undefined) {
    //   cacheOptions.backgroundRefresh = true;
    // }
    try {
      // using requestAsPost to _not cache_ the secrets for now
      const response = await operations.github.requestAsPost(
        this.authorize(AppPurpose.Data),
        'GET /repos/:owner/:repo/secret-scanning/alerts',
        parameters
      );
      return response as IGitHubSecretScanningAlert[];
    } catch (error) {
      if (
        error &&
        error.status == /* loose */ 404 &&
        error.message === 'Secret scanning is disabled on this repository.'
      ) {
        throw error;
      }
      throw error;
    }
  }

  async checkSecretScanning(cacheOptions?: ICacheOptions): Promise<boolean> {
    // NOTE: this is an experimental API as part of the program public beta, and likely not available
    // to most users. Expect this call to fail.
    cacheOptions = cacheOptions || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      repo_id: this.id.toString(),
    };
    if (!cacheOptions.maxAgeSeconds) {
      cacheOptions.maxAgeSeconds = getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoTeamsStaleSeconds
      );
    }
    if (cacheOptions.backgroundRefresh === undefined) {
      cacheOptions.backgroundRefresh = true;
    }
    try {
      await operations.github.requestAsPost(
        this.authorize(AppPurpose.Operations),
        'GET /repositories/:repo_id/secret-scanning',
        parameters
      );
      return true;
    } catch (error) {
      if (
        error &&
        error.status == /* loose */ 404 &&
        error.message === 'Secret scanning is disabled'
      ) {
        return false;
      }
      throw error;
    }
  }

  async getAdministrators(
    excludeOwners = true,
    excludeBroadAndSystemTeams = true
  ): Promise<string[]> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const opsSystemAccounts = operationsWithCapability<IOperationsServiceAccounts>(
      operations,
      CoreCapability.ServiceAccounts
    );
    const owners = await this._organization.getOwners();
    const ownersSet = new Set<string>(owners.map((o) => o.login.toLowerCase()));
    const actualCollaborators = await this.getCollaborators({
      affiliation: GitHubCollaboratorAffiliationQuery.Direct,
    });
    let collaborators = actualCollaborators.filter(
      (c) => c.permissions?.admin === true
    );
    // No system accounts or owners
    if (opsSystemAccounts) {
      collaborators = collaborators.filter(
        (c) => false === opsSystemAccounts.isSystemAccountByUsername(c.login)
      );
    }
    if (excludeOwners) {
      collaborators = collaborators.filter(
        (c) => false === ownersSet.has(c.login.toLowerCase())
      );
    }
    const users = new Set<string>(
      collaborators.map((c) => c.login.toLowerCase())
    );
    let teams = (await this.getTeamPermissions()).filter(
      (tp) => tp.permission === 'admin'
    );
    for (let i = 0; i < teams.length; i++) {
      const team = teams[i];
      if (
        excludeBroadAndSystemTeams &&
        (team.team.isSystemTeam || team.team.isBroadAccessTeam)
      ) {
        // Do not include broad access teams
        continue;
      }
      const members = await team.team.getMembers();
      for (let j = 0; j < members.length; j++) {
        const tm = members[j];
        const login = tm.login.toLowerCase();
        if (
          !ownersSet.has(login) &&
          (!opsSystemAccounts ||
            !opsSystemAccounts.isSystemAccountByUsername(login))
        ) {
          users.add(login.toLowerCase());
        }
      }
    }
    return Array.from(users.values());
  }

  async getPushers(): Promise<string[]> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const opsSystemAccounts = operationsWithCapability<IOperationsServiceAccounts>(
      operations,
      CoreCapability.ServiceAccounts
    );
    // duplicated code from getAdministrators
    const owners = await this._organization.getOwners();
    const ownersSet = new Set<string>(owners.map((o) => o.login.toLowerCase()));
    const actualCollaborators = await this.getCollaborators({
      affiliation: GitHubCollaboratorAffiliationQuery.Direct,
    });
    let collaborators = actualCollaborators.filter(
      (c) => c.permissions?.push === true
    );
    // No system accounts or owners
    if (opsSystemAccounts) {
      collaborators = collaborators.filter(
        (c) => false === opsSystemAccounts.isSystemAccountByUsername(c.login)
      );
    }
    collaborators = collaborators.filter(
      (c) => false === ownersSet.has(c.login.toLowerCase())
    );
    const users = new Set<string>(
      collaborators.map((c) => c.login.toLowerCase())
    );
    let teams = (await this.getTeamPermissions()).filter(
      (tp) => tp.permission === 'push'
    );
    for (let i = 0; i < teams.length; i++) {
      const team = teams[i];
      if (team.team.isSystemTeam || team.team.isBroadAccessTeam) {
        // Do not include broad access teams
        continue;
      }
      const members = await team.team.getMembers();
      for (let j = 0; j < members.length; j++) {
        const tm = members[j];
        const login = tm.login.toLowerCase();
        if (
          !ownersSet.has(login) &&
          (!opsSystemAccounts ||
            !opsSystemAccounts.isSystemAccountByUsername(login))
        ) {
          users.add(login.toLowerCase());
        }
      }
    }
    return Array.from(users.values());
  }

  async getPullers(
    excludeBroadTeamsAndOwners: boolean = true
  ): Promise<string[]> {
    // duplicated code from getAdministrators
    if (!this.private) {
      return [];
    }
    const operations = throwIfNotGitHubCapable(this._operations);
    const opsSystemAccounts = operationsWithCapability<IOperationsServiceAccounts>(
      operations,
      CoreCapability.ServiceAccounts
    );
    const owners = await this._organization.getOwners();
    const ownersSet = new Set<string>(owners.map((o) => o.login.toLowerCase()));
    const actualCollaborators = await this.getCollaborators({
      affiliation: GitHubCollaboratorAffiliationQuery.Direct,
    });
    let collaborators = actualCollaborators.filter(
      (c) => c.permissions?.pull === true
    );
    // No system accounts or owners
    if (opsSystemAccounts) {
      collaborators = collaborators.filter(
        (c) => false === opsSystemAccounts.isSystemAccountByUsername(c.login)
      );
    }
    if (excludeBroadTeamsAndOwners) {
      collaborators = collaborators.filter(
        (c) => false === ownersSet.has(c.login.toLowerCase())
      );
    }
    const users = new Set<string>(
      collaborators.map((c) => c.login.toLowerCase())
    );
    let teams = (await this.getTeamPermissions()).filter(
      (tp) => tp.permission === 'pull'
    );
    for (let i = 0; i < teams.length; i++) {
      const team = teams[i];
      if (
        excludeBroadTeamsAndOwners &&
        (team.team.isSystemTeam || team.team.isBroadAccessTeam)
      ) {
        // Do not include broad access teams
        continue;
      }
      const members = await team.team.getMembers();
      for (let j = 0; j < members.length; j++) {
        const tm = members[j];
        const login = tm.login.toLowerCase();
        if (
          !ownersSet.has(login) &&
          (!opsSystemAccounts ||
            !opsSystemAccounts.isSystemAccountByUsername(login))
        ) {
          users.add(login.toLowerCase());
        }
      }
    }
    return Array.from(users.values());
  }

  private authorize(purpose: AppPurpose): IGetAuthorizationHeader | string {
    const getAuthorizationHeader = this._getAuthorizationHeader.bind(
      this,
      purpose
    ) as IGetAuthorizationHeader;
    return getAuthorizationHeader;
  }

  public static SortByAwesomeness(a: Repository, b: Repository) {
    return b.computeAwesomeness() - a.computeAwesomeness();
  }

  public computeAwesomeness() {
    if (this._awesomeness) {
      return this._awesomeness;
    }
    const repo = this;
    const pushAwesomeness = 1000;
    const pushHalfLife = 42 * Math.E * Math.pow(10, -15);
    const starAwesomeness = (10 + Math.PI) * Math.pow(10, 13);
    const maxValue = 32767;
    if (!repo.pushed_at) {
      return 0;
    }
    const pushTicks =
      (moment.utc().valueOf() - moment.utc(repo.pushed_at).valueOf()) * 10000;
    const createdTicks =
      (moment.utc().valueOf() - moment.utc(repo.created_at).valueOf()) * 10000;
    // People power, if you have a high star factor (i.e. stars per day) then you
    // are definitely awesome.
    let awesomeness = (starAwesomeness * repo.stargazers_count) / createdTicks;
    // Make it so a recent contribution pushes you up the stack, but make the effect
    // fade quickly (as determined by the halflife of a push.
    awesomeness +=
      pushAwesomeness * Math.pow(Math.E, -1 * pushHalfLife * pushTicks);
    // Everyone who makes their code open source is a little bit awesome.
    ++awesomeness;
    if (awesomeness > maxValue) {
      awesomeness = maxValue;
    }
    this._awesomeness = awesomeness;
    return this._awesomeness;
  }

  async getIssues(
    options?: IRepositoryGetIssuesOptions
  ): Promise<RepositoryIssue[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const github = operations.github;
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      per_page: getPageSize(operations),
      milestone: options.milestone,
      state: options.state,
      assignee: options.assignee,
      creator: options.creator,
      mentioned: options.mentioned,
      labels: options.labels,
      sort: options.sort,
      direction: options.direction,
      since: options.since ? options.since.toISOString() : undefined,
    };
    const cacheOptions: IPagedCacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoTeamsStaleSeconds,
        options
      ),
      backgroundRefresh:
        options.backgroundRefresh !== undefined
          ? options.backgroundRefresh
          : true,
      pageRequestDelay: options.pageRequestDelay,
    };
    let issuesAndPullRequests = await github.collections.getRepoIssues(
      this.authorize(AppPurpose.Data),
      parameters,
      cacheOptions
    );
    let issuesOnly = issuesAndPullRequests.filter((r) => !r.pull_request);
    const issues = common.createInstances<RepositoryIssue>(
      this,
      issueFromEntity,
      issuesOnly
    );
    return issues;
  }

  issue(issueNumber: number, optionalEntity?: any): RepositoryIssue {
    const issue = new RepositoryIssue(
      this,
      issueNumber,
      this._operations,
      this._getAuthorizationHeader,
      optionalEntity
    );
    return issue;
  }

  async createIssue(title: string, body: string): Promise<RepositoryIssue> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = {
      owner: this.organization.name,
      repo: this.name,
      title,
      body,
    };
    // Operations has issue write permissions
    const details = await operations.github.post(
      this.authorize(AppPurpose.Operations),
      'issues.create',
      parameters
    );
    const issueNumber = details.number as number;
    const issue = new RepositoryIssue(
      this,
      issueNumber,
      this._operations,
      this._getAuthorizationHeader,
      details
    );
    return issue;
  }
}

function issueFromEntity(entity) {
  // 'this' is bound for this function to be a private method
  const operations = this._operations;
  const permission = new RepositoryIssue(
    this,
    entity.number,
    operations,
    this._getAuthorizationHeader,
    entity
  );
  return permission;
}

function teamPermissionFromEntity(entity) {
  // 'this' is bound for this function to be a private method
  const operations = this._operations;
  const permission = new TeamPermission(this.organization, entity, operations);
  return permission;
}

function collaboratorPermissionFromEntity(entity) {
  // 'this' is bound for this function to be a private method
  const permission = new Collaborator(entity);
  return permission;
}
