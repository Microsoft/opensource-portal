//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Response, Request } from 'express';
import redis from 'redis';
import crypto from 'crypto';
import { Pool as PostgresPool } from 'pg';
import type { TelemetryClient } from 'applicationinsights';
import { AccessToken, AuthorizationCode } from 'simple-oauth2';
import { Session } from 'express-session';
import githubUsernameRegex from 'github-username-regex';

import { IndividualContext } from './user';
import { IApprovalProvider } from './entities/teamJoinApproval/approvalProvider';
import { GitHubIssueQuery, ICorporateLink, Operations } from './business';
import { ITokenProvider } from './entities/token';
import { IMailAddressProvider } from './lib/mailAddressProvider';
import { IRepositoryMetadataProvider } from './entities/repositoryMetadata/repositoryMetadataProvider';
import { ILocalExtensionKeyProvider } from './entities/localExtensionKey';
import { Organization } from './business';
import { IGraphProvider } from './lib/graphProvider';
import { RestLibrary } from './lib/github';
import { Team } from './business';
import { IRepositoryCacheProvider } from './entities/repositoryCache/repositoryCacheProvider';
import { IRepositoryCollaboratorCacheProvider } from './entities/repositoryCollaboratorCache/repositoryCollaboratorCacheProvider';
import { ITeamCacheProvider } from './entities/teamCache/teamCacheProvider';
import { ITeamMemberCacheProvider } from './entities/teamMemberCache/teamMemberCacheProvider';
import { IRepositoryTeamCacheProvider } from './entities/repositoryTeamCache/repositoryTeamCacheProvider';
import { IOrganizationMemberCacheProvider } from './entities/organizationMemberCache/organizationMemberCacheProvider';
import { AppPurpose } from './github';
import QueryCache from './business/queryCache';
import { IMailProvider } from './lib/mailProvider';
import { GitHubRepositoryPermission } from './entities/repositoryMetadata/repositoryMetadata';
import { IOrganizationSettingProvider } from './entities/organizationSettings/organizationSettingProvider';
import { ILinkProvider } from './lib/linkProviders';
import { IAuditLogRecordProvider } from './entities/auditLogRecord/auditLogRecordProvider';
import { ICacheHelper } from './lib/caching';
import { ICampaignHelper } from './lib/campaigns';
import { ICorporateContactProvider } from './lib/corporateContactProvider';
import { IQueueProcessor } from './lib/queues';
import { IReposApplication } from './app';
import { IUserSettingsProvider } from './entities/userSettings';

import appPackage from './package.json';
import BlobCache from './lib/caching/blob';
import { ICreateRepositoryApiResult } from './api/createRepo';
import { Repository } from './business/repository';
import { ICorporationAdministrationSection } from './interfaces';
import { GitHubSortDirection } from './lib/github/collections';
import { AxiosError } from 'axios';
const packageVariableName = 'static-react-package-name';

export function hasStaticReactClientApp() {
  const staticClientPackageName = appPackage[packageVariableName];
  if (process.env.ENABLE_REACT_CLIENT && staticClientPackageName) {
    return staticClientPackageName;
  }
}

export interface ICallback<T> {
  (error: IReposError, result?: T): void;
}

export type Json =
  | string
  | number
  | boolean
  | null
  | { [property: string]: Json }
  | Json[];

export interface IGetOwnerToken {
  (): string;
}

export enum RequestTeamMemberAddType {
  Member = 'member',
  Maintainer = 'maintainer',
}

export enum GitHubTeamPrivacy {
  Closed = 'closed',
  Secret = 'secret',
}

export interface IPurposefulGetAuthorizationHeader {
  (purpose: AppPurpose): Promise<IAuthorizationHeaderValue>;
}

export interface IAuthorizationHeaderValue {
  value: string;
  purpose: AppPurpose;
  source?: string;
}

export interface ICacheDefaultTimes {
  orgReposStaleSeconds: number;
  orgRepoTeamsStaleSeconds: number;
  orgRepoCollaboratorsStaleSeconds: number;
  orgRepoCollaboratorStaleSeconds: number;
  orgRepoDetailsStaleSeconds: number;
  orgTeamsStaleSeconds: number;
  orgTeamDetailsStaleSeconds: number;
  orgTeamsSlugLookupStaleSeconds: number;
  orgMembersStaleSeconds: number;
  teamMaintainersStaleSeconds: number;
  orgMembershipStaleSeconds: number;
  orgMembershipDirectStaleSeconds: number;
  crossOrgsReposStaleSecondsPerOrg: number;
  crossOrgsReposParallelCalls: number;
  crossOrgsMembersStaleSecondsPerOrg: number;
  crossOrgsMembersParallelCalls: number;
  corporateLinksStaleSeconds: number;
  repoBranchesStaleSeconds: number;
  repoPullsStaleSeconds: number;
  accountDetailStaleSeconds: number;
  teamDetailStaleSeconds: number;
  orgRepoWebhooksStaleSeconds: number;
  teamRepositoryPermissionStaleSeconds: number;
}

export enum CoreCapability {
  GitHubRestApi = 'GitHub REST API', // IOperationsGitHubRestLibrary
  DefaultCacheTimes = 'Default cache times', // IOperationsDefaultCacheTimes
  GitHubCentralOperations = 'GitHub central operations', // IOperationsCentralOperationsToken
  Urls = 'urls', // IOperationsUrls
  LockdownFeatureFlags = 'Lockdown feature flags', // IOperationsLockdownFeatureFlags
  Providers = 'Providers', // IOperationsProviders
  LegalEntities = 'Legal entities', // IOperationsLegalEntities
  ServiceAccounts = 'Service Accounts', // IOperationsServiceAccounts
  Links = 'Links', // IOperationsLinks
  Templates = 'Templates', // IOperationsTemplates
}

export interface IOperationsInstance {
  hasCapability(capability: CoreCapability): boolean;
  throwIfNotCompatible(capability: CoreCapability): void;
} // Messy but allows for optional use with a cast

export interface IOperationsLegalEntities {
  getDefaultLegalEntities(): string[];
}

export interface IOperationsTemplates {
  getDefaultRepositoryTemplateNames(): string[];
}

export interface IOperationsProviders {
  providers: IProviders;
}

export interface IOperationsLinks {
  getLinks(options?: any): Promise<ICorporateLink[]>;
  getLinkByThirdPartyId(thirdPartyId: string): Promise<ICorporateLink>;
  getLinkByThirdPartyUsername(username: string): Promise<ICorporateLink>;
  tryGetLink(login: string): Promise<ICorporateLink>;
  fireLinkEvent(value): Promise<void>;
  fireUnlinkEvent(value): Promise<void>;
}

export interface IOperationsServiceAccounts {
  isSystemAccountByUsername(login: string): boolean;
}

export interface IOperationsGitHubRestLibrary {
  github: RestLibrary;
}

export interface IOperationsDefaultCacheTimes {
  defaults: ICacheDefaultTimes;
}

export interface IOperationsUrls {
  baseUrl: string;
  absoluteBaseUrl: string;
}

export interface IOperationsLockdownFeatureFlags {
  allowUnauthorizedNewRepositoryLockdownSystemFeature(): boolean;
  allowUnauthorizedForkLockdownSystemFeature(): boolean;
  allowTransferLockdownSystemFeature(): boolean;
}

export interface IOperationsCentralOperationsToken {
  getCentralOperationsToken(): IPurposefulGetAuthorizationHeader; // IGetAuthorizationHeader ?;
  getAccountByUsername(
    username: string,
    options?: ICacheOptions
  ): Promise<Account>;
}

export function operationsIsCapable<T>(
  operations: IOperationsInstance,
  capability: CoreCapability
): operations is IOperationsInstance & T {
  return operations.hasCapability(capability);
}

export function operationsWithCapability<T>(
  operations: IOperationsInstance,
  capability: CoreCapability
): T & IOperationsInstance {
  if (operationsIsCapable<T>(operations, capability)) {
    return operations as T & IOperationsInstance;
  }
  return null;
}

export function throwIfNotCapable<T>(
  operations: IOperationsInstance,
  capability: CoreCapability
) {
  operations.throwIfNotCompatible(capability);
  return (operations as any) as T & IOperationsInstance;
}

export function throwIfNotGitHubCapable(operations: IOperationsInstance) {
  return throwIfNotCapable<IOperationsGitHubRestLibrary & IOperationsInstance>(
    operations,
    CoreCapability.GitHubRestApi
  );
}

export interface IGetAuthorizationHeader {
  (): Promise<IAuthorizationHeaderValue>;
}

export interface IFunctionPromise<T> {
  (): Promise<T>;
}

export interface PromiseResolve<T> {
  (resolve: T[]): void;
}

export interface PromiseReject {
  (reject?: any): void;
}

export interface ICacheOptions {
  backgroundRefresh?: any | null | undefined;
  maxAgeSeconds?: number | null | undefined;
}

export interface IPagedCacheOptions extends ICacheOptions {
  pageRequestDelay?: number | null | undefined; // FUTURE: could be a function, too
}

export enum GetIssuesSort {
  Created = 'created',
  Updated = 'updated',
  Comments = 'comments',
}

export interface IRepositoryGetIssuesOptions extends IPagedCacheOptions {
  since?: Date;
  direction?: GitHubSortDirection;
  sort?: GetIssuesSort;
  labels?: string;
  mentioned?: string;
  creator?: string;
  assignee?: string; // user | 'none' | '*'
  state?: GitHubIssueQuery;
  milestone?: number | string; // '*'
}

export interface IPagedCrossOrganizationCacheOptions
  extends IPagedCacheOptions {
  individualMaxAgeSeconds?: number | null | undefined;
  individualRequestDelay?: number | null | undefined; // FUTURE: could be a function, too
}

export interface ILocalCacheOptions extends ICacheOptions {
  localMaxAgeSeconds?: number;
}

export interface ICacheOptionsPageLimiter extends ICacheOptions {
  pageLimit?: number;
}

export interface IMapPlusMetaCost extends Map<any, any> {
  headers?: any;
  cost?: IReposRestRedisCacheCost;
}

export interface IReposRestRedisCacheCost {
  github: {
    cacheHits: number;
    remainingApiTokens: string;
    restApiCalls: number;
    usedApiTokens: number;
  };
  local: {
    cacheHits: number;
    cacheMisses: number;
  };
  redis: {
    cacheHits: number;
    cacheMisses: number;
    expireCalls: number;
    getCalls: number;
    setCalls: number;
  };
}

export interface IDictionary<TValue> {
  [id: string]: TValue;
}

export const NoCacheNoBackground = {
  backgroundRefresh: false,
  maxAgeSeconds: -1,
};

export interface IProviders {
  app: IReposApplication;
  applicationProfile: IApplicationProfile;
  authorizationCodeClient?: AuthorizationCode;
  corporateAdministrationProfile?: ICorporationAdministrationSection;
  corporateViews?: any;
  approvalProvider?: IApprovalProvider;
  auditLogRecordProvider?: IAuditLogRecordProvider;
  basedir?: string;
  campaignStateProvider?: ICampaignHelper;
  campaign?: any; // campaign redirection route, poor variable name
  corporateContactProvider?: ICorporateContactProvider;
  config?: any;
  customizedNewRepositoryLogic?: ICustomizedNewRepositoryLogic;
  customizedTeamPermissionsWebhookLogic?: ICustomizedTeamPermissionsWebhookLogic;
  diagnosticsDrop?: BlobCache;
  healthCheck?: any;
  keyEncryptionKeyResolver?: any;
  github?: RestLibrary;
  graphProvider?: IGraphProvider;
  insights?: TelemetryClient;
  linkProvider?: ILinkProvider;
  localExtensionKeyProvider?: ILocalExtensionKeyProvider;
  mailAddressProvider?: IMailAddressProvider;
  mailProvider?: IMailProvider;
  operations?: Operations;
  organizationMemberCacheProvider?: IOrganizationMemberCacheProvider;
  organizationSettingsProvider?: IOrganizationSettingProvider;
  postgresPool?: PostgresPool;
  queryCache?: QueryCache;
  webhookQueueProcessor?: IQueueProcessor;
  sessionRedisClient?: redis.RedisClient;
  cacheProvider?: ICacheHelper;
  repositoryCacheProvider?: IRepositoryCacheProvider;
  repositoryCollaboratorCacheProvider?: IRepositoryCollaboratorCacheProvider;
  repositoryMetadataProvider?: IRepositoryMetadataProvider;
  repositoryTeamCacheProvider?: IRepositoryTeamCacheProvider;
  session?: any;
  teamCacheProvider?: ITeamCacheProvider;
  teamMemberCacheProvider?: ITeamMemberCacheProvider;
  userSettingsProvider?: IUserSettingsProvider;
  tokenProvider?: ITokenProvider;
  viewServices?: any;
  //redis?: RedisHelper;
  //redisClient?: redis.RedisClient;
}

export enum UserAlertType {
  Success = 'success',
  Warning = 'warning',
  Danger = 'danger',
}

export interface IApplicationProfile {
  applicationName: string;
  customErrorHandlerRender?: (
    errorView: any,
    err: Error,
    req: any,
    res: any,
    next: any
  ) => Promise<void>;
  customRoutes?: () => Promise<void>;
  logDependencies: boolean;
  serveClientAssets: boolean;
  serveStaticAssets: boolean;
  validate?: () => Promise<void>;
  startup?: (providers: IProviders) => Promise<void>;
  sessions: boolean;
  webServer: boolean;
}

export interface RedisOptions {
  auth_pass?: string;
  detect_buffers: boolean;
  tls?: {
    servername: string;
  };
}

export interface InnerError extends Error {
  inner?: Error;
}

export interface IReposError extends Error {
  skipLog?: boolean;
  status?: any; // status?: number;
  code?: any; // not sure this is used any longer by libraries
  originalUrl?: any;
  detailed?: any;
  redirect?: string;
  skipOops?: boolean;
  fancyLink?: {
    link: string;
    title: string;
  };
  fancySecondaryLink?: {
    link: string;
    title: string;
  };
  innerError?: IReposError;
}

export interface IReposAppContext {
  section?: string;
  pivotDirectlyToOtherOrg?: string;
  releaseTab?: boolean;
  organization?: Organization;
}

export interface IReposAppWithTeam extends ReposAppRequest {
  teamPermissions?: any;
  team2?: Team;
  teamUrl: string;
}

export enum LocalApiRepoAction {
  Delete = 'delete',
  Archive = 'archive',
}

export interface ReposAppRequest extends Request {
  // passport
  isAuthenticated(): boolean;
  user: any;

  // our extensions
  insights?: any;
  reposContext?: IReposAppContext;
  currentOrganizationMemberships?: any; // needs a redesign
  teamsPagerMode?: string;
  reposPagerMode?: string;
  link?: any; // not sure when this is set
  organization?: Organization;
  correlationId?: string;
  scrubbedUrl?: string;

  // FUTURE:
  apiContext: IndividualContext;
  individualContext: IndividualContext;
  oauthAccessToken: AccessToken;
}

export function getProviders(req: ReposAppRequest) {
  return req.app.settings.providers as IProviders;
}

export interface IReposAppResponse extends Response {}

export interface IReposRequestWithOrganization extends ReposAppRequest {
  organization?: any;
}

export interface IRequestTeams extends ReposAppRequest {
  team2?: any;
  teamUrl?: any;
}

export interface RequestWithSystemwidePermissions extends ReposAppRequest {
  systemWidePermissions?: any;
}

export interface IResponseForSettingsPersonalAccessTokens extends Response {
  newKey?: string;
}

interface ITooManyLinksError extends Error {
  links?: any;
  tooManyLinks?: boolean;
}

interface IExistingIdentityError extends Error {
  anotherAccount?: boolean;
  link?: any;
  skipLog?: boolean;
}

function tooManyLinksError(self, userLinks, callback) {
  const tooManyLinksError: ITooManyLinksError = new Error(
    `This account has ${userLinks.length} linked GitHub accounts.`
  );
  tooManyLinksError.links = userLinks;
  tooManyLinksError.tooManyLinks = true;
  return callback(tooManyLinksError, self);
}

function existingGitHubIdentityError(self, link, requestUser, callback) {
  const endUser = requestUser.azure.displayName || requestUser.azure.username;
  const anotherGitHubAccountError: IExistingIdentityError = new Error(
    `${endUser}, there is a different GitHub account linked to your corporate identity.`
  );
  anotherGitHubAccountError.anotherAccount = true;
  anotherGitHubAccountError.link = link;
  anotherGitHubAccountError.skipLog = true;
  return callback(anotherGitHubAccountError, self);
}

export function SettleToStateValue<T>(
  promise: Promise<T>
): Promise<ISettledValue<T>> {
  return promise.then(
    (value) => {
      return { value, state: SettledState.Fulfilled };
    },
    (reason) => {
      return { reason, state: SettledState.Rejected };
    }
  );
}

export function permissionsObjectToValue(
  permissions
): GitHubRepositoryPermission {
  if (permissions.admin === true) {
    return GitHubRepositoryPermission.Admin;
  } else if (permissions.push === true) {
    return GitHubRepositoryPermission.Push;
  } else if (permissions.triage === true) {
    return GitHubRepositoryPermission.Triage;
  } else if (permissions.maintain === true) {
    return GitHubRepositoryPermission.Maintain;
  } else if (permissions.pull === true) {
    return GitHubRepositoryPermission.Pull;
  }
  throw new Error(
    `Unsupported GitHubRepositoryPermission value inside permissions`
  );
}

export function isPermissionBetterThan(
  currentBest: GitHubRepositoryPermission,
  newConsideration: GitHubRepositoryPermission
) {
  switch (newConsideration) {
    case GitHubRepositoryPermission.Admin:
      return true;
    case GitHubRepositoryPermission.Maintain:
      if (currentBest !== GitHubRepositoryPermission.Admin) {
        return true;
      }
      break;
    case GitHubRepositoryPermission.Push:
      if (currentBest !== GitHubRepositoryPermission.Admin) {
        return true;
      }
      break;
    case GitHubRepositoryPermission.Pull:
      if (
        currentBest === null ||
        currentBest === GitHubRepositoryPermission.None
      ) {
        return true;
      }
      break;
    case GitHubRepositoryPermission.Triage:
      // not really great
      break;
    default:
      throw new Error(`Invalid permission type ${newConsideration}`);
  }
  return false;
}

export function MassagePermissionsToGitHubRepositoryPermission(
  value: string
): GitHubRepositoryPermission {
  // collaborator level APIs return a more generic read/write value, lead to some bad caches in the past...
  // TODO: support new collaboration values as they come online for Enterprise Cloud!
  switch (value) {
    case 'write':
    case 'push':
      return GitHubRepositoryPermission.Push;
    case 'admin':
      return GitHubRepositoryPermission.Admin;
    case 'triage':
      return GitHubRepositoryPermission.Triage;
    case 'maintain':
      return GitHubRepositoryPermission.Maintain;
    case 'pull':
    case 'read':
      return GitHubRepositoryPermission.Pull;
    default:
      throw new Error(
        `Invalid ${value} GitHub repository permission [massagePermissionsToGitHubRepositoryPermission]`
      );
  }
}

export interface ISettledValue<T> {
  reason?: any;
  value?: T;
  state: SettledState;
}

export enum SettledState {
  Fulfilled = 'fulfilled',
  Rejected = 'rejected',
}

export class CreateError {
  static CreateStatusCodeError(code: number, message?: string): Error {
    const error = new Error(message);
    error['status'] = code;
    return error;
  }

  static NotFound(message: string, innerError?: Error): Error {
    return ErrorHelper.SetInnerError(
      CreateError.CreateStatusCodeError(404, message),
      innerError
    );
  }

  static ParameterRequired(
    parameterName: string,
    optionalDetails?: string
  ): Error {
    const msg = `${parameterName} required`;
    return CreateError.CreateStatusCodeError(
      400,
      optionalDetails ? `${msg}: ${optionalDetails}` : msg
    );
  }

  static InvalidParameters(message: string, innerError?: Error): Error {
    return ErrorHelper.SetInnerError(
      CreateError.CreateStatusCodeError(400, message),
      innerError
    );
  }

  static NotAuthenticated(message: string): Error {
    return CreateError.CreateStatusCodeError(401, message);
  }

  static NotAuthorized(message: string): Error {
    return CreateError.CreateStatusCodeError(403, message);
  }

  static ServerError(message: string, innerError?: Error): Error {
    return ErrorHelper.SetInnerError(
      CreateError.CreateStatusCodeError(500, message),
      innerError
    );
  }
}

export class ErrorHelper {
  static EnsureHasStatus(error: Error, code: number): Error {
    if (!error['status']) {
      error['status'] = code;
    }
    return error;
  }

  public static WrapError(innerError: Error, message: string): Error {
    const err = new Error(message);
    err['innerError'] = innerError;
    return err;
  }

  public static SetInnerError(error: Error, innerError: Error) {
    if (error && innerError) {
      error['innerError'] = innerError;
    }
    return error;
  }

  public static HasStatus(error: Error): boolean {
    return error && error['status'];
  }

  public static IsNotFound(error: Error): boolean {
    const statusNumber = ErrorHelper.GetStatus(error);
    return statusNumber && statusNumber === 404;
  }

  public static IsConflict(error: Error): boolean {
    const statusNumber = ErrorHelper.GetStatus(error);
    if (statusNumber && statusNumber === 409) {
      return true;
    }
    // would be nice to be able to get rid of this clause someday
    if (error.message && error.message.includes('already exists')) {
      return true;
    }
    return false;
  }

  public static NotImplemented() {
    return new Error('Not implemented');
  }

  public static GetStatus(error: Error): number {
    const asAny = error as any;
    if (asAny?.isAxiosError === true) {
      const axiosError = asAny as AxiosError;
      if (axiosError?.response?.status) {
        return axiosError.response.status;
      }
    }
    if (asAny?.status) {
      const status = asAny.status;
      const type = typeof status;
      if (type === 'number') {
        return status;
      } else if (type === 'string') {
        return Number(status);
      } else {
        console.warn(`Unsupported error.status type: ${type}`);
        return null;
      }
    }
    return null;
  }
}

export function setImmediateAsync(f: IFunctionPromise<void>): void {
  const safeCall = () => {
    try {
      f().catch((error) => {
        console.warn(`setImmediateAsync caught error: ${error}`);
      });
    } catch (ignoredFailure) {
      console.warn(`setImmediateAsync call error: ${ignoredFailure}`);
    }
  };
  setImmediate(safeCall.bind(null));
}

export function stripDistFolderName(dirname: string) {
  // This is a hacky backup for init failure scenarios where the dirname may
  // not actually point at the app root.
  if (dirname.endsWith('dist')) {
    dirname = dirname.replace('\\dist', '');
    dirname = dirname.replace('/dist', '');
  }
  return dirname;
}

export interface IUserAlert {
  message: string;
  title: string;
  context: UserAlertType;
  optionalLink: string;
  optionalCaption: string;
}

interface IAppSessionProperties extends Session {
  enableMultipleAccounts: boolean;
  selectedGithubId: string;
  passport: any;
  id: string;
  alerts?: IUserAlert[];
  referer: string;
}

export function sha256(str: string) {
  const hash = crypto.createHash('sha256').update(str).digest('base64');
  return hash;
}

export interface IAppSession extends IAppSessionProperties {}

export interface ICustomizedNewRepositoryLogic {
  createContext(req: any): INewRepositoryContext;
  getAdditionalTelemetryProperties(
    context: INewRepositoryContext
  ): IDictionary<string>;
  validateRequest(context: INewRepositoryContext, req: any): Promise<void>;
  stripRequestBody(context: INewRepositoryContext, body: any): void;
  afterRepositoryCreated(
    context: INewRepositoryContext,
    corporateId: string,
    success: ICreateRepositoryApiResult
  ): Promise<void>;
  shouldNotifyManager(
    context: INewRepositoryContext,
    corporateId: string
  ): boolean;
  getNewMailViewProperties(
    context: INewRepositoryContext,
    repository: Repository
  ): Promise<ICustomizedNewRepoProperties>;
}

export interface ICustomizedNewRepoProperties {
  viewProperties: any;
  to?: string[];
  cc?: string[];
  bcc?: string[];
}

export interface ICustomizedTeamPermissionsWebhookLogic {
  shouldSkipEnforcement(repository: Repository): Promise<boolean>;
}

export interface INewRepositoryContext {
  isCustomContext: boolean;
}

export function validateGitHubLogin(username: string) {
  // There are some legitimate usernames at GitHub that have a dash
  // in them. While GitHub no longer allows this for new accounts,
  // they are grandfathered in.
  if (!githubUsernameRegex.test(username) && !username.endsWith('-')) {
    console.warn(`Invalid GitHub username format: ${username}`);
    // throw new Error(`Invalid GitHub username format: ${username}`);
  }
  return username;
}
