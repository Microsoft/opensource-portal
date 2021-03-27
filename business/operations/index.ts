//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import axios from 'axios';
import throat from 'throat';

import {
  ICacheOptions,
  IMapPlusMetaCost,
  IPagedCrossOrganizationCacheOptions,
  IGetAuthorizationHeader,
  IPurposefulGetAuthorizationHeader,
  IAuthorizationHeaderValue,
  IDictionary,
  CreateError,
  ErrorHelper,
  setImmediateAsync,
  IOperationsGitHubRestLibrary,
  ICacheDefaultTimes,
  IOperationsDefaultCacheTimes,
  IOperationsCentralOperationsToken,
  IOperationsLegalEntities,
  CoreCapability,
  IOperationsServiceAccounts,
  IOperationsTemplates,
  IOperationsLinks,
  IOperationsLockdownFeatureFlags,
} from '../../transitional';

import { Account } from '../account';
import { GraphManager } from '../graphManager';
import { Organization } from '../organization';
import { GitHubTokenManager } from '../../github/tokenManager';

import RenderHtmlMail from '../../lib/emailRender';

import { wrapError, sortByCaseInsensitive } from '../../utils';
import { ICorporateLink } from '../corporateLink';
import { Repository } from '../repository';
import { RestLibrary } from '../../lib/github';
import { Team, ICrossOrganizationTeamMembership } from '../team';
import { GitHubAppAuthenticationType } from '../../github';
import { OrganizationSetting } from '../../entities/organizationSettings/organizationSetting';
import { OrganizationSettingProvider } from '../../entities/organizationSettings/organizationSettingProvider';
import { IMail } from '../../lib/mailProvider';
import { ILinkProvider } from '../../lib/linkProviders';
import { IGraphEntryWithManager } from '../../lib/graphProvider';
import { ICacheHelper } from '../../lib/caching';
import { createPortalSudoInstance, IPortalSudo } from '../../features';
import { IOperationsCoreOptions, OperationsCore } from './core';

import { linkAccounts as linkAccountsMethod } from './link';
import { sendTerminatedAccountMail as sendTerminatedAccountMailMethod } from './unlinkMail';

export * from './core';

const throwIfOrganizationIdsMissing = true;

const SecondsBetweenOrganizationSettingUpdatesCheck = 60 * 2; // every 2 minutes, check for dynamic app updates
let DynamicRestartCheckHandle = null;

const ParallelLinkLookup = 4;

export const RedisPrefixManagerInfoCache = 'employeewithmanager:';

const defaultGitHubPageSize = 100;

export interface IUnlinkMailStatus {
  to: string[];
  bcc: string[];
  receipt: string;
}

export enum SupportedLinkType {
  User = 'user',
  ServiceAccount = 'serviceAccount',
}

export interface ISupportedLinkTypeOutcome {
  type: SupportedLinkType;
  graphEntry: IGraphEntryWithManager;
}

export enum UnlinkPurpose {
  Unknown = 'unknown',
  Termination = 'termination', // no longer listed as an employee
  Self = 'self', // the user self-service unlink themselves
  Operations = 'operations', // operational support
  Deleted = 'deleted', // the GitHub account has been deleted or does not exist
}

export enum LinkOperationSource {
  Portal = 'portal',
  Api = 'api',
}

export interface ICreateLinkOptions {
  link: ICorporateLink;
  operationSource: LinkOperationSource;
  skipCorporateValidation?: boolean;
  skipGitHubValidation?: boolean;
  skipSendingMail?: boolean;
  eventProperties?: IDictionary<string>;
  correlationId?: string;
}

export interface ICreatedLinkOutcome {
  linkId: string;
  resourceLink?: string;
}

export interface ICrossOrganizationMembershipBasics {
  id: string;
  login: string;
  avatar_url: string;
}

export interface ICrossOrganizationMembershipByOrganization {
  id: number; // ?
  orgs?: ICrossOrganizationMembershipBasics[]; // TODO: WARNING: This typing is incorrect. The object properties are the org name.
}

interface IPromisedLinks {
  headers: {
    type: 'links';
  };
  data: ICorporateLink[];
}

export interface ICrossOrganizationMembersResult
  extends Map<number, ICrossOrganizationMembershipByOrganization> {}

export interface ICachedEmployeeInformation {
  id: string;
  displayName: string;
  userPrincipalName: string;
  managerId: string;
  managerDisplayName: string;
  managerMail: string;
}

export interface IOperationsOptions extends IOperationsCoreOptions {
  // cacheProvider: ICacheHelper;
  // config: any;
  github: RestLibrary;
  // insights: TelemetryClient;
  // linkProvider: ILinkProvider;
  // mailAddressProvider: IMailAddressProvider;
  // mailProvider: IMailProvider;
}

export class Operations
  extends OperationsCore
  implements
    IOperationsLegalEntities,
    IOperationsServiceAccounts,
    IOperationsTemplates,
    IOperationsLinks,
    IOperationsLockdownFeatureFlags {
  private _cache: ICacheHelper;
  private _graphManager: GraphManager;
  private _organizationNames: string[];
  private _organizations: Map<string, Organization>;
  private _uncontrolledOrganizations: Map<string, Organization>;
  private _organizationOriginalNames: any;
  private _organizationNamesWithAuthorizationHeaders: Map<
    string,
    IPurposefulGetAuthorizationHeader
  >;
  private _defaultPageSize: number;
  private _organizationIds: Map<number, Organization>;
  private _dynamicOrganizationSettings: OrganizationSetting[];
  private _dynamicOrganizationIds: Set<number>;
  private _portalSudo: IPortalSudo;
  private _tokenManager: GitHubTokenManager;

  get graphManager(): GraphManager {
    return this._graphManager;
  }

  get defaultPageSize(): number {
    return this._defaultPageSize;
  }

  constructor(options: IOperationsOptions) {
    super(options);

    this.addCapability(CoreCapability.Providers);
    this.addCapability(CoreCapability.LegalEntities);
    this.addCapability(CoreCapability.ServiceAccounts);
    this.addCapability(CoreCapability.Templates);
    this.addCapability(CoreCapability.Links);
    this.addCapability(CoreCapability.LockdownFeatureFlags);
    this.addCapability(CoreCapability.GitHubCentralOperations);

    const providers = options.providers;
    const config = providers.config;
    this._cache = providers.cacheProvider;
    this._baseUrl = '/';
    this._graphManager = new GraphManager(this);
    this._uncontrolledOrganizations = new Map();
    this._defaultPageSize =
      this.config &&
      this.config.github &&
      this.config.github.api &&
      this.config.github.api.defaultPageSize
        ? this.config.github.api.defaultPageSize
        : defaultGitHubPageSize;
    const hasModernGitHubApps = config.github?.app;
    this._tokenManager = new GitHubTokenManager({
      customerFacingApp: hasModernGitHubApps ? config.github.app.ui : null,
      operationsApp: hasModernGitHubApps ? config.github.app.operations : null,
      dataApp: hasModernGitHubApps ? config.github.app.data : null,
      backgroundJobs: hasModernGitHubApps ? config.github.app.jobs : null,
      updatesApp: hasModernGitHubApps ? config.github.app.updates : null,
      app: this.providers.app,
    });
    this._dynamicOrganizationIds = new Set();
    this._dynamicOrganizationSettings = [];
  }

  protected get tokenManager() {
    return this._tokenManager;
  }

  async initialize() {
    await super.initialize();
    const hasModernGitHubApps = this.config.github && this.config.github.app;
    // const hasConfiguredOrganizations = this.config.github.organizations && this.config.github.organizations.length;
    const organizationSettingsProvider = this.providers
      .organizationSettingsProvider;
    if (hasModernGitHubApps && organizationSettingsProvider) {
      const dynamicOrganizations = (
        await organizationSettingsProvider.queryAllOrganizations()
      ).filter(
        (dynamicOrg) =>
          dynamicOrg.active === true && !dynamicOrg.hasFeature('ignore')
      );
      this._dynamicOrganizationSettings = dynamicOrganizations;
      this._dynamicOrganizationIds = new Set(
        dynamicOrganizations.map((org) => Number(org.organizationId))
      );
    }
    if (this._dynamicOrganizationSettings && organizationSettingsProvider) {
      DynamicRestartCheckHandle = setInterval(
        restartAfterDynamicConfigurationUpdate.bind(
          null,
          10,
          120,
          this.initialized,
          organizationSettingsProvider
        ),
        1000 * SecondsBetweenOrganizationSettingUpdatesCheck
      );
    }
    if (throwIfOrganizationIdsMissing) {
      this.getOrganizationIds();
    }
    this._portalSudo = createPortalSudoInstance(this.providers);
    return this;
  }

  get organizationNames(): string[] {
    if (!this._organizationNames) {
      const names = [];
      const processed = new Set<string>();
      for (const dynamic of this._dynamicOrganizationSettings) {
        const lowercase = dynamic.organizationName.toLowerCase();
        processed.add(lowercase);
        names.push(lowercase);
      }
      for (let i = 0; i < this.config.github.organizations.length; i++) {
        const lowercase = this.config.github.organizations[
          i
        ].name.toLowerCase();
        if (!processed.has(lowercase)) {
          names.push(lowercase);
          processed.add(lowercase);
        }
      }
      this._organizationNames = names.sort(sortByCaseInsensitive);
    }
    return this._organizationNames;
  }

  getOrganizationIds(): number[] {
    if (!this._organizationIds) {
      const organizations = this.organizations;
      this._organizationIds = new Map();
      this._dynamicOrganizationSettings.map((entry) => {
        if (entry.active) {
          const org = this.getOrganization(
            entry.organizationName.toLowerCase()
          );
          this._organizationIds.set(Number(entry.organizationId), org);
        }
      });
      // This check only runs on _static_ configuration entries, since adopted
      // GitHub App organizations must always have an organization ID.
      for (let i = 0; i < this.config.github.organizations.length; i++) {
        const organizationConfiguration = this.config.github.organizations[i];
        const organization = organizations.get(
          organizationConfiguration.name.toLowerCase()
        );
        if (!organization) {
          throw new Error(
            `Missing organization configuration ${organizationConfiguration.name}`
          );
        }
        if (!organizationConfiguration.id) {
          if (throwIfOrganizationIdsMissing) {
            throw new Error(
              `Organization ${organization.name} is not configured with an 'id' which can lead to issues if the organization is renamed. throwIfOrganizationIdsMissing is true: id is required`
            );
          } else {
            console.warn(
              `Organization ${organization.name} is not configured with an 'id' which can lead to issues if the organization is renamed.`
            );
          }
        } else if (!this._organizationIds.has(organizationConfiguration.id)) {
          this._organizationIds.set(organizationConfiguration.id, organization);
        }
      }
    }
    return Array.from(this._organizationIds.keys());
  }

  private createOrganization(
    name: string,
    settings: OrganizationSetting,
    centralOperationsFallbackToken: string,
    appAuthenticationType: GitHubAppAuthenticationType
  ): Organization {
    name = name.toLowerCase();
    let ownerToken = null;
    if (!settings) {
      let staticSettings = null;
      const group = this.config.github.organizations;
      for (let i = 0; i < group.length; i++) {
        if (group[i].name && group[i].name.toLowerCase() === name) {
          const staticOrganizationSettings = group[i];
          if (staticOrganizationSettings.ownerToken) {
            ownerToken = staticOrganizationSettings.ownerToken;
          }
          staticSettings = staticOrganizationSettings;
          break;
        }
      }
      try {
        settings = OrganizationSetting.CreateFromStaticSettings(staticSettings);
        settings.active = true;
      } catch (translateStaticSettingsError) {
        throw new Error(
          `This application is not able to translate the static configuration for the ${name} organization. Specific error: ${translateStaticSettingsError.message}`
        );
      }
    }
    if (!settings) {
      throw new Error(
        `This application is not configured for the ${name} organization`
      );
    }
    const hasDynamicSettings =
      this._dynamicOrganizationIds &&
      settings.organizationId &&
      this._dynamicOrganizationIds.has(Number(settings.organizationId));
    return new Organization(
      this,
      name,
      settings,
      this.getAuthorizationHeader.bind(
        this,
        name,
        settings,
        ownerToken,
        centralOperationsFallbackToken,
        appAuthenticationType
      ),
      this.getAuthorizationHeader.bind(
        this,
        name,
        settings,
        ownerToken,
        centralOperationsFallbackToken,
        GitHubAppAuthenticationType.ForceSpecificInstallation
      ),
      hasDynamicSettings
    );
  }

  get organizations() {
    if (!this._organizations) {
      const organizations = new Map<string, Organization>();
      const names = this.organizationNames;
      const centralOperationsToken = this.config.github.operations
        .centralOperationsToken;
      for (let i = 0; i < names.length; i++) {
        const name = names[i];
        let dynamicSettings: OrganizationSetting = null;
        this._dynamicOrganizationSettings.map((dos) => {
          if (
            dos.active &&
            dos.organizationName.toLowerCase() === name.toLowerCase()
          ) {
            dynamicSettings = dos;
          }
        });
        const organization = this.createOrganization(
          name,
          dynamicSettings,
          centralOperationsToken,
          GitHubAppAuthenticationType.BestAvailable
        );
        organizations.set(name, organization);
      }
      this._organizations = organizations;
    }
    return this._organizations;
  }

  private getAlternateOrganization(name: string, alternativeType) {
    // An 'alternate' organization is one whose static settings come from a
    // different location within the github.organizations config file.
    const lowercase = name.toLowerCase();
    const list = this.config.github.organizations[alternativeType];
    if (list) {
      for (let i = 0; i < list.length; i++) {
        const settings = list[i];
        if (
          settings &&
          settings.name &&
          settings.name.toLowerCase() === lowercase
        ) {
          const centralOperationsToken = this.config.github.operations
            .centralOperationsToken;
          return this.createOrganization(
            lowercase,
            settings,
            centralOperationsToken,
            GitHubAppAuthenticationType.BestAvailable
          );
        }
      }
    }
  }

  getOnboardingOrganization(name: string) {
    // Specialized method to retrieve a new organization via the onboarding configuration collection, if any
    const value = this.getAlternateOrganization(name, 'onboarding');
    if (value) {
      return value;
    }
    throw new Error(
      `No onboarding organization settings configured for the ${name} organization`
    );
  }

  getUnconfiguredOrganization(settings: OrganizationSetting): Organization {
    return this.createOrganization(
      settings.organizationName.toLowerCase(),
      settings,
      null,
      GitHubAppAuthenticationType.BestAvailable
    );
  }

  getUncontrolledOrganization(
    organizationName: string,
    organizationId?: number
  ): Organization {
    organizationName = organizationName.toLowerCase();
    const officialOrganization = this.organizations.get(organizationName);
    if (officialOrganization) {
      return officialOrganization;
    }
    if (this._uncontrolledOrganizations.has(organizationName)) {
      return this._uncontrolledOrganizations.get(organizationName);
    }
    const emptySettings = new OrganizationSetting();
    emptySettings.operationsNotes = `Uncontrolled Organization - ${organizationName}`;
    const centralOperationsToken = this.config.github.operations
      .centralOperationsToken;
    const org = this.createOrganization(
      organizationName,
      emptySettings,
      centralOperationsToken,
      GitHubAppAuthenticationType.ForceSpecificInstallation
    );
    this._uncontrolledOrganizations.set(organizationName, org);
    org.uncontrolled = true;
    return org;
  }

  isIgnoredOrganization(name: string): boolean {
    const value =
      this.getAlternateOrganization(name, 'onboarding') ||
      this.getAlternateOrganization(name, 'ignore');
    return !!value;
  }

  isManagedOrganization(name: string) {
    try {
      this.getOrganization(name.toLowerCase());
      return true;
    } catch (unmanaged) {
      return this.isIgnoredOrganization(name);
    }
  }

  getOrganizations(organizationList?: string[]): Organization[] {
    if (!organizationList) {
      return Array.from(this.organizations.values());
    }
    const references = [];
    organizationList.forEach((orgName) => {
      const organization = this.getOrganization(orgName);
      references.push(organization);
    });
    return references;
  }

  getPrimaryOrganizationName(): string {
    const id =
      this.config.github &&
      this.config.github.operations &&
      this.config.github.operations.primaryOrganizationId
        ? this.config.github.operations.primaryOrganizationId
        : null;
    if (id) {
      return this.getOrganizationById(Number(id)).name;
    }
    return this.getOrganizationOriginalNames()[0];
  }

  getOrganizationOriginalNames(): string[] {
    if (!this._organizationOriginalNames) {
      const names: string[] = [];
      const visited = new Set<string>();
      for (const entry of this._dynamicOrganizationSettings) {
        if (entry.active) {
          names.push(entry.organizationName);
          const lowercase = entry.organizationName.toLowerCase();
          visited.add(lowercase);
        }
      }
      for (let i = 0; i < this.config.github.organizations.length; i++) {
        const original = this.config.github.organizations[i].name;
        const lowercase = original.toLowerCase();
        if (!visited.has(lowercase)) {
          names.push(original);
          visited.add(lowercase);
        }
      }
      this._organizationOriginalNames = names.sort(sortByCaseInsensitive);
    }
    return this._organizationOriginalNames;
  }

  translateOrganizationNamesFromLowercase(object) {
    const orgs = this.getOrganizationOriginalNames();
    orgs.forEach((name) => {
      const lc = name.toLowerCase();
      if (name !== lc && object[lc] !== undefined) {
        object[name] = object[lc];
        delete object[lc];
      }
    });
    return object;
  }

  get organizationNamesWithAuthorizationHeaders() {
    if (!this._organizationNamesWithAuthorizationHeaders) {
      const tokens = new Map<string, IPurposefulGetAuthorizationHeader>();
      const visited = new Set<string>();
      for (const entry of this._dynamicOrganizationSettings) {
        const lowercase = entry.organizationName.toLowerCase();
        if (entry.active && !visited.has(lowercase)) {
          visited.add(lowercase);
          const orgInstance = this.getOrganization(lowercase);
          const token = orgInstance.getAuthorizationHeader();
          tokens.set(lowercase, token);
        }
      }
      for (let i = 0; i < this.config.github.organizations.length; i++) {
        const name = this.config.github.organizations[i].name.toLowerCase();
        if (visited.has(name)) {
          continue;
        }
        visited.add(name);
        const orgInstance = this.getOrganization(name);
        const token = orgInstance.getAuthorizationHeader();
        tokens.set(name, token);
      }
      this._organizationNamesWithAuthorizationHeaders = tokens;
    }
    return this._organizationNamesWithAuthorizationHeaders;
  }

  async getCachedEmployeeManagementInformation(
    corporateId: string
  ): Promise<ICachedEmployeeInformation> {
    const key = `${RedisPrefixManagerInfoCache}${corporateId}`;
    const currentManagerIfAny = await this._cache.getObjectCompressed(key);
    return currentManagerIfAny as ICachedEmployeeInformation;
  }

  linkAccounts(options: ICreateLinkOptions): Promise<ICreatedLinkOutcome> {
    return linkAccountsMethod(this, options);
  }

  async validateCorporateAccountCanLink(
    corporateId: string
  ): Promise<ISupportedLinkTypeOutcome> {
    const { graphProvider } = this.providers;
    const graphEntry = await graphProvider.getUserAndManagerById(corporateId);
    // NOTE: This assumption, that a user without a manager must be a Service Account,
    // is a bit of a hack. It means that the CEO will be flagged as a service account if
    // they find the time to use this app. This code prioritizes the more common scenario,
    // that a user without an assigned manager in the directory is a Service Account.
    if (graphEntry && !graphEntry.manager) {
      return { type: SupportedLinkType.ServiceAccount, graphEntry };
    }
    return { type: SupportedLinkType.User, graphEntry };
  }

  async terminateLinkAndMemberships(
    thirdPartyId,
    options?: any
  ): Promise<string[]> {
    const insights = this.insights;

    options = options || {};
    let history: string[] = [];
    const continueOnError = options.continueOnError || false;
    let errors = 0;

    const account: Account = this.getAccount(thirdPartyId);
    const reason = options.reason || 'Automated processPendingUnlink operation';
    const purpose = (options.purpose as UnlinkPurpose) || UnlinkPurpose.Unknown;

    try {
      // Uses an ID-based lookup on GitHub in case the user was renamed.
      // Also retrieves the link into memory in the account instance.
      await account.getDetailsAndDirectLink();
    } catch (noDirectDetails) {
      ++errors;
      insights?.trackException({ exception: noDirectDetails });
      // not a fatal error in this method however
      history.push(noDirectDetails.toString());
    }

    insights?.trackEvent({
      name: 'UserUnlinkStart',
      properties: {
        id: account.id,
        login: account.login,
        reason: reason,
        purpose,
        continueOnError: continueOnError
          ? 'continue on errors'
          : 'halt on errors',
      },
    });

    // GitHub memberships
    try {
      const removal = await account.removeManagedOrganizationMemberships();
      history.push(...removal.history);
      if (removal.error) {
        throw removal.error; // unclear if this is actually ideal
      }
    } catch (removeOrganizationsError) {
      ++errors;
      // If a removal error occurs, do not remove the link and throw and error
      // so that the link data and information is still present until the issue
      // can be cleared
      insights?.trackException({ exception: removeOrganizationsError });
      if (!continueOnError) {
        throw removeOrganizationsError;
      }
      history.push(
        `Organization removal error: ${removeOrganizationsError.toString()}`
      );
    }

    // Link
    try {
      if (account.link) {
        history.push(...(await account.removeLink()));
      }
    } catch (removeLinkError) {
      ++errors;
      insights?.trackException({ exception: removeLinkError });
      if (account.id && !account.login) {
        // If the account information could not be resolved through the
        // GitHub API for this _id_, then the user deleted their account,
        // or is using a different one now, etc., so try deleting the
        // associated link still by searching for it by _ID_.
        // TODO: Ported code from account.ts, remains unimp. It's OK.
      }
      if (!continueOnError) {
        throw removeLinkError;
      }
      history.push(`Unlink error: ${removeLinkError.toString()}`);
    }

    // Collaborator permissions to repositories
    try {
      const removed = await account.removeCollaboratorPermissions();
      history.push(...removed.history);
      if (removed.error) {
        throw removed.error;
      }
    } catch (removeCollaboratorsError) {
      ++errors;
      insights?.trackException({ exception: removeCollaboratorsError });
      if (account.id && !account.login) {
        // If the account information could not be resolved through the
        // GitHub API for this _id_, then the user deleted their account,
        // or is using a different one now, etc., so try deleting the
        // associated link still by searching for it by _ID_.
        // TODO: Ported code from account.ts, remains unimp. It's OK.
      }
      if (!continueOnError) {
        throw removeCollaboratorsError;
      }
      history.push(
        `Collaborator remove error: ${removeCollaboratorsError.toString()}`
      );
    }

    // Notify
    try {
      const status = await this.sendTerminatedAccountMail(
        account,
        purpose,
        history,
        errors
      );
      if (status) {
        history.push(
          `Unlink e-mail sent to manager: to=${status.to.join(
            ', '
          )} bcc=${status.bcc.join(', ')}, receipt=${status.receipt}`
        );
      } else {
        history.push('Service not configured to notify by mail');
      }
    } catch (notifyTerminationMailError) {
      insights?.trackException({ exception: notifyTerminationMailError });
      // Notification should never throw
      history.push('Unlink e-mail COULD NOT be sent to manager');
    }

    // Telemetry
    insights?.trackEvent({
      name: 'UserUnlink',
      properties: {
        id: account.id,
        login: account.login,
        reason: reason,
        purpose,
        continueOnError: continueOnError
          ? 'continue on errors'
          : 'halt on errors',
        history: JSON.stringify(history),
      },
    });

    return history;
  }

  getOperationsMailAddress(): string {
    return this.config.brand.operationsMail;
  }

  getInfrastructureNotificationsMail(): string {
    return (
      this.config.notifications.infrastructureNotificationsMail ||
      this.getOperationsMailAddress()
    );
  }

  getLinksNotificationMailAddress(): string {
    return (
      this.config.notifications.linksMailAddress ||
      this.getOperationsMailAddress()
    );
  }

  getRepositoriesNotificationMailAddress(): string {
    return (
      this.config.notifications.reposMailAddress ||
      this.getOperationsMailAddress()
    );
  }

  private sendTerminatedAccountMail(
    account: Account,
    purpose: UnlinkPurpose,
    details: string[],
    errorsCount: number
  ): Promise<IUnlinkMailStatus> {
    return sendTerminatedAccountMailMethod(
      this,
      account,
      purpose,
      details,
      errorsCount
    );
  }

  getDefaultRepositoryTemplateNames(): string[] {
    if (
      this.config.github &&
      this.config.github.templates &&
      this.config.github.templates.defaultTemplates &&
      this.config.github.templates.defaultTemplates.length > 0
    ) {
      return this.config.github.templates.defaultTemplates as string[];
    }
    return null;
  }

  getDefaultLegalEntities(): string[] {
    if (
      this.config.legalEntities &&
      this.config.legalEntities.defaultOrganizationEntities &&
      this.config.legalEntities.defaultOrganizationEntities.length > 0
    ) {
      return this.config.legalEntities.defaultOrganizationEntities as string[];
    }
    return null;
  }

  getOrganization(name: string): Organization {
    if (!name) {
      throw CreateError.ParameterRequired('name');
    }
    const lc = name.toLowerCase();
    const organization = this.organizations.get(lc);
    if (!organization) {
      throw CreateError.NotFound(
        `Could not find configuration for the "${name}" organization.`
      );
    }
    return organization;
  }

  isOrganizationManagedById(organizationId: number): boolean {
    try {
      this.getOrganizationById(organizationId);
      return true;
    } catch (notConfigured) {
      return false;
    }
  }

  getOrganizationById(organizationId: number): Organization {
    if (typeof organizationId === 'string') {
      organizationId = parseInt(organizationId, 10);
      console.warn(`getOrganizationById: organizationId must be a number`);
    }
    if (!this._organizationIds) {
      this.getOrganizationIds();
    }
    const org = this._organizationIds.get(organizationId);
    if (!org) {
      throw CreateError.NotFound(
        `getOrganizationById: no configured ID for an organization with ID ${organizationId}`
      );
    }
    return org;
  }

  async getRepos(options?: ICacheOptions): Promise<Repository[]> {
    const repos: Repository[] = [];
    const cacheOptions = options || {
      maxAgeSeconds: this.defaults.crossOrgsReposStaleSecondsPerOrg,
    };
    // CONSIDER: Cross-org functionality might be best in the GitHub library itself
    const orgs = this.organizations.values();
    for (let organization of orgs) {
      try {
        const organizationRepos = await organization.getRepositories(
          cacheOptions
        );
        repos.push(...organizationRepos);
      } catch (orgReposError) {
        console.dir(orgReposError);
      }
    }
    return repos;
  }

  getLinks(options?: any): Promise<ICorporateLink[]> {
    // Design change in the TypeScript version: this returns true link objects now,
    // but caches hydrated links behind the scenes
    options = options || {
      includeNames: true,
      includeId: true,
      includeServiceAccounts: true,
    };
    const caching = {
      maxAgeSeconds:
        options.maxAgeSeconds || this.defaults.corporateLinksStaleSeconds,
      backgroundRefresh: true,
    };
    delete options.maxAgeSeconds;
    delete options.backgroundRefresh;
    const linkProvider = this.providers.linkProvider;
    options.lp = linkProvider.serializationIdentifierVersion;
    return new Promise((resolve, reject) => {
      return this.github.links.getCachedLinks(
        getPromisedLinks.bind(null, linkProvider),
        options,
        caching,
        (ee, ll) => {
          let rehydratedLinks = null;
          if (!ee) {
            try {
              rehydratedLinks = linkProvider.rehydrateLinks(ll);
            } catch (rehydrationError) {
              ee = rehydrationError;
            }
          }
          if (ee) {
            return reject(ee);
          }
          return resolve(rehydratedLinks);
        }
      );
    });
  }

  async getLinksFromThirdPartyIds(
    thirdPartyIds: string[]
  ): Promise<ICorporateLink[]> {
    const corporateLinks: ICorporateLink[] = [];
    const throttle = throat(ParallelLinkLookup);
    await Promise.all(
      thirdPartyIds.map((thirdPartyId) =>
        throttle(async () => {
          try {
            const link = await this.getLinkByThirdPartyId(thirdPartyId);
            if (link) {
              corporateLinks.push(link);
            }
          } catch (noLinkError) {
            if (!ErrorHelper.IsNotFound(noLinkError)) {
              console.dir(noLinkError);
            }
          }
        })
      )
    );
    return corporateLinks;
  }

  async getLinksFromCorporateIds(
    corporateIds: string[]
  ): Promise<ICorporateLink[]> {
    const corporateLinks: ICorporateLink[] = [];
    const throttle = throat(ParallelLinkLookup);
    await Promise.all(
      corporateIds.map((corporateId) =>
        throttle(async () => {
          try {
            const links = await this.providers.linkProvider.queryByCorporateId(
              corporateId
            );
            if (links && links.length === 1) {
              corporateLinks.push(links[0]);
            } else if (links.length > 1) {
              throw new Error('Multiple links not supported');
            }
          } catch (noLinkError) {
            console.dir(noLinkError);
          }
        })
      )
    );
    return corporateLinks;
  }

  getLinkByThirdPartyId(thirdPartyId: string): Promise<ICorporateLink> {
    const linkProvider = this.providers.linkProvider;
    return linkProvider.getByThirdPartyId(thirdPartyId);
  }

  getLinkByThirdPartyUsername(username: string): Promise<ICorporateLink> {
    const linkProvider = this.providers.linkProvider;
    return linkProvider.getByThirdPartyUsername(username);
  }

  getMailAddressFromCorporateUsername(
    corporateUsername: string
  ): Promise<string> {
    if (!this.providers.mailAddressProvider) {
      throw new Error('No mailAddressProvider available');
    }
    return this.providers.mailAddressProvider.getAddressFromUpn(
      corporateUsername
    );
  }

  async getMailAddressesFromCorporateUsernames(
    corporateUsernames: string[]
  ): Promise<string[]> {
    // This is a best-faith effort but will not fail if some are not returned.
    const throttle = throat(2);
    const addresses: string[] = [];
    await Promise.all(
      corporateUsernames.map((username) =>
        throttle(async () => {
          try {
            const address = await this.getMailAddressFromCorporateUsername(
              username
            );
            if (address) {
              addresses.push(address);
            }
          } catch (ignoreError) {
            console.log('getMailAddressesFromCorporateUsernames error:');
            console.warn(ignoreError);
          }
        })
      )
    );
    return addresses;
  }

  async tryGetLink(id: string, options?): Promise<ICorporateLink> {
    if (this.providers.linkProvider) {
      try {
        const link = await this.getLinkByThirdPartyId(id);
        return link;
      } catch (error) {
        if (ErrorHelper.IsNotFound(error)) {
          return null;
        } else {
          throw error;
        }
      }
    }
    // This literally retrieves the cache of all links, built from a time before link provider.
    const links = await this.getLinks(options);
    const reduced = links.filter((link) => {
      // was 'ghid' in the prior implementation before link interfaces
      return link && link.thirdPartyId == id /* allow string comparisons */;
    });
    if (reduced.length > 1) {
      throw new Error(
        `Multiple links were present for the same GitHub user ${id}`
      );
    }
    return reduced.length === 1 ? reduced[0] : null;
  }

  getTeamsWithMembers(
    options?: ICrossOrganizationTeamMembership
  ): Promise<any> {
    const cacheOptions: IPagedCrossOrganizationCacheOptions = {};
    options = options || {};
    cacheOptions.backgroundRefresh =
      options.backgroundRefresh !== undefined
        ? options.backgroundRefresh
        : true;
    cacheOptions.maxAgeSeconds = options.maxAgeSeconds || 60 * 10;
    cacheOptions.individualMaxAgeSeconds = options.individualMaxAgeSeconds;
    delete options.backgroundRefresh;
    delete options.maxAgeSeconds;
    delete options.individualMaxAgeSeconds;
    return this.github.crossOrganization.teamMembers(
      this._organizationNamesWithAuthorizationHeaders,
      options,
      cacheOptions
    );
  }

  getRepoCollaborators(
    options: IPagedCrossOrganizationCacheOptions
  ): Promise<any> {
    const cacheOptions: IPagedCrossOrganizationCacheOptions = {};
    options = options || {};
    cacheOptions.backgroundRefresh =
      options.backgroundRefresh !== undefined
        ? options.backgroundRefresh
        : true;
    cacheOptions.maxAgeSeconds = options.maxAgeSeconds || 60 * 10;
    cacheOptions.individualMaxAgeSeconds = options.individualMaxAgeSeconds;
    delete options.backgroundRefresh;
    delete options.maxAgeSeconds;
    delete options.individualMaxAgeSeconds;
    return this.github.crossOrganization.repoCollaborators(
      this.organizationNamesWithAuthorizationHeaders,
      options,
      cacheOptions
    );
  }

  getRepoTeams(options: IPagedCrossOrganizationCacheOptions): Promise<any> {
    const cacheOptions: IPagedCrossOrganizationCacheOptions = {};
    options = options || {};
    cacheOptions.backgroundRefresh =
      options.backgroundRefresh !== undefined
        ? options.backgroundRefresh
        : true;
    cacheOptions.maxAgeSeconds = options.maxAgeSeconds || 60 * 10;
    cacheOptions.individualMaxAgeSeconds = options.individualMaxAgeSeconds;
    delete options.backgroundRefresh;
    delete options.maxAgeSeconds;
    delete options.individualMaxAgeSeconds;
    return this.github.crossOrganization.repoTeams(
      this.organizationNamesWithAuthorizationHeaders,
      options,
      cacheOptions
    );
  }

  async getCrossOrganizationTeams(
    options?: any
  ): Promise<ICrossOrganizationMembersResult> {
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = this.defaults.crossOrgsMembersStaleSecondsPerOrg;
    }
    if (options.backgroundRefresh === undefined) {
      options.backgroundRefresh = true;
    }
    const cacheOptions = {
      maxAgeSeconds: options.maxAgeSeconds,
      backgroundRefresh: options.backgroundRefresh,
    };
    delete options.maxAgeSeconds;
    delete options.backgroundRefresh;
    const values = await this.github.crossOrganization.teams(
      this.organizationNamesWithAuthorizationHeaders,
      options,
      cacheOptions
    );
    const results = crossOrganizationResults(this, values, 'id');
    return results;
  }

  async getMembers(
    options?: ICacheOptions
  ): Promise<ICrossOrganizationMembersResult> {
    options = options || {};
    if (!options.maxAgeSeconds) {
      options.maxAgeSeconds = this.defaults.crossOrgsMembersStaleSecondsPerOrg;
    }
    if (options.backgroundRefresh === undefined) {
      options.backgroundRefresh = true;
    }
    const cacheOptions = {
      maxAgeSeconds: options.maxAgeSeconds,
      backgroundRefresh: options.backgroundRefresh,
    };
    delete options.maxAgeSeconds;
    delete options.backgroundRefresh;
    const values = await this.github.crossOrganization.orgMembers(
      this.organizationNamesWithAuthorizationHeaders,
      options,
      cacheOptions
    );
    const crossOrgReturn = (crossOrganizationResults(
      this,
      values,
      'id'
    ) as any) as ICrossOrganizationMembersResult;
    return crossOrgReturn;
  }

  // Feature flags

  allowSelfServiceTeamMemberToMaintainerUpgrades() {
    return (
      this.config?.features?.allowTeamMemberToMaintainerSelfUpgrades === true
    );
  }

  allowUnauthorizedNewRepositoryLockdownSystemFeature() {
    return (
      this.config?.features?.allowUnauthorizedNewRepositoryLockdownSystem ===
      true
    );
  }

  allowUnauthorizedForkLockdownSystemFeature() {
    // This feature has a hard dependency on the new repo lockdown system itself
    return (
      this.allowUnauthorizedNewRepositoryLockdownSystemFeature() &&
      this.config &&
      this.config.features &&
      this.config.features.allowUnauthorizedForkLockdownSystem === true
    );
  }

  allowTransferLockdownSystemFeature() {
    // This feature has a hard dependency on the new repo lockdown system itself
    return (
      this.allowUnauthorizedNewRepositoryLockdownSystemFeature() &&
      this.config &&
      this.config.features &&
      this.config.features.allowUnauthorizedTransferLockdownSystem === true
    );
  }

  allowUndoSystem() {
    return this.config?.features?.features.allowUndoSystem === true;
  }

  // Eventually link/unlink should move from context into operations here to centralize more than just the events

  async fireLinkEvent(value): Promise<void> {
    await fireEvent(this.config, 'link', value);
  }

  async fireUnlinkEvent(value): Promise<void> {
    await fireEvent(this.config, 'unlink', value);
  }

  get systemAccountsByUsername(): string[] {
    return this.config?.github?.systemAccounts
      ? this.config.github.systemAccounts.logins
      : [];
  }

  isPortalSudoer(githubLogin: string, link: ICorporateLink) {
    if (!this.initialized) {
      throw new Error('The application is not yet initialized');
    }
    return this._portalSudo.isSudoer(githubLogin, link);
  }

  isSystemAccountByUsername(username: string): boolean {
    const lc = username.toLowerCase();
    const usernames = this.systemAccountsByUsername;
    for (let i = 0; i < usernames.length; i++) {
      if (usernames[i].toLowerCase() === lc) {
        return true;
      }
    }
    return false;
  }

  getCentralOperationsToken(): IGetAuthorizationHeader {
    const func = getCentralOperationsAuthorizationHeader.bind(
      null,
      this
    ) as IGetAuthorizationHeader;
    return func;
  }

  getAccount(id: string) {
    const entity = { id };
    return new Account(
      entity,
      this,
      getCentralOperationsAuthorizationHeader.bind(null, this)
    );
  }

  async getAccountWithDetailsAndLink(id: string): Promise<Account> {
    const account = this.getAccount(id);
    return await account.getDetailsAndLink();
  }

  async getAuthenticatedAccount(token: string): Promise<Account> {
    const github = this.github;
    const parameters = {};
    try {
      const entity = await github.post(
        `token ${token}`,
        'users.getAuthenticated',
        parameters
      );
      const account = new Account(
        entity,
        this,
        getCentralOperationsAuthorizationHeader.bind(null, this)
      );
      return account;
    } catch (error) {
      throw wrapError(
        error,
        'Could not get details about the authenticated account'
      );
    }
  }

  getTeamByIdWithOrganization(
    id: number,
    organizationName: string,
    entity?: any
  ): Team {
    const organization = this.getOrganization(organizationName);
    return organization.team(id, entity);
  }

  getRepositoryWithOrganization(
    name: string,
    organizationName: string,
    entity?: any
  ): Repository {
    const organization = this.getOrganization(organizationName);
    return organization.repository(name, entity);
  }

  async sendMail(mail: IMail): Promise<any> {
    const mailProvider = this.providers.mailProvider;
    const insights = this.providers.insights;
    const customData = {
      receipt: null,
      eventName: undefined,
    };
    try {
      const mailResult = await mailProvider.sendMail(mail);
      customData.receipt = mailResult;
      insights.trackEvent({ name: 'MailSuccess', properties: customData });
      return mailResult;
    } catch (mailError) {
      customData.eventName = 'MailFailure';
      insights.trackException({ exception: mailError, properties: customData });
      throw mailError;
    }
  }

  async emailRender(
    emailViewName: string,
    contentOptions: any
  ): Promise<string> {
    const appDirectory = this.config.typescript.appDirectory;
    return await RenderHtmlMail(appDirectory, emailViewName, contentOptions);
  }
}

interface IFireEventResult {
  url: string;
  value: string;
  body: string;
  headers: any;
  statusCode: any;
}

async function fireEvent(
  config,
  configurationName,
  value
): Promise<IFireEventResult[]> {
  if (
    !config ||
    !config.github ||
    !config.github.links ||
    !config.github.links.events ||
    !config.github.links.events
  ) {
    return;
  }
  const userAgent = config.userAgent || 'Unknown user agent';
  const httpUrls = config.github.links.events.http;
  if (!httpUrls || !httpUrls[configurationName]) {
    return;
  }
  const urlOrUrls = httpUrls[configurationName];
  let urls = Array.isArray(urlOrUrls) ? urlOrUrls : [urlOrUrls];
  let results: IFireEventResult[] = [];
  for (const httpUrl of urls) {
    try {
      const response = await axios({
        method: 'POST',
        url: httpUrl,
        data: value,
        headers: {
          'User-Agent': userAgent,
          'X-Repos-Event': configurationName,
        },
      });
      results.push({
        url: httpUrl,
        value,
        headers: response.headers,
        body: response.data,
        statusCode: response.status,
      });
    } catch (ignoredTechnicalError) {
      /* ignored */
      // TODO: telemetry
      console.warn(ignoredTechnicalError);
      console.log();
    }
  }
  return results;
}

export function getCentralOperationsAuthorizationHeader(
  self: Operations
): IPurposefulGetAuthorizationHeader {
  const s = (self || this) as Operations;
  if (
    s.config.github &&
    s.config.github.operations &&
    s.config.github.operations.centralOperationsToken
  ) {
    const capturedToken = s.config.github.operations.centralOperationsToken;
    return async () => {
      return {
        value: `token ${capturedToken}`,
        purpose: null, // legacy
        source: 'central operations token',
      };
    };
  } else if (s.getOrganizations.length === 0) {
    throw new Error(
      'No central operations token nor any organizations configured.'
    );
  }
  // Fallback to the first configured organization as a convenience
  // CONSIDER: would randomizing the organization be better, or a priority based on known-rate limit remaining?
  const firstOrganization = s.getOrganizations()[0];
  return firstOrganization.getAuthorizationHeader();
}

function crossOrganizationResults(
  operations: Operations,
  results,
  keyProperty
) {
  keyProperty = keyProperty || 'id';
  const map: IMapPlusMetaCost = new Map();
  operations.translateOrganizationNamesFromLowercase(results.orgs);
  for (const orgName of Object.getOwnPropertyNames(results.orgs)) {
    const orgValues = results.orgs[orgName];
    for (let i = 0; i < orgValues.length; i++) {
      const val = orgValues[i];
      const key = val[keyProperty];
      if (!key) {
        throw new Error(
          `Entity missing property ${key} during consolidation processing.`
        );
      }
      let mapEntry = map.get(key);
      if (!mapEntry) {
        mapEntry = {
          orgs: {},
        };
        mapEntry[keyProperty] = key;
        map.set(key, mapEntry);
      }
      mapEntry.orgs[orgName] = val;
    }
  }
  map.headers = results.headers;
  map.cost = results.cost;
  return map;
}

async function getPromisedLinks(
  linkProvider: ILinkProvider
): Promise<IPromisedLinks> {
  // TODO: consider looking at the options as to how to include/exclude properties etc.
  // today (TypeScript update with PGSQL) the 'options' have zero impact on what is actually returned...
  const links = await linkProvider.getAll();
  const jsonLinks = linkProvider.dehydrateLinks(links);
  const dataObject: IPromisedLinks = {
    headers: {
      type: 'links',
    },
    data: jsonLinks,
  };
  return dataObject;
}

function restartAfterDynamicConfigurationUpdate(
  minimumSeconds: number,
  maximumSeconds: number,
  appInitialized: Date,
  organizationSettingsProvider: OrganizationSettingProvider
) {
  didDynamicConfigurationUpdate(appInitialized, organizationSettingsProvider)
    .then((changed) => {
      if (changed) {
        const randomSeconds = Math.floor(
          Math.random() * (maximumSeconds - minimumSeconds + 1) + minimumSeconds
        );
        console.log(
          `changes to dynamic configuration detected since ${appInitialized}, restarting in ${randomSeconds}s`
        );
        setInterval(() => {
          console.log(
            `shutting down process due to dynamic configuration changes being detected at least ${randomSeconds} seconds ago...`
          );
          return process.exit(0);
        }, randomSeconds * 1000);
        if (DynamicRestartCheckHandle) {
          clearInterval(DynamicRestartCheckHandle);
        }
      }
    })
    .catch((error) => {
      console.dir(error);
    });
}

async function didDynamicConfigurationUpdate(
  appInitialized: Date,
  organizationSettingsProvider: OrganizationSettingProvider
): Promise<boolean> {
  try {
    const allOrganizations = await organizationSettingsProvider.queryAllOrganizations();
    const activeOrganizations = allOrganizations.filter((org) => org.active);
    for (const organization of activeOrganizations) {
      if (organization.updated > appInitialized) {
        console.log(
          `organization ${organization.organizationName} was updated ${organization.updated} vs app started time of ${appInitialized}`
        );
        return true;
      }
    }
  } catch (updateOrganizationsError) {
    console.dir(updateOrganizationsError);
  }
  return false;
}
