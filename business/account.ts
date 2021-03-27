//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import _ from 'lodash';

import {
  ICacheOptions,
  IReposError,
  IGetAuthorizationHeader,
  ErrorHelper,
  IOperationsInstance,
  operationsWithCapability,
  CoreCapability,
  IOperationsProviders,
  throwIfNotCapable,
  IOperationsLinks,
  throwIfNotGitHubCapable,
} from '../transitional';
import * as common from './common';

import { wrapError } from '../utils';
import { corporateLinkToJson, ICorporateLink } from './corporateLink';
import { Organization, OrganizationMembershipState } from './organization';
import { AppPurpose } from '../github';
import { ILinkProvider } from '../lib/linkProviders';
import { CacheDefault, getMaxAgeSeconds } from '.';

interface IRemoveOrganizationMembershipsResult {
  error?: IReposError;
  history: string[];
}

export enum AccountJsonFormat {
  GitHub = 'github',
  UplevelWithLink = 'github+link',
}

const primaryAccountProperties = ['id', 'login', 'avatar_url'];
const secondaryAccountProperties = [];

export class Account {
  private _operations: IOperationsInstance;
  private _getAuthorizationHeader: IGetAuthorizationHeader;

  private _link: ICorporateLink;
  private _id: number;

  private _login: string;
  private _avatar_url?: string;
  private _created_at?: any;
  private _updated_at?: any;

  private _originalEntity?: any;

  public asJson(format: AccountJsonFormat = AccountJsonFormat.GitHub) {
    const basic = {
      avatar_url: this.avatar_url,
      id: this.id,
      login: this.login,
      created_at: this.created_at,
      updated_at: this.updated_at,
    };
    switch (format) {
      case AccountJsonFormat.GitHub: {
        return basic;
      }
      case AccountJsonFormat.UplevelWithLink: {
        const link = this._link ? corporateLinkToJson(this._link) : undefined;
        return {
          account: basic,
          isLinked: !!link,
          link,
        };
      }
      default: {
        throw new Error(`Unsupported asJson format: ${format}`);
      }
    }
  }

  public get id(): number {
    return this._id;
  }

  public get link(): ICorporateLink {
    return this._link;
  }

  public get login(): string {
    return this._login;
  }

  public get avatar_url(): string {
    return this._avatar_url;
  }

  public get updated_at(): any {
    return this._updated_at;
  }

  public get created_at(): any {
    return this._created_at;
  }

  public get company(): string {
    return this._originalEntity ? this._originalEntity.company : undefined;
  }

  public get email(): string {
    return this._originalEntity ? this._originalEntity.email : undefined;
  }

  public get name(): string {
    return this._originalEntity ? this._originalEntity.name : undefined;
  }

  constructor(
    entity,
    operations: IOperationsInstance,
    getAuthorizationHeader: IGetAuthorizationHeader
  ) {
    common.assignKnownFieldsPrefixed(
      this,
      entity,
      'account',
      primaryAccountProperties,
      secondaryAccountProperties
    );
    this._originalEntity = entity;
    this._operations = operations;
    this._getAuthorizationHeader = getAuthorizationHeader;
  }

  getEntity() {
    return this._originalEntity;
  }

  // TODO: looks like we need to be able to resolve the link in here, too, to set instance.link

  // These were previously implemented in lib/user.js; functions may be needed
  // May also need to be in the org- and team- specific accounts, or built as proper objects

  contactName() {
    if (this._link) {
      return (
        this.link.corporateDisplayName || this.link['aadname'] || this._login
      );
    }
    return this._login;
  }

  contactEmail() {
    // NOTE: this field name is wrong, it is a UPN, not a mail address
    if (this._link) {
      return this.link.corporateUsername || this.link['aadupn'] || null;
    }
    return null;
  }

  corporateAlias() {
    // NOTE: this is a hack
    if (this.contactEmail()) {
      var email = this.contactEmail();
      var i = email.indexOf('@');
      if (i >= 0) {
        return email.substring(0, i);
      }
    }
  }

  corporateProfileUrl() {
    const operations = operationsWithCapability<IOperationsProviders>(
      this._operations,
      CoreCapability.Providers
    );
    if (operations) {
      const config = operations.providers.config;
      const alias = this.corporateAlias();
      const corporateSettings = config.corporate;
      if (
        alias &&
        corporateSettings &&
        corporateSettings.profile &&
        corporateSettings.profile.prefix
      ) {
        return corporateSettings.profile.prefix + alias;
      }
    }
  }

  avatar(optionalSize) {
    if (this._avatar_url) {
      return this._avatar_url + '&s=' + (optionalSize || 80);
    }
  }

  // End previous functions

  async getDetailsAndLink(options?: ICacheOptions): Promise<Account> {
    try {
      await this.getDetails(options || {});
    } catch (getDetailsError) {
      // If a GitHub account is deleted, this would fail
      console.dir(getDetailsError);
    }
    await this.tryGetLink();
    return this;
  }

  async tryGetLink() {
    const operations = throwIfNotCapable<IOperationsLinks>(
      this._operations,
      CoreCapability.Links
    );
    try {
      this._link = await operations.tryGetLink(this._id.toString());
    } catch (getLinkError) {
      // We do not assume that the link exists...
      console.dir(getLinkError);
    }
  }

  async getRecentEventsFirstPage(options?: ICacheOptions): Promise<any[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const login = this.login;
    if (!login) {
      throw new Error(
        'Must provide a GitHub login to retrieve account events.'
      );
    }
    const parameters = {
      username: login,
    };
    const cacheOptions: ICacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.accountDetailStaleSeconds,
        options,
        60
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    try {
      const entity = await operations.github.call(
        this.authorize(AppPurpose.Data),
        'activity.listEventsForAuthenticatedUser',
        parameters,
        cacheOptions
      );
      return entity;
    } catch (error) {
      console.dir(error);
      throw error;
    }
  }

  async getEvents(options?: ICacheOptions): Promise<any[]> {
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    const login = this.login;
    if (!login) {
      throw new Error(
        'Must provide a GitHub login to retrieve account events.'
      );
    }
    const parameters = {
      username: login,
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
      const events = await operations.github.collections.getUserActivity(
        this.authorize(AppPurpose.Data),
        parameters,
        cacheOptions
      );
      let cached = true;
      if (events && events.cost && events.cost.github.usedApiTokens > 0) {
        cached = false;
      }
      let arr = [...events];
      arr['cached'] = cached;
      return arr;
    } catch (error) {
      if (error.status && error.status === 404) {
        error = new Error(
          `The GitHub user ${login} could not be found (or has been deleted)`
        );
        error.status = 404;
        throw error;
      }
      throw wrapError(
        error,
        `Could not get events for login ${login}: ${error.message}`
      );
    }
  }

  async getDetailsAndDirectLink(): Promise<Account> {
    if (
      !throwIfNotCapable<IOperationsProviders>(
        this._operations,
        CoreCapability.Providers
      ).providers.linkProvider
    ) {
      throw new Error(
        'getDetailsAndDirectLink: this method can only be called when a linkProvider is used'
      );
    }
    const operations = throwIfNotCapable<IOperationsLinks>(
      this._operations,
      CoreCapability.Links
    );
    try {
      await this.getDetails();
    } catch (getDetailsError) {
      // If a GitHub account is deleted, this would fail
      // TODO: should this throw then?
      console.dir(getDetailsError);
    }
    try {
      let link = await operations.getLinkByThirdPartyId(this._id.toString());
      if (link) {
        this._link = link;
      }
    } catch (getLinkError) {
      // We do not assume that the link exists...
      // TODO: can we only ignore WHEN WE KNOW ITS a 404 no link vs a 500?
      console.dir(getLinkError);
    }
    return this;
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
    const id = this._id;
    if (!id) {
      throw new Error(
        'Must provide a GitHub user ID to retrieve account information.'
      );
    }
    const parameters = {
      id,
    };
    const cacheOptions: ICacheOptions = {
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.accountDetailStaleSeconds,
        options,
        60
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    try {
      const entity = await operations.github.request(
        this.authorize(AppPurpose.Data),
        'GET /user/:id',
        parameters,
        cacheOptions
      );
      common.assignKnownFieldsPrefixed(
        this,
        entity,
        'account',
        primaryAccountProperties,
        secondaryAccountProperties
      );
      return entity;
    } catch (error) {
      if (error.status && error.status === 404) {
        error = new Error(
          `The GitHub user ID ${id} could not be found (or has been deleted)`
        );
        error.status = 404;
        throw error;
      }
      throw wrapError(
        error,
        `Could not get details about account ID ${id}: ${error.message}`
      );
    }
  }

  async removeLink(): Promise<any> {
    const operations = throwIfNotCapable<IOperationsProviders>(
      this._operations,
      CoreCapability.Providers
    );
    const opsLinks = throwIfNotCapable<IOperationsLinks>(
      this._operations,
      CoreCapability.Links
    );
    const linkProvider = operations.providers.linkProvider as ILinkProvider;
    if (!linkProvider) {
      throw new Error('No link provider');
    }
    const id = this._id;
    try {
      await this.getDetailsAndDirectLink();
    } catch (getDetailsError) {
      // We ignore any error to make sure link removal always works
      const insights = operations.providers.insights;
      if (getDetailsError) {
        insights?.trackException({
          exception: getDetailsError,
          properties: {
            id,
            location: 'account.removeLink',
          },
        });
      }
    }
    let aadIdentity = undefined;
    const link = this._link;
    if (!link) {
      throw new Error(
        `No link is associated with the instance for account ${id}`
      );
    }
    aadIdentity = {
      preferredName: link.corporateDisplayName,
      userPrincipalName: link.corporateUsername,
      id: link.corporateId,
    };
    const eventData = {
      github: {
        id: id,
        login: this._login,
      },
      aad: aadIdentity,
    };
    const history = [];
    let finalError = null;
    try {
      await deleteLink(linkProvider, link);
    } catch (linkDeleteError) {
      const message =
        linkDeleteError.statusCode === 404
          ? `The link for ID ${id} no longer exists: ${linkDeleteError}`
          : `The link for ID ${id} could not be removed: ${linkDeleteError}`;
      history.push(message);
      finalError = linkDeleteError;
    }
    if (!finalError) {
      opsLinks.fireUnlinkEvent(eventData);
      history.push(
        `The link for ID ${id} has been removed from the link service`
      );
    }
    return history;
  }

  // TODO: implement getOrganizationMemberships, with caching; reuse below code

  async getOperationalOrganizationMemberships(): Promise<Organization[]> {
    const operations = throwIfNotGitHubCapable(this._operations);
    await this.getDetails();
    const username = this._login; // we want to make sure that we have an ID and username
    if (!username) {
      throw new Error(`No GitHub username available for user ID ${this._id}`);
    }
    let currentOrganizationMemberships: Organization[] = [];
    const checkOrganization = async (organization) => {
      try {
        const result = await organization.getOperationalMembership(username);
        if (
          result &&
          result.state &&
          (result.state === OrganizationMembershipState.Active ||
            result.state === OrganizationMembershipState.Pending)
        ) {
          currentOrganizationMemberships.push(organization);
        }
      } catch (ignoreErrors) {
        // getMembershipError ignored: if there is no membership that's fine
        console.log(
          `error from individual check of organization ${organization.name} membership for username ${username}: ${ignoreErrors}`
        );
      }
    };
    const opsAs = operations as any;
    if (!opsAs.organizations) {
      throw new Error('Operations does not expose an organizations Map getter');
    }
    const allOrganizations = Array.from(
      opsAs.organizations.values() as Organization[]
    );
    const staticOrganizations = allOrganizations.filter(
      (org) => org.hasDynamicSettings === false
    );
    const dynamicOrganizations = allOrganizations.filter(
      (org) => org.hasDynamicSettings
    );
    await Promise.all(dynamicOrganizations.map(checkOrganization));
    for (let organization of staticOrganizations) {
      await checkOrganization(organization);
    }
    return currentOrganizationMemberships;
  }

  async removeCollaboratorPermissions(): Promise<IRemoveOrganizationMembershipsResult> {
    const history = [];
    let error: IReposError = null;
    const operations = throwIfNotGitHubCapable(this._operations);
    const opsWithProvs = operationsWithCapability<IOperationsProviders>(
      operations,
      CoreCapability.Providers
    );
    const { queryCache } = opsWithProvs?.providers;
    if (!queryCache || !queryCache.supportsRepositoryCollaborators) {
      history.push(
        'The account may still have Collaborator permissions to repositories'
      );
      return { history };
    }
    if (!this.login) {
      await this.getDetails();
    }
    const collaborativeRepos = await queryCache.userCollaboratorRepositories(
      this.id.toString()
    );
    for (const entry of collaborativeRepos) {
      const { repository } = entry;
      try {
        await repository.getDetails();
        await repository.removeCollaborator(this.login);
        history.push(
          `Removed ${this.login} as a Collaborator from the repository ${repository.full_name}`
        );
      } catch (removeCollaboratorError) {
        if (ErrorHelper.IsNotFound(removeCollaboratorError)) {
          // The repo doesn't exist any longer, this is OK.
        } else {
          history.push(
            `Could not remove ${this.login} as a Collaborator from the repo: ${repository.full_name}`
          );
        }
      }
    }
    return { history, error };
  }

  async removeManagedOrganizationMemberships(): Promise<IRemoveOrganizationMembershipsResult> {
    const history = [];
    let error: IReposError = null;
    let organizations: Organization[];
    try {
      organizations = await this.getOperationalOrganizationMemberships();
    } catch (getMembershipError) {
      if (
        getMembershipError &&
        getMembershipError.status == /* loose */ '404'
      ) {
        history.push(getMembershipError.toString());
      } else if (getMembershipError) {
        throw getMembershipError;
      }
    }
    const username = this._login;
    if (organizations && organizations.length > 1) {
      const asText = _.map(organizations, (org) => {
        return org.name;
      }).join(', ');
      history.push(
        `${username} is a member of the following organizations: ${asText}`
      );
    } else if (organizations) {
      history.push(`${username} is not a member of any managed organizations`);
    }
    if (organizations && organizations.length) {
      for (let i = 0; i < organizations.length; i++) {
        const organization = organizations[i];
        try {
          await organization.removeMember(username);
          history.push(
            `Removed ${username} from the ${organization.name} organization`
          );
        } catch (removeError) {
          history.push(
            `Error while removing ${username} from the ${organization.name} organization: ${removeError}`
          );
          if (!error) {
            error = removeError;
          }
        }
      }
    }
    return { history, error };
  }

  private authorize(purpose: AppPurpose): IGetAuthorizationHeader | string {
    const getAuthorizationHeader = this._getAuthorizationHeader.bind(
      this,
      purpose
    ) as IGetAuthorizationHeader;
    return getAuthorizationHeader;
  }
}

function deleteLink(
  linkProvider: ILinkProvider,
  link: ICorporateLink
): Promise<any> {
  return linkProvider.deleteLink(link);
}
