//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  ReposAppRequest,
  IReposAppResponse,
  UserAlertType,
  IDictionary,
  IProviders,
  IAppSession,
} from '../transitional';

import { ICorporateLink } from '../business/corporateLink';

import { addBreadcrumb } from '../utils';
import { Operations } from '../business/operations';
import { UserContext } from './aggregate';

import pugLoad from 'pug-load';
import fs from 'fs';

import objectPath from 'object-path';

const debug = require('debug')('context');

// - - - identity

export enum GitHubIdentitySource {
  Link = 'link',
  Session = 'session',
}

export interface IGitHubIdentity {
  id: string;
  username: string;
  avatar?: string;
  displayName?: string;

  source: GitHubIdentitySource;
}

export interface ICorporateIdentity {
  id: string;
  username: string;
  displayName?: string;
}

// - - - web

export interface IWebContextOptions {
  baseUrl?: string;
  request: ReposAppRequest;
  response: IReposAppResponse;
  sessionUserProperties: SessionUserProperties;
}

export interface IWebPageRenderOptions {
  title?: string;
  view: string;

  state?: any;
  optionalObject?: any;
}

// legacy structure
interface IWebPageRenderUser {
  primaryAuthenticationScheme: string;
  primaryUsername: string;
  githubSignout: string;
  azureSignout: string;
  github?: {
    id?: string;
    username?: string;
    displayName?: string;
    avatarUrl?: string;
    accessToken?: boolean;
    increasedScope?: boolean;
  };
  azure?: {
    username: string;
    displayName?: string;
  };
}

export class SessionUserProperties {
  private _sessionUserProperties: any;

  constructor(sessionEntityReference: any) {
    this._sessionUserProperties = sessionEntityReference;
  }

  getValue(keyPath: string): string {
    return objectPath.get(this._sessionUserProperties, keyPath);
  }

  setValue(keyPath: string, value: string): boolean {
    return objectPath.set(this._sessionUserProperties, keyPath, value);
  }
}

export interface IReposGitHubTokens {
  gitHubReadToken: string;
  gitHubWriteOrganizationToken: string;
}

class ReposGitHubTokensSessionAdapter implements IReposGitHubTokens {
  private _sessionUserProperties: SessionUserProperties;

  constructor(sessionUserProperties: SessionUserProperties) {
    this._sessionUserProperties = sessionUserProperties;
  }

  get gitHubReadToken(): string {
    return this._sessionUserProperties.getValue('github.accessToken');
  }

  get gitHubWriteOrganizationToken(): string {
    const githubModernScope = this._sessionUserProperties.getValue(
      'github.scope'
    );
    // The newer GitHub App model supports user-to-server requests that should
    // be equivalent [once GitHub fixes some bugs]. Since GitHub App OAuth
    // does not have a scope, the user's primary token is the only thing to
    // return here.
    if (githubModernScope && githubModernScope === 'githubapp') {
      return this.gitHubReadToken;
    }
    return this._sessionUserProperties.getValue(
      'githubIncreasedScope.accessToken'
    );
  }
}

export class WebApiContext {
  constructor() {}
}

class PugPlugins {
  // This is used to alter the pug runtime environment. We use it today
  // to include conditional Pug files that may or may not exist, without
  // the typical error that Pug provides.

  private static _instance: PugPlugins = null;

  public static GetInstance(providers: IProviders) {
    if (!PugPlugins._instance) {
      PugPlugins._instance = new PugPlugins(providers);
    }
    return PugPlugins._instance;
  }

  _providers: IProviders;
  _plugins: any[];
  _analyzedCorporatePaths: Map<string, boolean>;

  constructor(providers: IProviders) {
    this._providers = providers;
    this._analyzedCorporatePaths = new Map();
    this._plugins = this.createPlugins();
  }

  get plugins() {
    return this._plugins;
  }

  private createPlugins() {
    if (
      !this._providers.corporateViews ||
      Object.getOwnPropertyNames(this._providers.corporateViews).length === 0
    ) {
      return [];
    }
    const analyzedCorporatePaths = this._analyzedCorporatePaths;
    const emptyFileContents = '';
    return [
      {
        read: function (filename, loadOptions) {
          const isPossibleCorporateView = filename.includes('corporate');
          if (isPossibleCorporateView) {
            if (analyzedCorporatePaths.get(filename) === false) {
              return emptyFileContents;
            } else if (analyzedCorporatePaths.get(filename) === true) {
              // no-op: this is a known-good file
            } else {
              try {
                fs.statSync(filename);
                analyzedCorporatePaths.set(filename, true);
                debug(`corporate view ${filename} validated to exist`);
              } catch (fileExists) {
                // This is a corporate view file that does not exist.
                // Instead of causing an error, this returns essentially
                // an empty file.
                analyzedCorporatePaths.set(filename, false);
                debug(
                  `corporate view ${filename} is not present in the application view folders, using an empty file`
                );
                return emptyFileContents;
              }
            }
          }
          return pugLoad.read(filename, loadOptions);
        },
      },
    ];
  }
}

export class WebContext {
  private _baseUrl: string;
  private _request: ReposAppRequest;
  private _response: IReposAppResponse;
  private _sessionUserProperties: SessionUserProperties;
  private _tokens: ReposGitHubTokensSessionAdapter;

  constructor(options: IWebContextOptions) {
    this._baseUrl = options.baseUrl || '/';
    this._request = options.request;
    this._response = options.response;
    this._sessionUserProperties = options.sessionUserProperties;

    this._tokens = new ReposGitHubTokensSessionAdapter(
      this._sessionUserProperties
    );
  }

  get baseUrl(): string {
    return this._baseUrl;
  }

  get tokens(): ReposGitHubTokensSessionAdapter {
    return this._tokens;
  }

  get correlationId(): string {
    const correlationId = this._request.correlationId;
    return correlationId;
  }

  getAbsoluteUrl(relative: string): string {
    const displayHostname = this._request.hostname;
    const approvalScheme = displayHostname === 'localhost' ? 'http' : 'https';
    const slashPrefix = relative.startsWith('/') ? '' : '/';
    return `${approvalScheme}://${displayHostname}${slashPrefix}${relative}`;
  }

  pushBreadcrumb(title: string, optionalLink?: string | boolean): void {
    const req = this._request;
    addBreadcrumb(req, title, optionalLink);
  }

  // NOTE: This function is direct from the legacy provider... it could move to
  // a dedicated alert provider or something else in the future.
  saveUserAlert(
    message: string,
    title: string,
    context: UserAlertType,
    optionalLink?,
    optionalCaption?
  ) {
    if (typeof message !== 'string') {
      console.warn(
        'First parameter message should be a string, not an object. Was the request object passed through by accident?'
      );
      throw new Error(
        'First parameter message should be a string, not an object. Was the request object passed through by accident?'
      );
    }
    // ----------------------------------------------------------------------------
    // Helper function for UI: Store in the user's session an alert message or
    // action to be shown in another successful render. Contexts come from Twitter
    // Bootstrap, i.e. 'success', 'info', 'warning', 'danger'.
    // ----------------------------------------------------------------------------
    const alert = {
      message,
      title: title || 'FYI',
      context: context || UserAlertType.Success,
      optionalLink: optionalLink,
      optionalCaption: optionalCaption,
    };
    const session = this._request['session'] as IAppSession;
    if (session) {
      if (session.alerts && session.alerts.length) {
        session.alerts.push(alert);
      } else {
        session.alerts = [alert];
      }
    }
  }

  render(options: IWebPageRenderOptions) {
    if (!this._request) {
      throw new Error('No request available');
    }
    if (!this._response) {
      throw new Error('No request available');
    }

    const individualContext = this._request.individualContext;

    const { view, title, optionalObject, state } = options;

    let viewState = state || optionalObject;
    if (state && optionalObject) {
      throw new Error(
        'Both state and optionalObject cannot be provided to a view render method'
      );
    }

    // LEGACY: this whole section
    const breadcrumbs = this._request['breadcrumbs'];
    if (breadcrumbs && breadcrumbs.length && breadcrumbs.length > 0) {
      breadcrumbs[breadcrumbs.length - 1].isLast = true;
    }
    const authScheme = 'aad';
    const user: IWebPageRenderUser = {
      primaryAuthenticationScheme: authScheme,
      primaryUsername: individualContext.corporateIdentity
        ? individualContext.corporateIdentity.username
        : null,
      githubSignout: '/signout/github',
      azureSignout: '/signout',
    };
    // TODO: if the user hasn't linked, we need to access their session/individual context's github identity here!
    const gitHubIdentity = individualContext.getGitHubIdentity();
    if (gitHubIdentity) {
      user.github = {
        id: gitHubIdentity.id,
        username: gitHubIdentity.username,
        displayName: gitHubIdentity.displayName,
        avatarUrl: gitHubIdentity.avatar,
        // OLD: accessToken; this is no longer stored
        increasedScope: individualContext.hasGitHubOrganizationWriteToken(),
      };
    }
    if (individualContext.corporateIdentity) {
      user.azure = {
        username: individualContext.corporateIdentity.username,
        displayName: individualContext.corporateIdentity.displayName,
      };
    }
    const reposContext = this._request.reposContext || {
      section: 'orgs',
      organization: this._request.organization,
    };
    const config = this._request.app.settings['runtimeConfig'];
    if (!config) {
      throw new Error('runtimeConfig is missing');
    }
    const simulatedLegacyLink = individualContext.link
      ? {
          aadupn: user.azure ? user.azure.username : null,
          ghu: user.github ? user.github.username : null,
        }
      : null;
    let session = this._request['session'] || null;

    const initialViewObject = individualContext
      ? individualContext.getInitialViewObject()
      : {};

    const providers = this._request.app.settings.providers as IProviders;
    const { corporateViews } = providers;
    const plugins = PugPlugins.GetInstance(providers).plugins;
    const obj = Object.assign(initialViewObject, {
      title,
      config,
      corporateViews,
      plugins,
      serviceBanner: config.serviceMessage
        ? config.serviceMessage.banner
        : null,
      user,
      // DESTROY: CONFIRM once 'ossline' is gone this way
      ossLink: simulatedLegacyLink,
      showBreadcrumbs: true,
      breadcrumbs,
      sudoMode: this._request['sudoMode'],
      view,
      site: 'github',
      enableMultipleAccounts: session
        ? session['enableMultipleAccounts']
        : false,
      reposContext: undefined,
      alerts: undefined,
    });
    if (obj.ossLink && reposContext) {
      obj.reposContext = reposContext;
    }
    if (viewState) {
      Object.assign(obj, viewState);
    }
    if (session && session['alerts'] && session['alerts'].length) {
      const alerts = [];
      Object.assign(alerts, session['alerts']);
      session['alerts'] = [];
      for (let i = 0; i < alerts.length; i++) {
        if (typeof alerts[i] == 'object') {
          alerts[i].number = i + 1;
        }
      }
      obj.alerts = alerts;
    }
    debug(`web render: view=${options.view}`);
    return this._response.render(view, obj);
    // ANCIENT: RESTORE A GOOD CALL HERE!
    /*
    if (reposContext && !reposContext.availableOrganizations) {
      this.getMyOrganizations((getMyOrgsError, organizations) => {
        if (!getMyOrgsError && organizations && Array.isArray(organizations)) {
          reposContext.availableOrganizations = organizations;
          res.render(view, obj);
        }
      });
    } else {
      res.render(view, obj);
    }
    */
  }
}

// - - - individual context

export interface IIndividualContextOptions {
  corporateIdentity: ICorporateIdentity;
  link: ICorporateLink | null | undefined;
  insights: any;
  webApiContext: WebApiContext | null | undefined;
  webContext: WebContext | null | undefined;
  operations: Operations;
}

export class IndividualContext {
  private _corporateIdentity: ICorporateIdentity;
  private _sessionBasedGitHubIdentity: IGitHubIdentity;
  private _link: ICorporateLink;
  private _webContext: WebContext;
  private _isPortalAdministrator: boolean | null;
  private _operations: Operations;
  private _aggregations: UserContext;
  private _initialView: IDictionary<any>;

  constructor(options: IIndividualContextOptions) {
    this._initialView = {};
    this._isPortalAdministrator = null;
    this._corporateIdentity = options.corporateIdentity;
    this._link = options.link;
    this._webContext = options.webContext;
    this._operations = options.operations;
  }

  get corporateIdentity(): ICorporateIdentity {
    return this._corporateIdentity;
  }

  set corporateIdentity(value: ICorporateIdentity) {
    if (this._corporateIdentity) {
      throw new Error('The context already has a corporate identity set');
    }
    this._corporateIdentity = value;
  }

  get link(): ICorporateLink {
    return this._link;
  }

  set link(value: ICorporateLink) {
    if (this._link) {
      throw new Error('The context already has had a link set');
    }
    this._link = value;
  }

  get webContext(): WebContext {
    return this._webContext;
  }

  hasGitHubOrganizationWriteToken(): boolean {
    const hasToken = !!this.webContext?.tokens?.gitHubWriteOrganizationToken;
    return hasToken;
  }

  get aggregations(): UserContext {
    if (this._aggregations) {
      return this._aggregations;
    }
    this._aggregations = new UserContext(
      this._operations,
      this._operations.providers.queryCache,
      Number(this.getGitHubIdentity().id)
    );
    return this._aggregations;
  }

  getGitHubIdentity(): IGitHubIdentity {
    if (this._link) {
      return {
        id: this._link.thirdPartyId,
        username: this._link.thirdPartyUsername,
        avatar: this._link.thirdPartyAvatar,
        source: GitHubIdentitySource.Link,
      };
    } else if (this._sessionBasedGitHubIdentity) {
      return this._sessionBasedGitHubIdentity;
    }
    return null;
  }

  getSessionBasedGitHubIdentity() {
    return this._sessionBasedGitHubIdentity;
  }

  setSessionBasedGitHubIdentity(identity: IGitHubIdentity) {
    this._sessionBasedGitHubIdentity = identity;
  }

  createGitHubLinkObject(): ICorporateLink {
    const corporateIdentity = this._corporateIdentity;
    if (!corporateIdentity) {
      throw new Error('Cannot create a link: no corporate identity');
    }

    const gitHubIdentity = this.getGitHubIdentity();
    if (!gitHubIdentity) {
      throw new Error('Cannot create a link: no corporate identity');
    }

    const newLink: ICorporateLink = {
      thirdPartyAvatar: gitHubIdentity.avatar,
      thirdPartyId: gitHubIdentity.id,
      thirdPartyUsername: gitHubIdentity.username,
      corporateId: corporateIdentity.id,
      corporateUsername: corporateIdentity.username,
      corporateDisplayName: corporateIdentity.displayName,
      corporateMailAddress: null,
      corporateAlias: null,
      isServiceAccount: false,
      serviceAccountMail: undefined,
    };
    return newLink;
  }

  async isPortalAdministrator(): Promise<boolean> {
    const operations = this._operations;
    const ghi = this.getGitHubIdentity().username;
    const link = this._link;
    this._isPortalAdministrator = await operations.isPortalSudoer(ghi, link);
    return this._isPortalAdministrator;
  }

  setInitialViewProperty(propertyName: string, value: any) {
    this._initialView[propertyName] = value;
  }

  getInitialViewObject() {
    return Object.assign({}, this._initialView);
  }
}
