//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  AppPurpose,
  IGitHubAppConfiguration,
  IGitHubAppsOptions,
  GitHubAppAuthenticationType,
} from '.';
import { GitHubAppTokens } from './appTokens';
import { IAuthorizationHeaderValue } from '../transitional';
import { OrganizationSetting } from '../entities/organizationSettings/organizationSetting';
import { readFileToText } from '../utils';

const fallbackPurposePriorities = [
  AppPurpose.Data,
  AppPurpose.CustomerFacing,
  AppPurpose.Operations,
  AppPurpose.BackgroundJobs,
];

const fallbackBackgroundJobPriorities = [
  AppPurpose.BackgroundJobs,
  AppPurpose.Data,
  AppPurpose.CustomerFacing,
  AppPurpose.Operations,
];

// Installation redirect format:
// /setup/app/41051?installation_id=1914154&setup_action=install

// 1: check if the app id is known
// 2: request installation info as the app
// 3: configure org setting/s, or add the org!

interface InstallationIdPurposePair {
  purpose: AppPurpose;
  installationId: number;
}

export class GitHubTokenManager {
  #options: IGitHubAppsOptions;
  private static _isBackgroundJob: boolean;
  // private _appConfiguration = new Map<AppPurpose, IGitHubAppConfiguration>();
  private _apps = new Map<AppPurpose, GitHubAppTokens>();
  private _appsById = new Map<number, GitHubAppTokens>();
  private _appIdToPurpose = new Map<number, AppPurpose>();
  private _appSlugs = new Map<number, string>();

  constructor(options: IGitHubAppsOptions) {
    if (!options) {
      throw new Error('options required');
    }
    this.#options = options;
    GitHubTokenManager._isBackgroundJob = options.app.isBackgroundJob;
  }

  async initialize() {
    await this.initializeApp(
      AppPurpose.CustomerFacing,
      this.#options.customerFacingApp
    );
    await this.initializeApp(
      AppPurpose.Operations,
      this.#options.operationsApp
    );
    await this.initializeApp(AppPurpose.Data, this.#options.dataApp);
    await this.initializeApp(
      AppPurpose.BackgroundJobs,
      this.#options.backgroundJobs
    );
    await this.initializeApp(AppPurpose.Updates, this.#options.updatesApp);
  }

  organizationSupportsAnyPurpose(
    organizationName: string,
    organizationSettings?: OrganizationSetting
  ) {
    return !!this.getPrioritizedOrganizationInstallationId(
      fallbackPurposePriorities[0],
      organizationName,
      organizationSettings,
      GitHubAppAuthenticationType.BestAvailable
    );
  }

  getAppById(id: number): GitHubAppTokens {
    return this._appsById.get(id);
  }

  getAppIds(): number[] {
    return Array.from(this._appsById.keys());
  }

  getSlugById(id: number): string {
    return this._appSlugs.get(id);
  }

  async getOrganizationAuthorizationHeader(
    organizationName: string,
    preferredPurpose: AppPurpose,
    organizationSettings: OrganizationSetting,
    appAuthenticationType: GitHubAppAuthenticationType
  ): Promise<IAuthorizationHeaderValue> {
    const installationIdPair = this.getPrioritizedOrganizationInstallationId(
      preferredPurpose,
      organizationName,
      organizationSettings,
      appAuthenticationType
    );
    if (!installationIdPair) {
      throw new Error(
        `GitHubTokenManager: organization ${organizationName} does not support the GitHub App model`
      );
    }
    if (
      appAuthenticationType === GitHubAppAuthenticationType.BestAvailable &&
      installationIdPair.purpose !== preferredPurpose
    ) {
      // console.log(`preferred GitHub App type ${preferredPurpose} not configured for organization ${organizationName}, falling back to ${installationIdPair.purpose}`);
    }
    const app = this._apps.get(installationIdPair.purpose);
    return app.getInstallationToken(
      installationIdPair.installationId,
      organizationName
    );
  }

  async getInstallationAuthorizationHeader(
    appId: number,
    installationId: number,
    organizationName: string
  ): Promise<IAuthorizationHeaderValue> {
    const app = this._appsById.get(appId);
    if (!app) {
      throw new Error(
        `App ID=${appId} is not configured in this application instance`
      );
    }
    return app.getInstallationToken(installationId, organizationName);
  }

  private getPrioritizedOrganizationInstallationId(
    preferredPurpose: AppPurpose,
    organizationName: string,
    organizationSettings: OrganizationSetting,
    appAuthenticationType: GitHubAppAuthenticationType
  ): InstallationIdPurposePair {
    if (!organizationSettings) {
      return null;
    }
    let order =
      GitHubTokenManager._isBackgroundJob === true
        ? fallbackBackgroundJobPriorities
        : [preferredPurpose, ...fallbackPurposePriorities];
    if (
      appAuthenticationType ===
      GitHubAppAuthenticationType.ForceSpecificInstallation
    ) {
      order = [preferredPurpose];
    }
    for (const purpose of order) {
      if (organizationSettings && organizationSettings.installations) {
        for (const {
          appId,
          installationId,
        } of organizationSettings.installations) {
          const purposeForApp = this._appIdToPurpose.get(appId);
          if (
            this._appsById.has(appId) &&
            purposeForApp &&
            purposeForApp === purpose
          ) {
            return { installationId, purpose };
          }
        }
      }
    }
  }

  private async initializeApp(
    purpose: AppPurpose,
    appConfig: IGitHubAppConfiguration
  ) {
    if (!appConfig || !appConfig.appId) {
      return;
    }
    const appId =
      typeof appConfig.appId === 'number'
        ? appConfig.appId
        : parseInt(appConfig.appId, 10);
    if (this._appsById.has(appId)) {
      return;
    }
    let key = appConfig.appKey;
    let fromLocalFile = false;
    if (appConfig.appKeyFile) {
      key = await readFileToText(appConfig.appKeyFile);
      fromLocalFile = true;
    }
    if (!key) {
      throw new Error(
        `appKey or appKeyFile required for ${purpose} GitHub App configuration`
      );
    }
    const friendlyName = appConfig.description || 'Unknown';
    const baseUrl = appConfig.baseUrl;
    const app = fromLocalFile
      ? GitHubAppTokens.CreateFromString(
          purpose,
          friendlyName,
          appId,
          key,
          baseUrl
        )
      : GitHubAppTokens.CreateFromBase64EncodedFileString(
          purpose,
          friendlyName,
          appId,
          key,
          baseUrl
        );
    this._apps.set(purpose, app);
    this._appsById.set(appId, app);
    this._appSlugs.set(appId, appConfig.slug);
    this._appIdToPurpose.set(appId, purpose);
  }
}
