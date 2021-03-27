//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IGetAuthorizationHeader,
  ICacheOptions,
  IOperationsGitHubRestLibrary,
  IOperationsDefaultCacheTimes,
} from '../transitional';
import { wrapError } from '../utils';

export interface IGitHubAppInstallationPermissions {
  issues: string; // write, ?
  metadata: string; // read, ...
  administration: string; // write, ...
}

export interface IGitHubWebhookEnterprise {
  id: number;
  slug: string;
  name: string;
  node_id: string;
  avatar_url: string;
  description: string;
  website_url: string;
  html_url: string;
  created_at: string;
  updated_at: string;
}

export interface IGitHubAppInstallation {
  id: number;
  account: {
    login: string;
    id: number;
  };
  app_id: number;
  target_id: number;
  target_type: string;
  permissions: IGitHubAppInstallationPermissions;
  events: string[];
  repository_selection: string;
  created_at: string;
  updated_at: string;
}

const primaryInstallationProperties = [
  'id',
  'account',
  'app_id',
  'target_id',
  'target_type',
  'permissions',
  'events',
];

export default class GitHubApplication {
  constructor(
    private operations: IOperationsGitHubRestLibrary &
      IOperationsDefaultCacheTimes,
    public id: number,
    public slug: string,
    public friendlyName: string,
    private getAuthorizationHeader: IGetAuthorizationHeader
  ) {}

  static PrimaryInstallationProperties = primaryInstallationProperties;

  static filterInstallations(installations: IGitHubAppInstallation[]) {
    return {
      valid: installations.filter(
        (install) =>
          GitHubApplication.isInvalidInstallation(install).length === 0
      ),
      invalid: installations.filter(
        (install) => GitHubApplication.isInvalidInstallation(install).length > 0
      ),
    };
  }

  static isInvalidInstallation(
    installation: IGitHubAppInstallation
  ): string[] /* invalid reasons*/ {
    let invalid: string[] = [];
    if (installation.target_type !== 'Organization') {
      invalid.push(
        `Installation has an unsupported target type of ${installation.target_type}.`
      );
    }
    if (
      installation.repository_selection &&
      installation.repository_selection !== 'all'
    ) {
      invalid.push(
        `This app can only be installed at the organization scope (all repos), please update the settings for the installation.`
      );
    }
    return invalid;
  }

  async getInstallation(
    installationId: number,
    options?: ICacheOptions
  ): Promise<IGitHubAppInstallation> {
    const operations = this.operations;
    const parameters = {
      installation_id: installationId.toString(),
    };
    const cacheOptions = { ...options };
    try {
      const entity = await operations.github.call(
        this.authorize(),
        'apps.getInstallation',
        parameters,
        cacheOptions
      );
      return entity as IGitHubAppInstallation;
    } catch (error) {
      // TODO: 404 vs error
      throw wrapError(
        error,
        `Could not get details about the ${this.id} app installation ID ${installationId}: ${error.message}`
      );
    }
  }

  deleteInstallation(installationId: number): Promise<void> {
    const parameters = {
      installation_id: installationId.toString(),
    };
    return this.operations.github.post(
      this.authorize(),
      'apps.deleteInstallation',
      parameters
    );
  }

  async getInstallations(
    options?: ICacheOptions
  ): Promise<IGitHubAppInstallation[]> {
    options = options || {};
    const operations = this.operations;
    const getAuthorizationHeader = this.getAuthorizationHeader.bind(
      this
    ) as IGetAuthorizationHeader;
    const github = operations.github;
    const caching = {
      maxAgeSeconds:
        options.maxAgeSeconds || operations.defaults.orgRepoDetailsStaleSeconds, // borrowing from another value
      backgroundRefresh: false,
      // pageRequestDelay: options.pageRequestDelay,
    };
    const installations = await github.collections.getAppInstallations(
      getAuthorizationHeader,
      {
        app_id: this.id.toString(),
      },
      caching
    );
    return installations;
  }

  private authorize(): IGetAuthorizationHeader | string {
    const getAuthorizationHeader = this.getAuthorizationHeader.bind(
      this
    ) as IGetAuthorizationHeader;
    return getAuthorizationHeader;
  }
}
