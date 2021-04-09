//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IReposApplication } from '../interfaces';

export enum AppPurpose {
  Data = 'Data',
  CustomerFacing = 'CustomerFacing',
  Operations = 'Operations',
  BackgroundJobs = 'BackgroundJobs', // "secondary" / "default" fallback
  Updates = 'Updates',
}

export enum GitHubAppAuthenticationType {
  ForceSpecificInstallation = 'force',
  BestAvailable = 'best',
}

export interface IGitHubAppConfiguration {
  clientId?: string;
  clientSecret?: string;
  appId?: number;
  appKey?: string;
  appKeyFile?: string;
  webhookSecret?: string;
  slug?: string;
  description?: string;
  baseUrl: string;
}

export interface IGitHubAppsOptions {
  backgroundJobs?: IGitHubAppConfiguration;
  dataApp?: IGitHubAppConfiguration;
  customerFacingApp?: IGitHubAppConfiguration;
  operationsApp?: IGitHubAppConfiguration;
  updatesApp?: IGitHubAppConfiguration;
  app: IReposApplication;
}
