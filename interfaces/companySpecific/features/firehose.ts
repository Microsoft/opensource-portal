//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IGitHubAppInstallation,
  IGitHubWebhookEnterprise,
} from '../../../business';
import { IProviders } from '../../../transitional';

export interface ICompanySpecificFeatureFirehose {
  processWebhook: (
    providers: IProviders,
    body: any,
    eventType: string,
    enterprise: IGitHubWebhookEnterprise,
    installation: IGitHubAppInstallation,
    acknowledgeEvent: () => void
  ) => Promise<boolean>;
}
