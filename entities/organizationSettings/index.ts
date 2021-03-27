//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IOrganizationSettingProvider,
  IOrganizationSettingCreateOptions,
  OrganizationSettingProvider,
} from './organizationSettingProvider';

export async function createAndInitializeOrganizationSettingProviderInstance(
  options?: IOrganizationSettingCreateOptions
): Promise<IOrganizationSettingProvider> {
  const provider = new OrganizationSettingProvider(options);
  await provider.initialize();
  return provider;
}
