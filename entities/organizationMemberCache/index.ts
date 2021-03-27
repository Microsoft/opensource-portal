//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  OrganizationMemberCacheProvider,
  IOrganizationMemberCacheCreateOptions,
  IOrganizationMemberCacheProvider,
} from './organizationMemberCacheProvider';

export async function CreateOrganizationMemberCacheProviderInstance(
  options?: IOrganizationMemberCacheCreateOptions
): Promise<IOrganizationMemberCacheProvider> {
  const provider = new OrganizationMemberCacheProvider(options);
  await provider.initialize();
  return provider;
}
