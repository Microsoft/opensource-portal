//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  RepositoryTeamCacheProvider,
  IRepositoryTeamCacheCreateOptions,
  IRepositoryTeamCacheProvider,
} from './repositoryTeamCacheProvider';

export async function CreateRepositoryTeamCacheProviderInstance(
  options?: IRepositoryTeamCacheCreateOptions
): Promise<IRepositoryTeamCacheProvider> {
  const provider = new RepositoryTeamCacheProvider(options);
  await provider.initialize();
  return provider;
}
