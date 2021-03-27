//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IRepositoryCacheProvider,
  IRepositoryCacheCreateOptions,
  RepositoryCacheProvider,
} from './repositoryCacheProvider';

export async function CreateRepositoryCacheProviderInstance(
  options?: IRepositoryCacheCreateOptions
): Promise<IRepositoryCacheProvider> {
  const provider = new RepositoryCacheProvider(options);
  await provider.initialize();
  return provider;
}
