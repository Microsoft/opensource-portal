//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IRepositoryMetadataProvider,
  IRepositoryMetadataCreateOptions,
  RepositoryMetadataApprovalProvider,
} from './repositoryMetadataProvider';

export async function createAndInitializeRepositoryMetadataProviderInstance(
  options?: IRepositoryMetadataCreateOptions
): Promise<IRepositoryMetadataProvider> {
  const provider = new RepositoryMetadataApprovalProvider(options);
  await provider.initialize();
  return provider;
}
