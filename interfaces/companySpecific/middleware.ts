//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Repository } from '../../business';
import { IContextualRepositoryPermissions } from '../../middleware/github/repoPermissions';
import { IProviders } from '../../transitional';
import { IndividualContext } from '../../user';

export interface ICompanySpecificRepoPermissionsMiddlewareCalls {
  afterPermissionsInitialized?: (
    providers: IProviders,
    permissions: IContextualRepositoryPermissions,
    activeContext: IndividualContext
  ) => void;
  afterPermissionsComputed?: (
    providers: IProviders,
    permissions: IContextualRepositoryPermissions,
    activeContext: IndividualContext,
    repository: Repository
  ) => Promise<void>;
}

export interface IAttachCompanySpecificMiddleware {
  repoPermissions?: ICompanySpecificRepoPermissionsMiddlewareCalls;
}
