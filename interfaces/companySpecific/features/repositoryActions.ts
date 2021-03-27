//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Repository } from '../../../business';
import { IProviders, LocalApiRepoAction } from '../../../transitional';
import { IndividualContext } from '../../../user';

export interface ICompanySpecificRepositoryStateStatus {}

export interface ICompanySpecificFeatureRepositoryState {
  getCurrentRepositoryState(
    providers: IProviders,
    repository: Repository
  ): Promise<ICompanySpecificRepositoryStateStatus>;
  sendActionReceipt(
    providers: IProviders,
    context: IndividualContext,
    repository: Repository,
    action: LocalApiRepoAction,
    currentState: ICompanySpecificRepositoryStateStatus
  ): Promise<void>;
}
