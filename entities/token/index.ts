//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { TokenProvider } from './tokenProvider';
import { PersonalAccessToken } from './token';
import { IEntityMetadataBaseOptions } from '../../lib/entityMetadataProvider/entityMetadata';

export interface ITokenProvider {
  initialize(): Promise<void>;

  getToken(token: string): Promise<PersonalAccessToken>;
  saveNewToken(token: PersonalAccessToken): Promise<void>;
  updateToken(token: PersonalAccessToken): Promise<void>;
  deleteToken(token: PersonalAccessToken): Promise<void>;

  queryTokensForCorporateId(
    thirdPartyId: string
  ): Promise<PersonalAccessToken[]>;
  getAllTokens(): Promise<PersonalAccessToken[]>;
}

export interface ITokenProviderCreateOptions
  extends IEntityMetadataBaseOptions {}

export async function createTokenProvider(
  options: ITokenProviderCreateOptions
): Promise<ITokenProvider> {
  const provider = new TokenProvider(options);
  await provider.initialize();
  return provider;
}
