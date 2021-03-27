//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { EntityMetadataBase } from '../../lib/entityMetadataProvider/entityMetadata';
import { PersonalAccessToken, EntityImplementation } from './token';
import { ITokenProvider, ITokenProviderCreateOptions } from '.';
import {
  IEntityMetadataFixedQuery,
  FixedQueryType,
} from '../../lib/entityMetadataProvider/query';
import { Type } from './type';

const thisProviderType = Type;

export class TokenProvider
  extends EntityMetadataBase
  implements ITokenProvider {
  constructor(options: ITokenProviderCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getToken(token: string): Promise<PersonalAccessToken> {
    this.ensureHelpers(thisProviderType);
    const metadata = await this._entities.getMetadata(thisProviderType, token);
    return this.deserialize<PersonalAccessToken>(thisProviderType, metadata);
  }

  async saveNewToken(token: PersonalAccessToken): Promise<void> {
    const entity = this.serialize(thisProviderType, token);
    await this._entities.setMetadata(entity);
  }

  async updateToken(token: PersonalAccessToken): Promise<void> {
    const entity = this.serialize(thisProviderType, token);
    return await this._entities.updateMetadata(entity);
  }

  async deleteToken(token: PersonalAccessToken): Promise<void> {
    const entity = this.serialize(thisProviderType, token);
    return await this._entities.deleteMetadata(entity);
  }

  async getAllTokens(): Promise<PersonalAccessToken[]> {
    const query = new QueryTokensGetAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<PersonalAccessToken>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryTokensForCorporateId(
    corporateId: string
  ): Promise<PersonalAccessToken[]> {
    const query = new QueryTokensByCorporateID(corporateId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<PersonalAccessToken>(
      thisProviderType,
      metadatas
    );
    return results;
  }
}

export class QueryTokensByCorporateID implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TokensByCorporateId;
  constructor(public corporateId: string) {}
}

export class QueryTokensGetAll implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType = FixedQueryType.TokensGetAll;
}
