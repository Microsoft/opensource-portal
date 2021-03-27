//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { EntityMetadataBase } from '../../lib/entityMetadataProvider/entityMetadata';
import { LocalExtensionKey, EntityImplementation } from './localExtensionKey';
import {
  ILocalExtensionKeyProvider,
  ILocalExtensionKeyProviderOptions,
} from '.';
import {
  IEntityMetadataFixedQuery,
  FixedQueryType,
} from '../../lib/entityMetadataProvider/query';

const thisProviderType = EntityImplementation.Type;

export class LocalExtensionKeyProvider
  extends EntityMetadataBase
  implements ILocalExtensionKeyProvider {
  constructor(options: ILocalExtensionKeyProviderOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getForCorporateId(corporateId: string): Promise<LocalExtensionKey> {
    this.ensureHelpers(thisProviderType);
    const metadata = await this._entities.getMetadata(
      thisProviderType,
      corporateId
    );
    return this.deserialize<LocalExtensionKey>(thisProviderType, metadata);
  }

  async getAllKeys(): Promise<LocalExtensionKey[]> {
    const query = new QueryLocalExtensionKeysGetAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<LocalExtensionKey>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createNewForCorporateId(
    localExtensionKey: LocalExtensionKey
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, localExtensionKey);
    await this._entities.setMetadata(entity);
  }

  async updateForCorporateId(
    localExtensionKey: LocalExtensionKey
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, localExtensionKey);
    return await this._entities.updateMetadata(entity);
  }

  async delete(localExtensionKey: LocalExtensionKey): Promise<void> {
    const entity = this.serialize(thisProviderType, localExtensionKey);
    return await this._entities.deleteMetadata(entity);
  }
}

export class QueryLocalExtensionKeysGetAll
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.LocalExtensionKeysGetAll;
}
