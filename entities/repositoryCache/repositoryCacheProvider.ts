//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IEntityMetadata,
  EntityMetadataBase,
  IEntityMetadataBaseOptions,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  RepositoryCacheEntity,
  EntityImplementation,
  RepositoryCacheFixedQueryAll,
  RepositoryCacheFixedQueryByOrganizationId,
  RepositoryCacheDeleteByOrganizationId,
  RepositoryCacheGetOrganizationIdsQuery,
} from './repositoryCache';

const thisProviderType = EntityImplementation.Type;

export interface IRepositoryCacheCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IRepositoryCacheProvider {
  initialize(): Promise<void>;

  getRepository(repositoryId: string): Promise<RepositoryCacheEntity>;
  createRepositoryCache(metadata: RepositoryCacheEntity): Promise<string>;
  updateRepositoryCache(metadata: RepositoryCacheEntity): Promise<void>;
  deleteRepositoryCache(metadata: RepositoryCacheEntity): Promise<void>;
  queryAllRepositories(): Promise<RepositoryCacheEntity[]>;
  queryRepositoriesByOrganizationId(
    organizationId: string
  ): Promise<RepositoryCacheEntity[]>;
  queryAllOrganizationIds(): Promise<string[]>;
  deleteByOrganizationId(organizationId: string): Promise<void>;
}

export class RepositoryCacheProvider
  extends EntityMetadataBase
  implements IRepositoryCacheProvider {
  constructor(options: IRepositoryCacheCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getRepository(repositoryId: string): Promise<RepositoryCacheEntity> {
    this.ensureHelpers(thisProviderType);
    let metadata: IEntityMetadata = null;
    if (this._entities.supportsPointQueryForType(thisProviderType)) {
      metadata = await this._entities.getMetadata(
        thisProviderType,
        repositoryId
      );
    } else {
      throw new Error(
        'fixed point queries are required as currently implemented'
      );
    }
    if (!metadata) {
      const error = new Error(
        `No metadata available for repository ${repositoryId}`
      );
      error['status'] = 404;
      throw error;
    }
    return this.deserialize<RepositoryCacheEntity>(thisProviderType, metadata);
  }

  async queryAllRepositories(): Promise<RepositoryCacheEntity[]> {
    const query = new RepositoryCacheFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryRepositoriesByOrganizationId(
    organizationId: string
  ): Promise<RepositoryCacheEntity[]> {
    const query = new RepositoryCacheFixedQueryByOrganizationId(organizationId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createRepositoryCache(
    metadata: RepositoryCacheEntity
  ): Promise<string> {
    const entity = this.serialize(thisProviderType, metadata);
    if (!this._entities.supportsPointQueryForType(thisProviderType)) {
      throw new Error(
        'fixed point queries are required as currently implemented'
      );
    }
    await this._entities.setMetadata(entity);
    return entity.entityId;
  }

  async updateRepositoryCache(metadata: RepositoryCacheEntity): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }

  async deleteRepositoryCache(metadata: RepositoryCacheEntity): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.deleteMetadata(entity);
  }

  async queryAllOrganizationIds(): Promise<string[]> {
    const query = new RepositoryCacheGetOrganizationIdsQuery();
    const results = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    return results.map((row) => row['organizationid']);
  }

  async deleteByOrganizationId(organizationId: string): Promise<void> {
    const query = new RepositoryCacheDeleteByOrganizationId(organizationId);
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }
}
