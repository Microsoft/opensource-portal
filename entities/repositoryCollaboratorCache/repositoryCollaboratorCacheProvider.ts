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
  RepositoryCollaboratorCacheEntity,
  EntityImplementation,
  RepositoryCollaboratorCacheFixedQueryAll,
  RepositoryCollaboratorCacheFixedQueryByOrganizationId,
  RepositoryCollaboratorCacheFixedQueryByUserId,
  RepositoryCollaboratorCacheFixedQueryByRepositoryId,
  RepositoryCollaboratorCacheDeleteByOrganizationId,
  RepositoryCollaboratorCacheGetOrganizationIdsQuery,
  RepositoryCollaboratorCacheDeleteByRepositoryId,
} from './repositoryCollaboratorCache';

const thisProviderType = EntityImplementation.Type;

export interface IRepositoryCollaboratorCacheCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IRepositoryCollaboratorCacheProvider {
  initialize(): Promise<void>;

  getRepositoryCollaboratorCache(
    uniqueId: string
  ): Promise<RepositoryCollaboratorCacheEntity>;
  getRepositoryCollaboratorCacheByUserId(
    organizationId: string,
    repositoryId: string,
    userId: string
  ): Promise<RepositoryCollaboratorCacheEntity>;
  createRepositoryCollaboratorCache(
    metadata: RepositoryCollaboratorCacheEntity
  ): Promise<string>;
  updateRepositoryCollaboratorCache(
    metadata: RepositoryCollaboratorCacheEntity
  ): Promise<void>;
  deleteRepositoryCollaboratorCache(
    metadata: RepositoryCollaboratorCacheEntity
  ): Promise<void>;
  queryAllCollaborators(): Promise<RepositoryCollaboratorCacheEntity[]>;
  queryCollaboratorsByOrganizationId(
    organizationId: string
  ): Promise<RepositoryCollaboratorCacheEntity[]>;
  queryCollaboratorsByRepositoryId(
    organizationId: string
  ): Promise<RepositoryCollaboratorCacheEntity[]>;
  queryCollaboratorsByUserId(
    userId: string
  ): Promise<RepositoryCollaboratorCacheEntity[]>;
  queryAllOrganizationIds(): Promise<string[]>;
  deleteByOrganizationId(organizationId: string): Promise<void>;
  deleteByRepositoryId(repositoryId: string): Promise<void>;
}

export class RepositoryCollaboratorCacheProvider
  extends EntityMetadataBase
  implements IRepositoryCollaboratorCacheProvider {
  constructor(options: IRepositoryCollaboratorCacheCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getRepositoryCollaboratorCacheByUserId(
    organizationId: string,
    repositoryId: string,
    userId: string
  ): Promise<RepositoryCollaboratorCacheEntity> {
    return this.getRepositoryCollaboratorCache(
      RepositoryCollaboratorCacheEntity.GenerateIdentifier(
        organizationId,
        repositoryId,
        userId
      )
    );
  }

  async getRepositoryCollaboratorCache(
    uniqueId: string
  ): Promise<RepositoryCollaboratorCacheEntity> {
    this.ensureHelpers(thisProviderType);
    let metadata: IEntityMetadata = null;
    if (this._entities.supportsPointQueryForType(thisProviderType)) {
      metadata = await this._entities.getMetadata(thisProviderType, uniqueId);
    } else {
      throw new Error(
        'fixed point queries are required as currently implemented'
      );
    }
    if (!metadata) {
      const error = new Error(
        `No metadata available for collaborator with unique ID ${uniqueId}`
      );
      error['status'] = 404;
      throw error;
    }
    return this.deserialize<RepositoryCollaboratorCacheEntity>(
      thisProviderType,
      metadata
    );
  }

  async queryAllCollaborators(): Promise<RepositoryCollaboratorCacheEntity[]> {
    const query = new RepositoryCollaboratorCacheFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryCollaboratorCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryCollaboratorsByOrganizationId(
    organizationId: string
  ): Promise<RepositoryCollaboratorCacheEntity[]> {
    const query = new RepositoryCollaboratorCacheFixedQueryByOrganizationId(
      organizationId
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryCollaboratorCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryCollaboratorsByRepositoryId(
    repositoryId: string
  ): Promise<RepositoryCollaboratorCacheEntity[]> {
    const query = new RepositoryCollaboratorCacheFixedQueryByRepositoryId(
      repositoryId
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryCollaboratorCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryCollaboratorsByUserId(
    userId: string
  ): Promise<RepositoryCollaboratorCacheEntity[]> {
    const query = new RepositoryCollaboratorCacheFixedQueryByUserId(userId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryCollaboratorCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createRepositoryCollaboratorCache(
    metadata: RepositoryCollaboratorCacheEntity
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

  async updateRepositoryCollaboratorCache(
    metadata: RepositoryCollaboratorCacheEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }

  async deleteRepositoryCollaboratorCache(
    metadata: RepositoryCollaboratorCacheEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.deleteMetadata(entity);
  }

  async queryAllOrganizationIds(): Promise<string[]> {
    const query = new RepositoryCollaboratorCacheGetOrganizationIdsQuery();
    const results = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    return results.map((row) => row['organizationid']);
  }

  async deleteByOrganizationId(organizationId: string): Promise<void> {
    const query = new RepositoryCollaboratorCacheDeleteByOrganizationId(
      organizationId
    );
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }

  async deleteByRepositoryId(repositoryId: string): Promise<void> {
    const query = new RepositoryCollaboratorCacheDeleteByRepositoryId(
      repositoryId
    );
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }
}
