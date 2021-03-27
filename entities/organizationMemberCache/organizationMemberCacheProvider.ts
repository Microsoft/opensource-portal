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
  OrganizationMemberCacheEntity,
  EntityImplementation,
  OrganizationMemberCacheFixedQueryAll,
  OrganizationMemberCacheFixedQueryByOrganizationId,
  OrganizationMemberCacheFixedQueryByUserId,
  OrganizationBasicsFixedQuery,
  OrganizationMemberCacheDeleteByOrganizationId,
  OrganizationOwnersQuery,
} from './organizationMemberCache';

const thisProviderType = EntityImplementation.Type;

export interface IOrganizationMemberCacheCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IOrganizationMemberCacheProvider {
  initialize(): Promise<void>;

  getOrganizationMemberCache(
    uniqueId: string
  ): Promise<OrganizationMemberCacheEntity>;
  getOrganizationMemberCacheByUserId(
    organizationId: string,
    userId: string
  ): Promise<OrganizationMemberCacheEntity>;
  createOrganizationMemberCache(
    metadata: OrganizationMemberCacheEntity
  ): Promise<string>;
  updateOrganizationMemberCache(
    metadata: OrganizationMemberCacheEntity
  ): Promise<void>;
  deleteOrganizationMemberCache(
    metadata: OrganizationMemberCacheEntity
  ): Promise<void>;
  queryAllOrganizationMembers(): Promise<OrganizationMemberCacheEntity[]>;
  queryOrganizationMembersByOrganizationId(
    organizationId: string
  ): Promise<OrganizationMemberCacheEntity[]>;
  queryOrganizationMembersByUserId(
    userId: string
  ): Promise<OrganizationMemberCacheEntity[]>;
  queryAllOrganizationIds(): Promise<string[]>;
  queryAllOrganizationOwners(): Promise<OrganizationMemberCacheEntity[]>;
  deleteByOrganizationId(organizationId: string): Promise<void>;
}

export class OrganizationMemberCacheProvider
  extends EntityMetadataBase
  implements IOrganizationMemberCacheProvider {
  constructor(options: IOrganizationMemberCacheCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getOrganizationMemberCacheByUserId(
    organizationId: string,
    userId: string
  ): Promise<OrganizationMemberCacheEntity> {
    return this.getOrganizationMemberCache(
      OrganizationMemberCacheEntity.GenerateIdentifier(organizationId, userId)
    );
  }

  async getOrganizationMemberCache(
    uniqueId: string
  ): Promise<OrganizationMemberCacheEntity> {
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
    return this.deserialize<OrganizationMemberCacheEntity>(
      thisProviderType,
      metadata
    );
  }

  async queryAllOrganizationMembers(): Promise<
    OrganizationMemberCacheEntity[]
  > {
    const query = new OrganizationMemberCacheFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<OrganizationMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryAllOrganizationIds(): Promise<string[]> {
    const query = new OrganizationBasicsFixedQuery();
    const results = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    return results.map((row) => row['organizationid']);
  }

  async deleteByOrganizationId(organizationId: string): Promise<void> {
    const query = new OrganizationMemberCacheDeleteByOrganizationId(
      organizationId
    );
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }

  async queryOrganizationMembersByOrganizationId(
    organizationId: string
  ): Promise<OrganizationMemberCacheEntity[]> {
    const query = new OrganizationMemberCacheFixedQueryByOrganizationId(
      organizationId
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<OrganizationMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryOrganizationMembersByUserId(
    userId: string
  ): Promise<OrganizationMemberCacheEntity[]> {
    const query = new OrganizationMemberCacheFixedQueryByUserId(userId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<OrganizationMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryAllOrganizationOwners(): Promise<OrganizationMemberCacheEntity[]> {
    const query = new OrganizationOwnersQuery();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<OrganizationMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createOrganizationMemberCache(
    metadata: OrganizationMemberCacheEntity
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

  async updateOrganizationMemberCache(
    metadata: OrganizationMemberCacheEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }

  async deleteOrganizationMemberCache(
    metadata: OrganizationMemberCacheEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.deleteMetadata(entity);
  }
}
