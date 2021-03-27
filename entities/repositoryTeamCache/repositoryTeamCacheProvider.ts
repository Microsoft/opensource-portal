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
  RepositoryTeamCacheEntity,
  EntityImplementation,
  RepositoryTeamCacheFixedQueryAll,
  RepositoryTeamCacheFixedQueryByOrganizationId,
  RepositoryTeamCacheFixedQueryByTeamId,
  RepositoryTeamCacheFixedQueryByRepositoryId,
  RepositoryTeamCacheFixedQueryByTeamIds,
  RepositoryTeamCacheDeleteByOrganizationId,
  RepositoryTeamCacheGetOrganizationIdsQuery,
  RepositoryTeamCacheDeleteByRepositoryId,
} from './repositoryTeamCache';

const thisProviderType = EntityImplementation.Type;

export interface IRepositoryTeamCacheCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IRepositoryTeamCacheProvider {
  initialize(): Promise<void>;

  getRepositoryTeamCache(uniqueId: string): Promise<RepositoryTeamCacheEntity>;
  getRepositoryTeamCacheByTeamId(
    organizationId: string,
    repositoryId: string,
    teamId: string
  ): Promise<RepositoryTeamCacheEntity>;
  createRepositoryTeamCache(
    metadata: RepositoryTeamCacheEntity
  ): Promise<string>;
  updateRepositoryTeamCache(metadata: RepositoryTeamCacheEntity): Promise<void>;
  deleteRepositoryTeamCache(metadata: RepositoryTeamCacheEntity): Promise<void>;
  queryAllTeams(): Promise<RepositoryTeamCacheEntity[]>;
  queryByRepositoryId(
    repositoryId: string
  ): Promise<RepositoryTeamCacheEntity[]>;
  queryByTeamId(teamId: string): Promise<RepositoryTeamCacheEntity[]>;
  queryByTeamIds(teamIds: string[]): Promise<RepositoryTeamCacheEntity[]>;
  queryAllOrganizationIds(): Promise<string[]>;
  deleteByOrganizationId(organizationId: string): Promise<void>;
  deleteByRepositoryId(repositoryId: string): Promise<void>;
}

export class RepositoryTeamCacheProvider
  extends EntityMetadataBase
  implements IRepositoryTeamCacheProvider {
  constructor(options: IRepositoryTeamCacheCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getRepositoryTeamCacheByTeamId(
    organizationId: string,
    repositoryId: string,
    teamId: string
  ): Promise<RepositoryTeamCacheEntity> {
    return this.getRepositoryTeamCache(
      RepositoryTeamCacheEntity.GenerateIdentifier(
        organizationId,
        repositoryId,
        teamId
      )
    );
  }

  async getRepositoryTeamCache(
    uniqueId: string
  ): Promise<RepositoryTeamCacheEntity> {
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
        `No metadata available for Team with unique ID ${uniqueId}`
      );
      error['status'] = 404;
      throw error;
    }
    return this.deserialize<RepositoryTeamCacheEntity>(
      thisProviderType,
      metadata
    );
  }

  async queryAllTeams(): Promise<RepositoryTeamCacheEntity[]> {
    const query = new RepositoryTeamCacheFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryTeamCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryByRepositoryId(
    repositoryId: string
  ): Promise<RepositoryTeamCacheEntity[]> {
    const query = new RepositoryTeamCacheFixedQueryByRepositoryId(repositoryId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryTeamCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryByTeamId(teamId: string): Promise<RepositoryTeamCacheEntity[]> {
    const query = new RepositoryTeamCacheFixedQueryByTeamId(teamId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryTeamCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryByTeamIds(
    teamIds: string[]
  ): Promise<RepositoryTeamCacheEntity[]> {
    const query = new RepositoryTeamCacheFixedQueryByTeamIds(teamIds);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryTeamCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createRepositoryTeamCache(
    metadata: RepositoryTeamCacheEntity
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

  async updateRepositoryTeamCache(
    metadata: RepositoryTeamCacheEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }

  async deleteRepositoryTeamCache(
    metadata: RepositoryTeamCacheEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.deleteMetadata(entity);
  }

  async queryAllOrganizationIds(): Promise<string[]> {
    const query = new RepositoryTeamCacheGetOrganizationIdsQuery();
    const results = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    return results.map((row) => row['organizationid']);
  }

  async deleteByOrganizationId(organizationId: string): Promise<void> {
    const query = new RepositoryTeamCacheDeleteByOrganizationId(organizationId);
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }

  async deleteByRepositoryId(repositoryId: string): Promise<void> {
    const query = new RepositoryTeamCacheDeleteByRepositoryId(repositoryId);
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }
}
