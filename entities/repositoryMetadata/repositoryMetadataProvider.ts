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
  RepositoryMetadataEntity,
  RepositoryMetadataFixedQueryAll,
  RepositoryMetadataFixedQueryByRepositoryId,
} from './repositoryMetadata';
import { EntityImplementation } from './repositoryMetadata';

const thisProviderType = EntityImplementation.Type;

export interface IRepositoryMetadataCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IRepositoryMetadataProvider {
  initialize(): Promise<void>;

  getRepositoryMetadata(
    repositoryId: string
  ): Promise<RepositoryMetadataEntity>;
  createRepositoryMetadata(metadata: RepositoryMetadataEntity): Promise<string>;
  updateRepositoryMetadata(metadata: RepositoryMetadataEntity): Promise<void>;
  queryAllRepositoryMetadatas(): Promise<RepositoryMetadataEntity[]>;
  clearAllRepositoryMetadatas(): Promise<void>;
}

export class RepositoryMetadataApprovalProvider
  extends EntityMetadataBase
  implements IRepositoryMetadataProvider {
  constructor(options: IRepositoryMetadataCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getRepositoryMetadata(
    repositoryId: string
  ): Promise<RepositoryMetadataEntity> {
    this.ensureHelpers(thisProviderType);
    let metadata: IEntityMetadata = null;
    if (this._entities.supportsPointQueryForType(thisProviderType)) {
      metadata = await this._entities.getMetadata(
        thisProviderType,
        repositoryId
      );
    } else {
      const query = new RepositoryMetadataFixedQueryByRepositoryId(
        repositoryId
      );
      const metadatas = await this._entities.fixedQueryMetadata(
        thisProviderType,
        query
      );
      if (metadatas.length > 1) {
        const error = new Error(
          `Only a single metadata result was expected for repository ${repositoryId}`
        );
        error['status'] = 409;
        throw error;
      }
      metadata = metadatas.length === 1 ? metadatas[0] : null;
    }
    if (!metadata) {
      const error = new Error(
        `No metadata available for repository ${repositoryId}`
      );
      error['status'] = 404;
      throw error;
    }
    return this.deserialize<RepositoryMetadataEntity>(
      thisProviderType,
      metadata
    );
  }

  async clearAllRepositoryMetadatas(): Promise<void> {
    await this._entities.clearMetadataStore(thisProviderType);
  }

  async queryAllRepositoryMetadatas(): Promise<RepositoryMetadataEntity[]> {
    const query = new RepositoryMetadataFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<RepositoryMetadataEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createRepositoryMetadata(
    metadata: RepositoryMetadataEntity
  ): Promise<string> {
    const repoId = metadata.repositoryId;
    const entity = this.serialize(thisProviderType, metadata);
    if (!this._entities.supportsPointQueryForType(thisProviderType)) {
      let exists = false;
      try {
        await this.getRepositoryMetadata(repoId);
        exists = true;
      } catch (getError) {
        if (getError['status'] === 404) {
          // ok
        } else if (getError['status'] === 409) {
          exists = true;
        } else {
          throw getError;
        }
      }
      if (exists) {
        // TODO: which storage error code would this be?
        throw new Error(
          `There is already an entity of type ${thisProviderType} in the table for ID ${repoId}`
        );
      }
    }
    await this._entities.setMetadata(entity);
    return entity.entityId;
  }

  async updateRepositoryMetadata(
    metadata: RepositoryMetadataEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }
}
