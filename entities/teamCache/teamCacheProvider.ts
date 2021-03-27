//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IEntityMetadata,
  EntityMetadataBase,
  IEntityMetadataBaseOptions,
} from '../../lib/entityMetadataProvider/entityMetadata';
import { TeamCacheEntity } from './teamCache';
import {
  TeamCacheFixedQueryAll,
  TeamCacheFixedQueryByOrganizationId,
  TeamCacheGetOrganizationIdsQuery,
  TeamCacheDeleteByOrganizationId,
} from '.';
import { EntityImplementation } from './teamCache';

const thisProviderType = EntityImplementation.Type;

export interface ITeamCacheCreateOptions extends IEntityMetadataBaseOptions {}

export interface ITeamCacheProvider {
  initialize(): Promise<void>;

  getTeam(teamId: string): Promise<TeamCacheEntity>;
  createTeamCache(metadata: TeamCacheEntity): Promise<string>;
  updateTeamCache(metadata: TeamCacheEntity): Promise<void>;
  deleteTeamCache(metadata: TeamCacheEntity): Promise<void>;
  queryAllTeams(): Promise<TeamCacheEntity[]>;
  queryTeamsByOrganizationId(
    organizationId: string
  ): Promise<TeamCacheEntity[]>;
  queryAllOrganizationIds(): Promise<string[]>;
  deleteByOrganizationId(organizationId: string): Promise<void>;
}

export class TeamCacheProvider
  extends EntityMetadataBase
  implements ITeamCacheProvider {
  constructor(options: ITeamCacheCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getTeam(teamId: string): Promise<TeamCacheEntity> {
    this.ensureHelpers(thisProviderType);
    let metadata: IEntityMetadata = null;
    if (this._entities.supportsPointQueryForType(thisProviderType)) {
      metadata = await this._entities.getMetadata(thisProviderType, teamId);
    } else {
      throw new Error(
        'fixed point queries are required as currently implemented'
      );
    }
    if (!metadata) {
      const error = new Error(`No metadata available for team ${teamId}`);
      error['status'] = 404;
      throw error;
    }
    return this.deserialize<TeamCacheEntity>(thisProviderType, metadata);
  }

  async queryAllTeams(): Promise<TeamCacheEntity[]> {
    const query = new TeamCacheFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryTeamsByOrganizationId(
    organizationId: string
  ): Promise<TeamCacheEntity[]> {
    const query = new TeamCacheFixedQueryByOrganizationId(organizationId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createTeamCache(metadata: TeamCacheEntity): Promise<string> {
    const entity = this.serialize(thisProviderType, metadata);
    if (!this._entities.supportsPointQueryForType(thisProviderType)) {
      throw new Error(
        'fixed point queries are required as currently implemented'
      );
    }
    await this._entities.setMetadata(entity);
    return entity.entityId;
  }

  async updateTeamCache(metadata: TeamCacheEntity): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }

  async deleteTeamCache(metadata: TeamCacheEntity): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.deleteMetadata(entity);
  }

  async queryAllOrganizationIds(): Promise<string[]> {
    const query = new TeamCacheGetOrganizationIdsQuery();
    const results = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    return results.map((row) => row['organizationid']);
  }

  async deleteByOrganizationId(organizationId: string): Promise<void> {
    const query = new TeamCacheDeleteByOrganizationId(organizationId);
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }
}
