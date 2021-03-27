//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IEntityMetadata,
  EntityMetadataBase,
  IEntityMetadataBaseOptions,
} from '../../lib/entityMetadataProvider/entityMetadata';
import { TeamMemberCacheEntity, EntityImplementation } from './teamMemberCache';
import {
  TeamMemberCacheFixedQueryAll,
  TeamMemberCacheFixedQueryByOrganizationId,
  TeamMemberCacheFixedQueryByUserId,
  TeamMemberCacheFixedQueryByTeamId,
  TeamMemberCacheFixedQueryByOrganizationIdAndUserId,
  TeamMemberCacheGetOrganizationIdsQuery,
  TeamMemberCacheDeleteByOrganizationId,
} from '.';

const thisProviderType = EntityImplementation.Type;

export interface ITeamMemberCacheCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface ITeamMemberCacheProvider {
  initialize(): Promise<void>;

  getTeamMemberCache(uniqueId: string): Promise<TeamMemberCacheEntity>;
  getTeamMemberCacheByUserId(
    organizationId: string,
    teamId: string,
    userId: string
  ): Promise<TeamMemberCacheEntity>;
  createTeamMemberCache(metadata: TeamMemberCacheEntity): Promise<string>;
  updateTeamMemberCache(metadata: TeamMemberCacheEntity): Promise<void>;
  deleteTeamMemberCache(metadata: TeamMemberCacheEntity): Promise<void>;
  queryAllTeamMembers(): Promise<TeamMemberCacheEntity[]>;
  queryTeamMembersByOrganizationId(
    organizationId: string
  ): Promise<TeamMemberCacheEntity[]>;
  queryTeamMembersByUserId(userId: string): Promise<TeamMemberCacheEntity[]>;
  queryTeamMembersByTeamId(teamId: string): Promise<TeamMemberCacheEntity[]>;
  queryTeamMembersByOrganizationIdAndUserId(
    organizationId: string,
    userId: string
  ): Promise<TeamMemberCacheEntity[]>;
  queryAllOrganizationIds(): Promise<string[]>;
  deleteByOrganizationId(organizationId: string): Promise<void>;
}

export class TeamMemberCacheProvider
  extends EntityMetadataBase
  implements ITeamMemberCacheProvider {
  constructor(options: ITeamMemberCacheCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getTeamMemberCacheByUserId(
    organizationId: string,
    teamId: string,
    userId: string
  ): Promise<TeamMemberCacheEntity> {
    return this.getTeamMemberCache(
      TeamMemberCacheEntity.GenerateIdentifier(organizationId, teamId, userId)
    );
  }

  async getTeamMemberCache(uniqueId: string): Promise<TeamMemberCacheEntity> {
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
        `No metadata available for team member ${uniqueId}`
      );
      error['status'] = 404;
      throw error;
    }
    return this.deserialize<TeamMemberCacheEntity>(thisProviderType, metadata);
  }

  async queryAllTeamMembers(): Promise<TeamMemberCacheEntity[]> {
    const query = new TeamMemberCacheFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryTeamMembersByOrganizationId(
    organizationId: string
  ): Promise<TeamMemberCacheEntity[]> {
    const query = new TeamMemberCacheFixedQueryByOrganizationId(organizationId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryTeamMembersByUserId(
    userId: string
  ): Promise<TeamMemberCacheEntity[]> {
    const query = new TeamMemberCacheFixedQueryByUserId(userId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryTeamMembersByOrganizationIdAndUserId(
    organizationId: string,
    userId: string
  ): Promise<TeamMemberCacheEntity[]> {
    const query = new TeamMemberCacheFixedQueryByOrganizationIdAndUserId(
      organizationId,
      userId
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryTeamMembersByTeamId(
    teamId: string
  ): Promise<TeamMemberCacheEntity[]> {
    const query = new TeamMemberCacheFixedQueryByTeamId(teamId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamMemberCacheEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createTeamMemberCache(
    metadata: TeamMemberCacheEntity
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

  async updateTeamMemberCache(metadata: TeamMemberCacheEntity): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }

  async deleteTeamMemberCache(metadata: TeamMemberCacheEntity): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.deleteMetadata(entity);
  }

  async queryAllOrganizationIds(): Promise<string[]> {
    const query = new TeamMemberCacheGetOrganizationIdsQuery();
    const results = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    return results.map((row) => row['organizationid']);
  }

  async deleteByOrganizationId(organizationId: string): Promise<void> {
    const query = new TeamMemberCacheDeleteByOrganizationId(organizationId);
    await this._entities.fixedQueryMetadata(thisProviderType, query);
  }
}
