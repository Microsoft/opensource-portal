//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IEntityMetadataProvider, IEntityMetadataSerializationHelper, IEntityMetadataDeserializationHelper } from "./entityMetadataProvider";

export enum EntityMetadataType {
  // READ THIS:
  // IMPORTANT: When adding new types, also add it to EntityMetadataTypes array below!

  // GitHub entities METADATA
  Repository = 'Repository',

  AuditLogRecord = 'AuditLogRecord',
  EventRecord = 'EventRecord',

  // App-specific entities
  TeamJoinRequest = 'TeamJoinRequest',

  // Fast query cache entities
  OrganizationMemberCache = 'OrganizationMemberCache',
  RepositoryCache = 'RepositoryCache',
  RepositoryCollaboratorCache = 'RepositoryCollaboratorCache',
  RepositoryTeamCache = 'RepositoryTeamCache',
  TeamCache = 'TeamCache',
  TeamMemberCache = 'TeamMemberCache',

  // Setting entities
  Token = 'Token',
  LocalExtensionKey = 'LocalExtensionKey',
  OrganizationSetting = 'OrganizationSetting',
}

export const EntityMetadataTypes = [
  EntityMetadataType.Repository,
  EntityMetadataType.AuditLogRecord,
  EntityMetadataType.EventRecord,
  EntityMetadataType.TeamJoinRequest,
  EntityMetadataType.Token,
  EntityMetadataType.LocalExtensionKey,
  EntityMetadataType.OrganizationMemberCache,
  EntityMetadataType.OrganizationSetting,
  EntityMetadataType.RepositoryCache,
  EntityMetadataType.RepositoryCollaboratorCache,
  EntityMetadataType.RepositoryTeamCache,
  EntityMetadataType.TeamCache,
  EntityMetadataType.TeamMemberCache,
];

export interface IEntityMetadata {
  entityType: EntityMetadataType;
  entityId: string;
  entityFieldNames: string[];
  entityCreated?: Date;
}

export interface IEntityMetadataBaseOptions {
  entityMetadataProvider: IEntityMetadataProvider;
}

export abstract class EntityMetadataBase {
  protected _entities: IEntityMetadataProvider;
  protected _serialize: Map<EntityMetadataType, IEntityMetadataSerializationHelper>;
  protected _deserialize: Map<EntityMetadataType, IEntityMetadataDeserializationHelper>;

  constructor(options: IEntityMetadataBaseOptions) {
    this._entities = options.entityMetadataProvider;
  }

  async initialize(): Promise<void> {}

  protected serialize(type: EntityMetadataType, obj: any): IEntityMetadata {
    this.ensureHelpers(type);
    const serializer = this._serialize.get(type);
    const metadata = serializer(obj);
    return metadata;
  }

  protected deserialize<T>(type: EntityMetadataType, metadata: IEntityMetadata) {
    this.ensureHelpers(type);
    const entity = this._deserialize.get(type)(metadata) as T;
    return entity;
  }

  protected deserializeArray<T>(type: EntityMetadataType, array: IEntityMetadata[]): T[] {
    return array.map(metadata => this.deserialize(type, metadata));
  }

  protected ensureHelpers(type: EntityMetadataType) {
    if (!this._serialize) {
      this._serialize = new Map<EntityMetadataType, IEntityMetadataSerializationHelper>();
    }
    if (!this._serialize.has(type)) {
      const helper = this._entities.getSerializationHelper(type);
      if (!helper) {
        throw new Error(`No serialization helper available to the ${this._entities.name} entity provider for the type ${type}`);
      }
      this._serialize.set(type, helper);
    }
    if (!this._deserialize) {
      this._deserialize = new Map<EntityMetadataType, IEntityMetadataDeserializationHelper>();
    }
    if (!this._deserialize.has(type)) {
      const helper = this._entities.getDeserializationHelper(type);
      if (!helper) {
        throw new Error(`No deserialization helper available to the ${this._entities.name} entity provider for the type ${type}`);
      }
      this._deserialize.set(type, this._entities.getDeserializationHelper(type));
    }
  }
}
