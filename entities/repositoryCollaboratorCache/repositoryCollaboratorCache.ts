//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { EntityField } from '../../lib/entityMetadataProvider/entityMetadataProvider';
import {
  EntityMetadataType,
  IEntityMetadata,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  IEntityMetadataFixedQuery,
  FixedQueryType,
} from '../../lib/entityMetadataProvider/query';
import {
  EntityMetadataMappings,
  MetadataMappingDefinition,
} from '../../lib/entityMetadataProvider/declarations';
import { GitHubRepositoryPermission } from '../repositoryMetadata/repositoryMetadata';
import { GitHubCollaboratorType } from '../../business/repository';
import {
  PostgresGetAllEntities,
  PostgresJsonEntityQuery,
  PostgresSettings,
  PostgresConfiguration,
} from '../../lib/entityMetadataProvider/postgres';
import { stringOrNumberAsString } from '../../utils';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = new EntityMetadataType('RepositoryCollaboratorCache');

interface IRepositoryCollaboratorCacheProperties {
  // entity ID: orgid:repoid:userid
  uniqueId: any;

  cacheUpdated: any;
  organizationId: any;
  repositoryId: any;
  userId: any;
  permission: any;

  repositoryPrivate: any;
  repositoryName: any;

  login: any;
  avatar: any;
  collaboratorType: any;
}

const Field: IRepositoryCollaboratorCacheProperties = {
  uniqueId: 'uniqueId',
  cacheUpdated: 'cacheUpdated',

  organizationId: 'organizationId',
  repositoryId: 'repositoryId',
  repositoryName: 'repositoryName',
  repositoryPrivate: 'repositoryPrivate',
  userId: 'userId',
  permission: 'permission',
  login: 'login',
  avatar: 'avatar',
  collaboratorType: 'collaboratorType',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class RepositoryCollaboratorCacheEntity
  implements IRepositoryCollaboratorCacheProperties {
  uniqueId: string;
  cacheUpdated: Date;

  organizationId: string;
  repositoryId: string;
  userId: string;

  repositoryName: string;
  repositoryPrivate: boolean;

  permission: GitHubRepositoryPermission;
  login: string;
  avatar: string;
  collaboratorType: GitHubCollaboratorType;

  public static GenerateIdentifier(
    organizationId: string,
    repositoryId: string,
    userId: string
  ) {
    if (!organizationId) {
      throw new Error('organizationId required');
    }
    if (!repositoryId) {
      throw new Error('repositoryId required');
    }
    if (!userId) {
      throw new Error('userId required');
    }
    return `${organizationId}:${repositoryId}:${userId}`;
  }

  constructor() {
    this.cacheUpdated = new Date();
  }
}

export class RepositoryCollaboratorCacheFixedQueryAll
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCollaboratorCacheGetAll;
}

export class RepositoryCollaboratorCacheGetOrganizationIdsQuery
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCollaboratorCacheGetOrganizationIds;
}

export class RepositoryCollaboratorCacheDeleteByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCollaboratorCacheDeleteByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`organizationId ${organizationId} must be a string`);
    }
  }
}

export class RepositoryCollaboratorCacheDeleteByRepositoryId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCollaboratorCacheDeleteByRepositoryId;
  constructor(public repositoryId: string) {
    if (typeof this.repositoryId !== 'string') {
      throw new Error(`repositoryId ${repositoryId} must be a string`);
    }
  }
}

export class RepositoryCollaboratorCacheFixedQueryByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCollaboratorCacheByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class RepositoryCollaboratorCacheFixedQueryByUserId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCollaboratorCacheByUserId;
  constructor(public userId: string) {
    if (typeof this.userId !== 'string') {
      throw new Error(`${userId} must be a string`);
    }
  }
}

export class RepositoryCollaboratorCacheFixedQueryByRepositoryId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCollaboratorCacheByRepositoryId;
  constructor(public repositoryId: string) {
    if (typeof this.repositoryId !== 'string') {
      throw new Error(`${repositoryId} must be a string`);
    }
  }
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new RepositoryCollaboratorCacheEntity();
  }
);
EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityIdColumnName,
  Field.uniqueId
);

EntityMetadataMappings.Register(
  type,
  MemorySettings.MemoryMapping,
  new Map<string, string>([
    [Field.avatar, 'avatar'],
    [Field.cacheUpdated, 'cached'],
    [Field.login, 'login'],
    [Field.organizationId, 'orgid'],
    [Field.permission, 'permission'],
    [Field.repositoryId, 'repoid'],
    [Field.repositoryName, 'reponame'],
    [Field.repositoryPrivate, 'repoprivate'],
    [Field.uniqueId, 'unique'],
    [Field.userId, 'userid'],
    [Field.collaboratorType, 'collaboratorType'],
  ])
);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  MemorySettings.MemoryMapping,
  fieldNames,
  []
);

PostgresConfiguration.SetDefaultTableName(type, 'repositorycollaboratorcache');
EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'repositorycollaboratorcache'
);
PostgresConfiguration.MapFieldsToColumnNames(
  type,
  new Map<string, string>([
    [Field.avatar, (Field.avatar as string).toLowerCase()],
    [Field.cacheUpdated, (Field.cacheUpdated as string).toLowerCase()],
    [Field.login, (Field.login as string).toLowerCase()],
    [Field.organizationId, (Field.organizationId as string).toLowerCase()], // net new
    [Field.repositoryName, (Field.repositoryName as string).toLowerCase()], // net new
    [
      Field.repositoryPrivate,
      (Field.repositoryPrivate as string).toLowerCase(),
    ], // net new
    [Field.permission, (Field.permission as string).toLowerCase()],
    [Field.repositoryId, (Field.repositoryId as string).toLowerCase()],
    [Field.uniqueId, (Field.uniqueId as string).toLowerCase()],
    [Field.userId, (Field.userId as string).toLowerCase()],
    [Field.collaboratorType, (Field.collaboratorType as string).toLowerCase()],
  ])
);
PostgresConfiguration.ValidateMappings(type, fieldNames, []);

EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresQueries,
  (
    query: IEntityMetadataFixedQuery,
    mapMetadataPropertiesToFields: string[],
    metadataColumnName: string,
    tableName: string,
    getEntityTypeColumnValue
  ) => {
    const entityTypeColumn = mapMetadataPropertiesToFields[EntityField.Type];
    const entityTypeValue = getEntityTypeColumnValue(type);
    switch (query.fixedQueryType) {
      case FixedQueryType.RepositoryCollaboratorCacheGetAll:
        return PostgresGetAllEntities(
          tableName,
          entityTypeColumn,
          entityTypeValue
        );
      case FixedQueryType.RepositoryCollaboratorCacheByOrganizationId: {
        const {
          organizationId,
        } = query as RepositoryCollaboratorCacheFixedQueryByOrganizationId;
        if (!organizationId) {
          throw new Error('organizationId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            organizationid: stringOrNumberAsString(organizationId),
          }
        );
      }
      case FixedQueryType.RepositoryCollaboratorCacheByUserId: {
        const {
          userId,
        } = query as RepositoryCollaboratorCacheFixedQueryByUserId;
        if (!userId) {
          throw new Error('userId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            userid: stringOrNumberAsString(userId),
          }
        );
      }
      case FixedQueryType.RepositoryCollaboratorCacheDeleteByOrganizationId: {
        const {
          organizationId,
        } = query as RepositoryCollaboratorCacheDeleteByOrganizationId;
        return {
          sql: `DELETE FROM ${tableName} WHERE ${metadataColumnName}->>'organizationid' = $1`,
          values: [organizationId],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.RepositoryCollaboratorCacheDeleteByRepositoryId: {
        const {
          repositoryId,
        } = query as RepositoryCollaboratorCacheDeleteByRepositoryId;
        return {
          sql: `DELETE FROM ${tableName} WHERE ${metadataColumnName}->>'repositoryid' = $1`,
          values: [repositoryId],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.RepositoryCollaboratorCacheGetOrganizationIds: {
        return {
          sql: `
          SELECT DISTINCT(${metadataColumnName}->>'organizationid') as organizationid
          FROM ${tableName}`,
          values: [],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.RepositoryCollaboratorCacheByRepositoryId: {
        const {
          repositoryId,
        } = query as RepositoryCollaboratorCacheFixedQueryByRepositoryId;
        if (!repositoryId) {
          throw new Error('repositoryId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            repositoryid: stringOrNumberAsString(repositoryId),
          }
        );
      }
      default:
        throw new Error(
          `The fixed query type "${query.fixedQueryType}" is not implemented by this provider for the type ${type}, or is of an unknown type`
        );
    }
  }
);

EntityMetadataMappings.Register(
  type,
  MemorySettings.MemoryQueries,
  (query: IEntityMetadataFixedQuery, allInTypeBin: IEntityMetadata[]) => {
    switch (query.fixedQueryType) {
      case FixedQueryType.RepositoryCollaboratorCacheGetAll:
        return allInTypeBin;

      case FixedQueryType.RepositoryCollaboratorCacheByOrganizationId:
        const {
          organizationId,
        } = query as RepositoryCollaboratorCacheFixedQueryByOrganizationId;
        if (!organizationId) {
          throw new Error('organizationId required');
        }
        throw new Error('Not implemented yet');
      default:
        throw new Error(
          `The fixed query type "${query.fixedQueryType}" is not implemented by this provider for the type ${type}, or is of an unknown type`
        );
    }
  }
);

// Runtime validation of FieldNames
for (let i = 0; i < fieldNames.length; i++) {
  const fn = fieldNames[i];
  if (Field[fn] !== fn) {
    throw new Error(`Field name ${fn} and value do not match in ${__filename}`);
  }
}

export const EntityImplementation = {
  EnsureDefinitions: () => {},
  Type: type,
};
