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
import { OrganizationMembershipRole } from '../../business/organization';
import { stringOrNumberAsString } from '../../utils';
import {
  PostgresJsonEntityQuery,
  PostgresGetAllEntities,
  PostgresSettings,
  PostgresConfiguration,
} from '../../lib/entityMetadataProvider/postgres';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = new EntityMetadataType('OrganizationMemberCache');

interface IOrganizationMemberCacheProperties {
  // entity ID: orgid:userid
  uniqueId: any;
  cacheUpdated: any;
  organizationId: any;
  userId: any;
  role: any;
}

const Field: IOrganizationMemberCacheProperties = {
  uniqueId: 'uniqueId',
  cacheUpdated: 'cacheUpdated',
  organizationId: 'organizationId',
  userId: 'userId',
  role: 'role',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class OrganizationMemberCacheEntity
  implements IOrganizationMemberCacheProperties {
  uniqueId: string;
  cacheUpdated: Date;
  organizationId: string;
  userId: string;
  role: OrganizationMembershipRole;

  public static GenerateIdentifier(organizationId: string, userId: string) {
    if (!organizationId) {
      throw new Error('organizationId required');
    }
    if (!userId) {
      throw new Error('userId required');
    }
    return `${organizationId}:${userId}`;
  }

  constructor() {
    this.cacheUpdated = new Date();
  }
}

export class OrganizationBasicsFixedQuery implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.OrganizationCacheGetAllBasics;
}

export class OrganizationMemberCacheDeleteByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.OrganizationMemberCacheDeleteByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class OrganizationOwnersQuery implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.OrganizationOwnersCache;
}

export class OrganizationMemberCacheFixedQueryAll
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.OrganizationMemberCacheGetAll;
}

export class OrganizationMemberCacheFixedQueryByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.OrganizationMemberCacheByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class OrganizationMemberCacheFixedQueryByUserId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.OrganizationMemberCacheByUserId;
  constructor(public userId: string) {
    if (typeof this.userId !== 'string') {
      throw new Error(`${userId} must be a string`);
    }
  }
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new OrganizationMemberCacheEntity();
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
    [Field.cacheUpdated, 'cached'],
    [Field.organizationId, 'orgid'],
    [Field.uniqueId, 'unique'],
    [Field.userId, 'userid'],
    [Field.role, 'role'],
  ])
);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  MemorySettings.MemoryMapping,
  fieldNames,
  []
);

PostgresConfiguration.SetDefaultTableName(type, 'organizationmembercache');
EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'organizationmembercache'
);
PostgresConfiguration.MapFieldsToColumnNames(
  type,
  new Map<string, string>([
    [Field.cacheUpdated, (Field.cacheUpdated as string).toLowerCase()],
    [Field.organizationId, (Field.organizationId as string).toLowerCase()], // net new
    [Field.uniqueId, (Field.uniqueId as string).toLowerCase()],
    [Field.userId, (Field.userId as string).toLowerCase()],
    [Field.role, (Field.role as string).toLowerCase()],
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
      case FixedQueryType.OrganizationMemberCacheGetAll:
        return PostgresGetAllEntities(
          tableName,
          entityTypeColumn,
          entityTypeValue
        );
      case FixedQueryType.OrganizationMemberCacheDeleteByOrganizationId: {
        const {
          organizationId,
        } = query as OrganizationMemberCacheDeleteByOrganizationId;
        return {
          sql: `DELETE FROM ${tableName} WHERE ${metadataColumnName}->>'organizationid' = $1`,
          values: [organizationId],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.OrganizationCacheGetAllBasics: {
        return {
          sql: `
          SELECT DISTINCT(${metadataColumnName}->>'organizationid') as organizationid
          FROM ${tableName}`,
          values: [],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.OrganizationOwnersCache: {
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            role: 'admin',
          }
        );
      }
      case FixedQueryType.OrganizationOwnersCache: {
        const {
          organizationId,
        } = query as OrganizationMemberCacheFixedQueryByOrganizationId;
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
      case FixedQueryType.OrganizationMemberCacheByOrganizationId: {
        const {
          organizationId,
        } = query as OrganizationMemberCacheFixedQueryByOrganizationId;
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
      case FixedQueryType.OrganizationMemberCacheByUserId: {
        const { userId } = query as OrganizationMemberCacheFixedQueryByUserId;
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
      case FixedQueryType.OrganizationMemberCacheGetAll:
        return allInTypeBin;

      case FixedQueryType.OrganizationMemberCacheByOrganizationId:
        const {
          organizationId,
        } = query as OrganizationMemberCacheFixedQueryByOrganizationId;
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
