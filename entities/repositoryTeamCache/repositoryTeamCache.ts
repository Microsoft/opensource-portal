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
import {
  PostgresGetAllEntities,
  PostgresJsonEntityQuery,
  PostgresJsonEntityQueryMultiple,
  PostgresSettings,
  PostgresConfiguration,
} from '../../lib/entityMetadataProvider/postgres';
import { stringOrNumberAsString } from '../../utils';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = new EntityMetadataType('RepositoryTeamCache');

interface IRepositoryTeamCacheProperties {
  // entity ID: orgid:repoid:teamid
  uniqueId: any;
  cacheUpdated: any;

  organizationId: any;
  repositoryId: any;
  repositoryName: any;
  teamId: any;

  repositoryPrivate: any;

  permission: any;
}

const Field: IRepositoryTeamCacheProperties = {
  uniqueId: 'uniqueId',
  cacheUpdated: 'cacheUpdated',

  organizationId: 'organizationId',
  repositoryId: 'repositoryId',
  repositoryName: 'repositoryName',
  teamId: 'teamId',
  permission: 'permission',

  repositoryPrivate: 'repositoryPrivate',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class RepositoryTeamCacheEntity
  implements IRepositoryTeamCacheProperties {
  uniqueId: string;
  cacheUpdated: Date;

  organizationId: string;
  repositoryId: string;
  repositoryName: string;
  teamId: string;

  repositoryPrivate: boolean;

  permission: GitHubRepositoryPermission;

  public static GenerateIdentifier(
    organizationId: string,
    repositoryId: string,
    teamId: string
  ) {
    if (!organizationId) {
      throw new Error('organizationId required');
    }
    if (!repositoryId) {
      throw new Error('repositoryId required');
    }
    if (!teamId) {
      throw new Error('teamId required');
    }
    return `${organizationId}:${repositoryId}:${teamId}`;
  }

  constructor() {
    this.cacheUpdated = new Date();
  }
}

export class RepositoryTeamCacheGetOrganizationIdsQuery
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheGetOrganizationIds;
}

export class RepositoryTeamCacheDeleteByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheDeleteByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class RepositoryTeamCacheDeleteByRepositoryId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheDeleteByRepositoryId;
  constructor(public repositoryId: string) {
    if (typeof this.repositoryId !== 'string') {
      throw new Error(`repositoryId ${repositoryId} must be a string`);
    }
  }
}

export class RepositoryTeamCacheFixedQueryAll
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheGetAll;
}

export class RepositoryTeamCacheFixedQueryByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class RepositoryTeamCacheFixedQueryByTeamId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheByTeamId;
  constructor(public teamId: string) {
    if (typeof this.teamId !== 'string') {
      throw new Error(`${teamId} must be a string`);
    }
  }
}

export class RepositoryTeamCacheFixedQueryByTeamIds
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheByTeamIds;
  constructor(public teamIds: string[]) {
    if (!Array.isArray(this.teamIds)) {
      throw new Error(`teamIds must be an array`);
    }
    // should also make sure the array is of strings, not numbers
  }
}

export class RepositoryTeamCacheFixedQueryByRepositoryId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryTeamCacheByRepositoryId;
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
    return new RepositoryTeamCacheEntity();
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
    [Field.permission, 'permission'],
    [Field.repositoryId, 'repoid'],
    [Field.repositoryPrivate, 'repoprivate'],
    [Field.uniqueId, 'unique'],
    [Field.teamId, 'teamId'],
    [Field.repositoryName, 'repositoryName'],
  ])
);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  MemorySettings.MemoryMapping,
  fieldNames,
  []
);

PostgresConfiguration.SetDefaultTableName(type, 'repositoryteamcache');
EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'repositoryteamcache'
);
PostgresConfiguration.MapFieldsToColumnNames(
  type,
  new Map<string, string>([
    [Field.cacheUpdated, (Field.cacheUpdated as string).toLowerCase()],
    [Field.organizationId, (Field.organizationId as string).toLowerCase()], // net new
    [Field.permission, (Field.permission as string).toLowerCase()],
    [Field.repositoryId, (Field.repositoryId as string).toLowerCase()],
    [Field.repositoryName, (Field.repositoryName as string).toLowerCase()],
    [Field.uniqueId, (Field.uniqueId as string).toLowerCase()],
    [Field.teamId, (Field.teamId as string).toLowerCase()],
    [
      Field.repositoryPrivate,
      (Field.repositoryPrivate as string).toLowerCase(),
    ],
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
      case FixedQueryType.RepositoryTeamCacheGetAll:
        return PostgresGetAllEntities(
          tableName,
          entityTypeColumn,
          entityTypeValue
        );
      case FixedQueryType.RepositoryTeamCacheByOrganizationId: {
        const {
          organizationId,
        } = query as RepositoryTeamCacheFixedQueryByOrganizationId;
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
      case FixedQueryType.RepositoryTeamCacheByRepositoryId: {
        const {
          repositoryId,
        } = query as RepositoryTeamCacheFixedQueryByRepositoryId;
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
      case FixedQueryType.RepositoryTeamCacheByTeamId: {
        const { teamId } = query as RepositoryTeamCacheFixedQueryByTeamId;
        if (!teamId) {
          throw new Error('teamId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            teamid: stringOrNumberAsString(teamId),
          }
        );
      }
      case FixedQueryType.RepositoryTeamCacheDeleteByRepositoryId: {
        const {
          repositoryId,
        } = query as RepositoryTeamCacheDeleteByRepositoryId;
        return {
          sql: `DELETE FROM ${tableName} WHERE ${metadataColumnName}->>'repositoryid' = $1`,
          values: [repositoryId],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.RepositoryTeamCacheDeleteByOrganizationId: {
        const {
          organizationId,
        } = query as RepositoryTeamCacheDeleteByOrganizationId;
        return {
          sql: `DELETE FROM ${tableName} WHERE ${metadataColumnName}->>'organizationid' = $1`,
          values: [organizationId],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.RepositoryTeamCacheGetOrganizationIds: {
        return {
          sql: `
          SELECT DISTINCT(${metadataColumnName}->>'organizationid') as organizationid
          FROM ${tableName}`,
          values: [],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.RepositoryTeamCacheByTeamIds: {
        const { teamIds } = query as RepositoryTeamCacheFixedQueryByTeamIds;
        if (!teamIds) {
          throw new Error('teamIds required');
        }
        if (teamIds.length === 0) {
          throw new Error('teamIds must have at least 1 team ID');
        }
        return PostgresJsonEntityQueryMultiple(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          teamIds.map((teamId) => {
            return {
              teamid: stringOrNumberAsString(teamId),
            };
          })
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
      case FixedQueryType.RepositoryTeamCacheGetAll:
        return allInTypeBin;

      case FixedQueryType.RepositoryTeamCacheByOrganizationId:
        const {
          organizationId,
        } = query as RepositoryTeamCacheFixedQueryByOrganizationId;
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
