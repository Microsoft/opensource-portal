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
import {
  PostgresGetAllEntities,
  PostgresJsonEntityQuery,
  PostgresSettings,
  PostgresConfiguration,
} from '../../lib/entityMetadataProvider/postgres';
import { stringOrNumberAsString } from '../../utils';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';
import { Operations } from '../../business/operations';

const type = new EntityMetadataType('RepositoryCache');

interface IRepositoryCacheProperties {
  organizationId: any;
  repositoryName: any;
  repositoryDetails: any;
  cacheUpdated: any;
}

const repositoryId = 'repositoryId';

const Field: IRepositoryCacheProperties = {
  organizationId: 'organizationId',
  repositoryName: 'repositoryName',
  repositoryDetails: 'repositoryDetails',
  cacheUpdated: 'cacheUpdated',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class RepositoryCacheEntity implements IRepositoryCacheProperties {
  repositoryId: string;
  repositoryName: string;
  organizationId: string;
  repositoryDetails: any;
  cacheUpdated: Date;

  constructor() {
    this.cacheUpdated = new Date();
  }

  hydrateToInstance(operations: Operations) {
    try {
      const organization = operations.getOrganizationById(
        Number(this.organizationId)
      );
      const clone = { ...this.repositoryDetails };
      clone.id = Number(this.repositoryId); // GitHub entities are numbers
      return organization.repository(this.repositoryName, clone);
    } catch (noConfiguredOrganization) {
      throw noConfiguredOrganization;
    }
  }
}

export class RepositoryCacheFixedQueryAll implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCacheGetAll;
}

export class RepositoryCacheGetOrganizationIdsQuery
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCacheGetOrganizationIds;
}

export class RepositoryCacheDeleteByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCacheDeleteByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class RepositoryCacheFixedQueryByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.RepositoryCacheGetByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new RepositoryCacheEntity();
  }
);
EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityIdColumnName,
  repositoryId
);

EntityMetadataMappings.Register(
  type,
  MemorySettings.MemoryMapping,
  new Map<string, string>([
    [Field.organizationId, 'orgid'],
    [Field.repositoryName, 'repoName'],
    [Field.repositoryDetails, 'repoDetails'],
    [Field.cacheUpdated, 'cached'],
  ])
);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  MemorySettings.MemoryMapping,
  fieldNames,
  [repositoryId]
);

PostgresConfiguration.SetDefaultTableName(type, 'repositorycache');
EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'repositorycache'
);
PostgresConfiguration.MapFieldsToColumnNames(
  type,
  new Map<string, string>([
    [Field.organizationId, (Field.organizationId as string).toLowerCase()], // net new
    [Field.repositoryName, (Field.repositoryName as string).toLowerCase()],
    [
      Field.repositoryDetails,
      (Field.repositoryDetails as string).toLowerCase(),
    ],
    [Field.cacheUpdated, (Field.cacheUpdated as string).toLowerCase()],
  ])
);
PostgresConfiguration.ValidateMappings(type, fieldNames, [repositoryId]);

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
      case FixedQueryType.RepositoryCacheGetAll:
        return PostgresGetAllEntities(
          tableName,
          entityTypeColumn,
          entityTypeValue
        );
      case FixedQueryType.RepositoryCacheGetByOrganizationId: {
        const {
          organizationId,
        } = query as RepositoryCacheFixedQueryByOrganizationId;
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
      case FixedQueryType.RepositoryCacheDeleteByOrganizationId: {
        const {
          organizationId,
        } = query as RepositoryCacheDeleteByOrganizationId;
        return {
          sql: `DELETE FROM ${tableName} WHERE ${metadataColumnName}->>'organizationid' = $1`,
          values: [organizationId],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.RepositoryCacheGetOrganizationIds: {
        return {
          sql: `
          SELECT DISTINCT(${metadataColumnName}->>'organizationid') as organizationid
          FROM ${tableName}`,
          values: [],
          skipEntityMapping: true,
        };
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
      case FixedQueryType.RepositoryCacheGetAll:
        return allInTypeBin;
      case FixedQueryType.RepositoryCacheGetByOrganizationId: {
        const {
          organizationId,
        } = query as RepositoryCacheFixedQueryByOrganizationId;
        if (!organizationId) {
          throw new Error('organizationId required');
        }
        throw new Error('Not implemented yet');
      }
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
