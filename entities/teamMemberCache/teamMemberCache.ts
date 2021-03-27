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
  TeamMemberCacheFixedQueryByOrganizationId,
  TeamMemberCacheFixedQueryByUserId,
  TeamMemberCacheFixedQueryByTeamId,
  TeamMemberCacheFixedQueryByOrganizationIdAndUserId,
  TeamMemberCacheDeleteByOrganizationId,
} from '.';
import { GitHubTeamRole } from '../../business/team';
import {
  PostgresGetAllEntities,
  PostgresJsonEntityQuery,
  PostgresSettings,
  PostgresConfiguration,
} from '../../lib/entityMetadataProvider/postgres';
import { stringOrNumberAsString } from '../../utils';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = new EntityMetadataType('TeamMemberCache');

interface ITeamMemberCacheProperties {
  uniqueId: any; // orgid:teamid:userid
  cacheUpdated: any;
  organizationId: any;
  teamId: any;
  userId: any;
  teamRole: any;
  login: any;
  avatar: any;
}

const teamId = 'teamId';

const Field: ITeamMemberCacheProperties = {
  uniqueId: 'uniqueId',
  cacheUpdated: 'cacheUpdated',
  organizationId: 'organizationId',
  teamId: 'teamId',
  userId: 'userId',
  teamRole: 'teamRole',
  login: 'login',
  avatar: 'avatar',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class TeamMemberCacheEntity implements ITeamMemberCacheProperties {
  uniqueId: string;
  cacheUpdated: Date;

  organizationId: string;
  teamId: string;
  userId: string;
  teamRole: GitHubTeamRole;

  login: string;
  avatar: string;

  public static GenerateIdentifier(
    organizationId: string,
    teamId: string,
    userId: string
  ) {
    if (!organizationId) {
      throw new Error('organizationId required');
    }
    if (!teamId) {
      throw new Error('teamId required');
    }
    if (!userId) {
      throw new Error('userId required');
    }
    return `${organizationId}:${teamId}:${userId}`;
  }

  constructor() {
    this.cacheUpdated = new Date();
  }
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new TeamMemberCacheEntity();
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
    [Field.organizationId, 'orgid'],
    [Field.teamId, 'teamName'],
    [Field.userId, 'teamSlug'],
    [Field.uniqueId, 'uniqueId'],
    [Field.cacheUpdated, 'cached'],
    [Field.teamRole, 'role'],
    [Field.login, 'login'],
    [Field.avatar, 'avatar'],
  ])
);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  MemorySettings.MemoryMapping,
  fieldNames,
  []
);
PostgresConfiguration.SetDefaultTableName(type, 'teammembercache');
EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'teammembercache'
);
PostgresConfiguration.MapFieldsToColumnNames(
  type,
  new Map<string, string>([
    [Field.organizationId, (Field.organizationId as string).toLowerCase()],
    [Field.uniqueId, (Field.uniqueId as string).toLowerCase()],
    [Field.teamId, (Field.teamId as string).toLowerCase()],
    [Field.userId, (Field.userId as string).toLowerCase()],
    [Field.cacheUpdated, (Field.cacheUpdated as string).toLowerCase()],
    [Field.teamRole, (Field.teamRole as string).toLowerCase()],
    [Field.login, (Field.login as string).toLowerCase()],
    [Field.avatar, (Field.avatar as string).toLowerCase()],
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
      case FixedQueryType.TeamMemberCacheGetAll:
        return PostgresGetAllEntities(
          tableName,
          entityTypeColumn,
          entityTypeValue
        );
      case FixedQueryType.TeamMemberCacheGetByOrganizationId: {
        const {
          organizationId,
        } = query as TeamMemberCacheFixedQueryByOrganizationId;
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
      case FixedQueryType.TeamMemberCacheGetByUserId: {
        const { userId } = query as TeamMemberCacheFixedQueryByUserId;
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
      case FixedQueryType.TeamMemberCacheGetByTeamId: {
        const { teamId } = query as TeamMemberCacheFixedQueryByTeamId;
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
      case FixedQueryType.TeamMemberCacheDeleteByOrganizationId: {
        const {
          organizationId,
        } = query as TeamMemberCacheDeleteByOrganizationId;
        return {
          sql: `DELETE FROM ${tableName} WHERE ${metadataColumnName}->>'organizationid' = $1`,
          values: [organizationId],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.TeamMemberCacheGetOrganizationIds: {
        return {
          sql: `
          SELECT DISTINCT(${metadataColumnName}->>'organizationid') as organizationid
          FROM ${tableName}`,
          values: [],
          skipEntityMapping: true,
        };
      }
      case FixedQueryType.TeamMemberCacheGetByOrganizationIdAndUserId: {
        const {
          organizationId,
          userId,
        } = query as TeamMemberCacheFixedQueryByOrganizationIdAndUserId;
        if (!organizationId) {
          throw new Error('organizationId required');
        }
        if (!userId) {
          throw new Error('userId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            organizationid: stringOrNumberAsString(organizationId),
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
      case FixedQueryType.TeamMemberCacheGetAll:
        return allInTypeBin;

      case FixedQueryType.TeamMemberCacheGetByOrganizationId:
        const {
          organizationId,
        } = query as TeamMemberCacheFixedQueryByOrganizationId;
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
  Type: type,
  EnsureDefinitions: () => {},
};
