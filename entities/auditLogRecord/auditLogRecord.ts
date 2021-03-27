//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { EntityField } from '../../lib/entityMetadataProvider/entityMetadataProvider';
import {
  IEntityMetadata,
  EntityMetadataType,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  IEntityMetadataFixedQuery,
  FixedQueryType,
} from '../../lib/entityMetadataProvider/query';
import {
  EntityMetadataMappings,
  MetadataMappingDefinition,
} from '../../lib/entityMetadataProvider/declarations';
import { Type } from './type';
import {
  PostgresGetAllEntities,
  PostgresJsonEntityQuery,
  PostgresGetByID,
  PostgresSettings,
  PostgresConfiguration,
} from '../../lib/entityMetadataProvider/postgres';
import { IDictionary } from '../../transitional';
import { v4 } from 'uuid';
import { AuditLogSource } from '.';
import { stringOrNumberAsString } from '../../utils';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = Type;

interface IAuditLogRecordProperties {
  // THIS IS THE PRIMARY ID: recordId: any;
  recordSource: any;
  action: any;
  actorUsername: any;
  actorId: any;
  actorCorporateId: any;
  actorCorporateUsername: any;
  userUsername: any;
  userId: any;
  userCorporateId: any;
  userCorporateUsername: any;
  organizationName: any;
  organizationId: any;
  repositoryName: any;
  repositoryId: any;
  created: any;
  inserted: any;
  incomingUsername: any;
  incomingId: any;
  teamName: any;
  teamId: any;
  additionalData: any;
}

const recordId = 'recordId';

const Field: IAuditLogRecordProperties = {
  // recordId: 'recordId',
  recordSource: 'recordSource',
  action: 'action',
  actorUsername: 'actorUsername',
  actorId: 'actorId',
  actorCorporateId: 'actorCorporateId',
  actorCorporateUsername: 'actorCorporateUsername',
  created: 'created',
  inserted: 'inserted',
  organizationName: 'organizationName',
  organizationId: 'organizationId',
  repositoryName: 'repositoryName',
  repositoryId: 'repositoryId',
  userUsername: 'userUsername',
  userId: 'userId',
  userCorporateId: 'userCorporateId',
  userCorporateUsername: 'userCorporateUsername',
  incomingUsername: 'incomingUsername',
  incomingId: 'incomingId',
  teamName: 'teamName',
  teamId: 'teamId',
  additionalData: 'additionalData',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class AuditLogRecord implements IAuditLogRecordProperties {
  recordId: string;
  recordSource: AuditLogSource;

  action: string;

  additionalData: IDictionary<any>;

  repositoryId: string;
  repositoryName: string;

  organizationId: string;
  organizationName: string;

  created: Date;
  inserted: Date;

  actorUsername: string;
  actorId: string;
  actorCorporateId: string;
  actorCorporateUsername: string;

  userUsername: string;
  userId: string;
  userCorporateId: string;
  userCorporateUsername: string;

  incomingUsername: string;
  incomingId: string;
  teamName: string;
  teamId: string;

  constructor() {
    this.recordId = v4();
  }
}

export class AuditLogRecordQueryUndoCandidatesByThirdPartyId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.AuditLogUndoCandidateRecordsByThirdPartyId;
  constructor(public thirdPartyId: string) {}
}

export class AuditLogRecordQueryRecordsByRepositoryId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.AuditLogRecordsByRepositoryId;
  constructor(public repositoryId: string) {}
}

export class AuditLogRecordQueryRecordsByActorThirdPartyId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.AuditLogRecordsByActorThirdPartyId;
  constructor(public thirdPartyId: string) {}
}

export class AuditLogRecordQueryRecordsByUserThirdPartyId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.AuditLogRecordsByUserThirdPartyId;
  constructor(public thirdPartyId: string) {}
}

export class AuditLogRecordQueryRecordsByTeamId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.AuditLogRecordsByTeamId;
  constructor(public teamId: string) {}
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new AuditLogRecord();
  }
);
EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityIdColumnName,
  recordId
);

EntityMetadataMappings.Register(
  type,
  MemorySettings.MemoryMapping,
  new Map<string, string>([
    [Field.recordSource, Field.recordSource.toLowerCase()],
    [Field.action, Field.action.toLowerCase()],
    [Field.actorId, Field.actorId.toLowerCase()],
    [Field.actorUsername, Field.actorUsername.toLowerCase()],
    [Field.additionalData, Field.additionalData.toLowerCase()],
    [Field.actorCorporateUsername, Field.actorCorporateUsername.toLowerCase()],
    [Field.actorCorporateId, Field.actorCorporateId.toLowerCase()],
    [Field.teamId, Field.teamId.toLowerCase()],
    [Field.teamName, Field.teamName.toLowerCase()],
    [Field.incomingId, Field.incomingId.toLowerCase()],
    [Field.incomingUsername, Field.incomingUsername.toLowerCase()],
    [Field.inserted, Field.inserted.toLowerCase()],
    [Field.created, Field.created.toLowerCase()],
    [Field.repositoryId, Field.repositoryId.toLowerCase()],
    [Field.repositoryName, Field.repositoryName.toLowerCase()],
    [Field.organizationId, Field.organizationId.toLowerCase()],
    [Field.organizationName, Field.organizationName.toLowerCase()],
    [Field.userId, Field.userId.toLowerCase()],
    [Field.userUsername, Field.userUsername.toLowerCase()],
    [Field.userCorporateId, Field.userCorporateId.toLowerCase()],
    [Field.userCorporateUsername, Field.userCorporateUsername.toLowerCase()],
  ])
);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  MemorySettings.MemoryMapping,
  fieldNames,
  [recordId]
);

PostgresConfiguration.SetDefaultTableName(type, 'auditlog');
EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'auditlogrecord'
);
PostgresConfiguration.MapFieldsToColumnNames(
  type,
  new Map<string, string>([
    [Field.recordSource, Field.recordSource.toLowerCase()],
    [Field.action, Field.action.toLowerCase()],
    [Field.actorId, Field.actorId.toLowerCase()],
    [Field.actorUsername, Field.actorUsername.toLowerCase()],
    [Field.actorCorporateUsername, Field.actorCorporateUsername.toLowerCase()],
    [Field.actorCorporateId, Field.actorCorporateId.toLowerCase()],
    [Field.additionalData, Field.additionalData.toLowerCase()],
    [Field.teamId, Field.teamId.toLowerCase()],
    [Field.teamName, Field.teamName.toLowerCase()],
    [Field.incomingId, Field.incomingId.toLowerCase()],
    [Field.incomingUsername, Field.incomingUsername.toLowerCase()],
    [Field.inserted, Field.inserted.toLowerCase()],
    [Field.created, Field.created.toLowerCase()],
    [Field.repositoryId, Field.repositoryId.toLowerCase()],
    [Field.repositoryName, Field.repositoryName.toLowerCase()],
    [Field.organizationId, Field.organizationId.toLowerCase()],
    [Field.organizationName, Field.organizationName.toLowerCase()],
    [Field.userId, Field.userId.toLowerCase()],
    [Field.userUsername, Field.userUsername.toLowerCase()],
    [Field.userCorporateId, Field.userCorporateId.toLowerCase()],
    [Field.userCorporateUsername, Field.userCorporateUsername.toLowerCase()],
  ])
);
PostgresConfiguration.ValidateMappings(type, fieldNames, [recordId]);

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
      case FixedQueryType.AuditLogRecordsByActorThirdPartyId: {
        const {
          thirdPartyId,
        } = query as AuditLogRecordQueryRecordsByActorThirdPartyId;
        if (!thirdPartyId) {
          throw new Error('thirdPartyId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            actorid: stringOrNumberAsString(thirdPartyId),
          },
          Field.created.toLowerCase(),
          true
        );
      }
      case FixedQueryType.AuditLogRecordsByRepositoryId: {
        const {
          repositoryId,
        } = query as AuditLogRecordQueryRecordsByRepositoryId;
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
          },
          Field.created.toLowerCase(),
          true
        );
      }
      case FixedQueryType.AuditLogRecordsByTeamId: {
        const { teamId } = query as AuditLogRecordQueryRecordsByTeamId;
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
          },
          Field.created.toLowerCase(),
          true
        );
      }
      case FixedQueryType.AuditLogUndoCandidateRecordsByThirdPartyId: {
        const {
          thirdPartyId,
        } = query as AuditLogRecordQueryUndoCandidatesByThirdPartyId;
        if (!thirdPartyId) {
          throw new Error('thirdPartyId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            actorid: stringOrNumberAsString(thirdPartyId),
            additionaldata: {
              undoCandidate: true,
            },
          },
          Field.created.toLowerCase(),
          true
        );
      }
      case FixedQueryType.AuditLogRecordsByUserThirdPartyId: {
        const {
          thirdPartyId,
        } = query as AuditLogRecordQueryRecordsByUserThirdPartyId;
        if (!thirdPartyId) {
          throw new Error('thirdPartyId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            userid: stringOrNumberAsString(thirdPartyId),
          },
          Field.created.toLowerCase(),
          true
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
