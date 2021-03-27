//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { v4 as uuidV4 } from 'uuid';
import azure from 'azure-storage';

import { IObjectWithDefinedKeys } from '../../lib/entityMetadataProvider/entityMetadataProvider';
import {
  EntityMetadataType,
  IEntityMetadata,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  MetadataMappingDefinition,
  EntityMetadataMappings,
} from '../../lib/entityMetadataProvider/declarations';
import {
  FixedQueryType,
  IEntityMetadataFixedQuery,
} from '../../lib/entityMetadataProvider/query';
import {
  stringOrNumberAsString,
  stringOrNumberArrayAsStringArray,
} from '../../utils';
import {
  PostgresGetAllEntities,
  PostgresJsonEntityQuery,
  PostgresSettings,
  PostgresConfiguration,
} from '../../lib/entityMetadataProvider/postgres';
import { TableSettings } from '../../lib/entityMetadataProvider/table';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = new EntityMetadataType('TeamJoinRequest');

interface ITeamJoinApprovalEntityProperties {
  // approvalId: any;

  thirdPartyId?: any;
  thirdPartyUsername?: any;

  corporateDisplayName?: any;
  corporateId?: any;
  corporateUsername?: any;

  active?: any;
  created?: any;

  justification?: any;

  organizationName?: any;
  teamId?: any;
  teamName?: any;

  mailSentTo?: any;
  mailSentToApprovers?: any;

  decision?: any;
  decisionMessage?: any;
  decisionTime?: any;
  decisionThirdPartyUsername?: any;
  decisionThirdPartyId?: any;
  decisionCorporateUsername?: any;
  decisionCorporateId?: any;

  ticketType?: any; // REMOVE
}

const Field: ITeamJoinApprovalEntityProperties = {
  // approvalId: 'approvalId',
  thirdPartyId: 'thirdPartyId',
  thirdPartyUsername: 'thirdPartyUsername',
  corporateDisplayName: 'corporateDisplayName',
  corporateId: 'corporateId',
  corporateUsername: 'corporateUsername',
  active: 'active',
  created: 'created',
  justification: 'justification',
  organizationName: 'organizationName',
  teamId: 'teamId',
  teamName: 'teamName',
  mailSentTo: 'mailSentTo',
  mailSentToApprovers: 'mailSentToApprovers',
  decision: 'decision',
  decisionThirdPartyUsername: 'decisionThirdPartyUsername',
  decisionThirdPartyId: 'decisionThirdPartyId',
  decisionCorporateUsername: 'decisionCorporateUsername',
  decisionCorporateId: 'decisionCorporateId',
  decisionMessage: 'decisionMessage',
  decisionTime: 'decisionTime',
  ticketType: 'ticketType', // TODO: remove tt
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class TeamJoinApprovalEntity
  implements IObjectWithDefinedKeys, ITeamJoinApprovalEntityProperties {
  approvalId: string;

  thirdPartyId?: string;
  thirdPartyUsername?: string;

  corporateDisplayName?: string;
  corporateId?: string;
  corporateUsername?: string;

  active?: boolean;
  created?: Date;

  justification?: string;

  organizationName?: string;
  teamId?: string;
  teamName?: string;

  mailSentTo?: string;
  mailSentToApprovers?: string;

  decision?: string;
  decisionThirdPartyUsername?: string;
  decisionThirdPartyId?: string;
  decisionCorporateUsername?: string;
  decisionCorporateId?: string;
  decisionMessage?: string;
  decisionTime?: Date;

  ticketType?: string;

  constructor() {
    this.approvalId = uuidV4();
    this.ticketType = this.type();
  }

  getObjectFieldNames(): string[] {
    return fieldNames;
  }

  type() {
    return 'joinTeam'; // legacy, used in a few views
  }
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new TeamJoinApprovalEntity();
  }
);
EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityIdColumnName,
  'approvalId'
);

EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultTableName,
  'pending'
);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultFixedPartitionKey,
  'pk'
);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableMapping,
  new Map<string, string>([
    [Field.thirdPartyId, 'ghid'],
    [Field.thirdPartyUsername, 'ghu'],

    [Field.corporateDisplayName, 'name'],
    [Field.corporateId, 'aadid'],
    [Field.corporateUsername, 'email'],

    [Field.justification, 'justification'],

    [Field.active, 'active'],
    [Field.created, 'requested'],

    [Field.organizationName, 'org'],
    [Field.teamId, 'teamid'],
    [Field.teamName, 'teamname'],

    [Field.mailSentTo, 'mailSentTo'],
    [Field.mailSentToApprovers, 'mailSentToApprovers'],

    [Field.decision, 'decision'],
    [Field.decisionTime, 'decisionTime'],
    [Field.decisionMessage, 'decisionNote'],
    [Field.decisionThirdPartyUsername, 'decisionBy'],
    [Field.decisionThirdPartyId, 'decisionById'],
    [Field.decisionCorporateUsername, 'decisionEmail'],
    [Field.decisionCorporateId, 'decisionCorporateId'],

    [Field.ticketType, 'tickettype'],
  ])
);
EntityMetadataMappings.Register(type, TableSettings.TablePossibleDateColumns, [
  Field.created,
  Field.decisionTime,
]);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  TableSettings.TableMapping,
  fieldNames,
  []
);

EntityMetadataMappings.Register(
  type,
  MemorySettings.MemoryMapping,
  new Map<string, string>([
    [Field.thirdPartyId, '__ghi'],
    [Field.thirdPartyUsername, '__ghu'],

    [Field.corporateDisplayName, '__cdn'],
    [Field.corporateId, '__cid'],
    [Field.corporateUsername, '__cu'],

    [Field.justification, '__justification__'],

    [Field.active, '__active'],
    [Field.created, '__created'],

    [Field.organizationName, '__orgName'],
    [Field.teamId, '__teamid'],
    [Field.teamName, '__teamname'],

    [Field.mailSentTo, '/mailSentTo'],
    [Field.mailSentToApprovers, '/senttoapprovers'],

    [Field.decision, 'decision'],
    [Field.decisionTime, 'dt'],
    [Field.decisionMessage, 'd'],
    [Field.decisionThirdPartyUsername, 'decisionBy'],
    [Field.decisionThirdPartyId, 'decisionById'],
    [Field.decisionCorporateUsername, 'decisionCorporateUsername'],
    [Field.decisionCorporateId, 'decisionCorporateId'],

    [Field.ticketType, '/type'],
  ])
);
EntityMetadataMappings.RuntimeValidateMappings(
  type,
  MemorySettings.MemoryMapping,
  fieldNames,
  []
);

EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'teamjoin'
);
PostgresConfiguration.SetDefaultTableName(type, 'approvals');
PostgresConfiguration.MapFieldsToColumnNames(
  type,
  new Map<string, string>([
    [Field.thirdPartyId, (Field.thirdPartyId as string).toLowerCase()],
    [
      Field.thirdPartyUsername,
      (Field.thirdPartyUsername as string).toLowerCase(),
    ],

    [
      Field.corporateDisplayName,
      (Field.corporateDisplayName as string).toLowerCase(),
    ],
    [Field.corporateId, (Field.corporateId as string).toLowerCase()],
    [
      Field.corporateUsername,
      (Field.corporateUsername as string).toLowerCase(),
    ],

    [Field.justification, (Field.justification as string).toLowerCase()],

    [Field.active, (Field.active as string).toLowerCase()],
    [Field.created, (Field.created as string).toLowerCase()],

    [Field.organizationName, (Field.organizationName as string).toLowerCase()],
    [Field.teamId, (Field.teamId as string).toLowerCase()],
    [Field.teamName, (Field.teamName as string).toLowerCase()],

    [Field.mailSentTo, (Field.mailSentTo as string).toLowerCase()],
    [
      Field.mailSentToApprovers,
      (Field.mailSentToApprovers as string).toLowerCase(),
    ],

    [Field.decision, (Field.decision as string).toLowerCase()],
    [Field.decisionTime, (Field.decisionTime as string).toLowerCase()],
    [Field.decisionMessage, (Field.decisionMessage as string).toLowerCase()],
    [
      Field.decisionThirdPartyUsername,
      (Field.decisionThirdPartyUsername as string).toLowerCase(),
    ],
    [
      Field.decisionThirdPartyId,
      (Field.decisionThirdPartyId as string).toLowerCase(),
    ],
    [
      Field.decisionCorporateUsername,
      (Field.decisionCorporateUsername as string).toLowerCase(),
    ],
    [
      Field.decisionCorporateId,
      (Field.decisionCorporateId as string).toLowerCase(),
    ],

    [Field.ticketType, 'tickettype'], // TODO: remove ticket type from team join approvals
  ])
);
PostgresConfiguration.ValidateMappings(type, fieldNames, []);

export class TeamJoinRequestFixedQueryAll implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.AllTeamJoinApprovals;
}

export class TeamJoinRequestFixedQueryByTeams
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.ActiveTeamJoinApprovalsByTeams;
  public ids: string[];
  constructor(ids: string[]) {
    if (!ids || !Array.isArray(ids)) {
      throw new Error('ids must be an array of team IDs');
    }
    this.ids = ids.map((id) => {
      if (typeof id === 'number') {
        id = (id as number).toString();
      }
      if (typeof id !== 'string') {
        throw new Error(
          `TeamJoinRequestFixedQueryByTeams: team ID must be a string: ${id}`
        );
      }
      return id;
    });
  }
}

export class TeamJoinRequestFixedQueryByTeam
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.ActiveTeamJoinApprovalsByTeam;
  constructor(public id: string) {
    if (typeof id === 'number') {
      id = (id as number).toString();
    }
    if (typeof id !== 'string') {
      throw new Error(
        `TeamJoinRequestFixedQueryByTeam: team ID must be a string: ${id}`
      );
    }
  }
}

export class TeamJoinRequestFixedQueryByThirdPartyUserId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.ActiveTeamJoinApprovalsByThirdPartyId;
  constructor(public thirdPartyId: string) {
    if (typeof thirdPartyId === 'number') {
      thirdPartyId = (thirdPartyId as number).toString();
    }
    if (typeof thirdPartyId !== 'string') {
      throw new Error(
        `TeamJoinRequestFixedQueryByThirdPartyUserId: thirdPartyId must be a string: ${thirdPartyId}`
      );
    }
  }
}

export class TeamJoinRequestFixedQueryAllActiveRequests
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.AllActiveTeamJoinApprovals;
  constructor() {}
}

EntityMetadataMappings.Register(
  type,
  TableSettings.TableQueries,
  (query: IEntityMetadataFixedQuery, fixedPartitionKey: string) => {
    switch (query.fixedQueryType) {
      case FixedQueryType.ActiveTeamJoinApprovalsByTeams:
        const { ids } = query as TeamJoinRequestFixedQueryByTeams;
        if (!ids || !Array.isArray(ids)) {
          throw new Error('ids must be an Array');
        }
        const qids = new azure.TableQuery()
          .where('PartitionKey eq ?', fixedPartitionKey)
          .and('tickettype eq ?string?', 'joinTeam')
          .and('active eq ?', true);
        const args = [
          ids
            .map((id) => {
              return 'teamid eq ?string?';
            })
            .join(' or '),
        ].concat(ids);
        const temp = qids.and.apply(qids, args);
        return temp;

      case FixedQueryType.AllActiveTeamJoinApprovals:
        return new azure.TableQuery()
          .where('PartitionKey eq ?', fixedPartitionKey)
          .and('tickettype eq ?string?', 'joinTeam')
          .and('active eq ?', true);

      case FixedQueryType.AllTeamJoinApprovals:
        return new azure.TableQuery()
          .where('PartitionKey eq ?', fixedPartitionKey)
          .and('tickettype eq ?string?', 'joinTeam');

      case FixedQueryType.ActiveTeamJoinApprovalsByTeam:
        const { id } = query as TeamJoinRequestFixedQueryByTeam;
        if (!id) {
          throw new Error('id required');
        }
        const qid = new azure.TableQuery()
          .where('PartitionKey eq ?', fixedPartitionKey)
          .and('tickettype eq ?string?', 'joinTeam')
          .and('active eq ?', true)
          .and('teamid eq ?string?', id);
        return qid;

      case FixedQueryType.ActiveTeamJoinApprovalsByThirdPartyId:
        const {
          thirdPartyId,
        } = query as TeamJoinRequestFixedQueryByThirdPartyUserId;
        if (!thirdPartyId) {
          throw new Error('thirdPartyId required');
        }
        const qtpid = new azure.TableQuery()
          .where('PartitionKey eq ?', fixedPartitionKey)
          .and('tickettype eq ?string?', 'joinTeam')
          .and('active eq ?', true)
          .and('ghid eq ?string?', thirdPartyId);
        return qtpid;

      default:
        throw new Error(
          `The fixed query type "${query.fixedQueryType}" is not implemented by this provider for the type ${type}, or is of an unknown type`
        );
    }
  }
);

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
    const entityTypeColumn = mapMetadataPropertiesToFields['entityType'];
    const entityTypeValue = getEntityTypeColumnValue(type);
    let sql = '',
      values = [];
    switch (query.fixedQueryType) {
      case FixedQueryType.ActiveTeamJoinApprovalsByTeams:
        const { ids } = query as TeamJoinRequestFixedQueryByTeams;
        if (!ids || !Array.isArray(ids)) {
          throw new Error('ids must be an Array');
        }
        let valueCounter = 1;
        const groupSet = ids
          .map((skip) => {
            return '$' + ++valueCounter;
          })
          .join(', ');
        sql = `
        SELECT *
        FROM ${tableName}
        WHERE
          ${entityTypeColumn} = $1 AND
          ${metadataColumnName}->>'teamid' IN ( ${groupSet} ) AND
          ${metadataColumnName} @> \$${++valueCounter}
      `;
        values = [
          entityTypeValue,
          ...stringOrNumberArrayAsStringArray(ids),
          { active: true },
        ];
        return { sql, values };
      case FixedQueryType.AllTeamJoinApprovals:
        return PostgresGetAllEntities(
          tableName,
          entityTypeColumn,
          entityTypeValue
        );
      case FixedQueryType.AllActiveTeamJoinApprovals:
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            active: true,
          }
        );
      case FixedQueryType.ActiveTeamJoinApprovalsByTeam: {
        const { id } = query as TeamJoinRequestFixedQueryByTeam;
        if (!id) {
          throw new Error('id required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            active: true,
            teamid: stringOrNumberAsString(id),
          }
        );
      }
      case FixedQueryType.ActiveTeamJoinApprovalsByThirdPartyId: {
        const {
          thirdPartyId,
        } = query as TeamJoinRequestFixedQueryByThirdPartyUserId;
        if (!thirdPartyId) {
          throw new Error('thirdPartyId required');
        }
        return PostgresJsonEntityQuery(
          tableName,
          entityTypeColumn,
          entityTypeValue,
          metadataColumnName,
          {
            active: true,
            thirdpartyid: stringOrNumberAsString(thirdPartyId),
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
    function translatedField(type: EntityMetadataType, key: string): string {
      const mapTeamApprovalObjectToMemoryFields = EntityMetadataMappings.GetDefinition(
        type,
        MemorySettings.MemoryMapping,
        true
      );
      const value = mapTeamApprovalObjectToMemoryFields.get(key);
      if (!value) {
        throw new Error(
          `No translation exists for field ${key} in memory provider`
        );
      }
      return value;
    }
    const columnActive = translatedField(type, Field.active);
    const columnTeamId = translatedField(type, Field.teamId);
    const columnThirdPartyId = translatedField(type, Field.thirdPartyId);

    switch (query.fixedQueryType) {
      case FixedQueryType.ActiveTeamJoinApprovalsByTeams:
        const { ids } = query as TeamJoinRequestFixedQueryByTeams;
        return allInTypeBin.filter((entity) => {
          return (
            entity[columnActive] &&
            entity[columnActive] === true &&
            entity[columnTeamId] &&
            ids.includes(entity[columnTeamId])
          );
        });

      case FixedQueryType.ActiveTeamJoinApprovalsByTeam:
        const { id } = query as TeamJoinRequestFixedQueryByTeam;
        return allInTypeBin.filter((entity) => {
          return (
            entity[columnActive] &&
            entity[columnActive] === true &&
            entity[columnTeamId] &&
            entity[columnTeamId] === id
          );
        });

      case FixedQueryType.ActiveTeamJoinApprovalsByThirdPartyId:
        const {
          thirdPartyId,
        } = query as TeamJoinRequestFixedQueryByThirdPartyUserId;
        return allInTypeBin.filter((entity) => {
          return (
            entity[columnActive] &&
            entity[columnActive] === true &&
            entity[columnThirdPartyId] &&
            entity[columnThirdPartyId] === thirdPartyId
          );
        });

      case FixedQueryType.AllTeamJoinApprovals:
        return allInTypeBin;

      case FixedQueryType.AllActiveTeamJoinApprovals:
        return allInTypeBin.filter((entity) => {
          return entity[columnActive] && entity[columnActive] === true;
        });

      default:
        throw new Error(
          'fixed query type not implemented in the memory provider'
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
