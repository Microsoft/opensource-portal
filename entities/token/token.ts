//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import azure from 'azure-storage';
import crypto from 'crypto';

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
  IEntityMetadataFixedQuery,
  FixedQueryType,
} from '../../lib/entityMetadataProvider/query';
import { TokenGenerator } from './tokenGenerator';
import { QueryTokensByCorporateID } from './tokenProvider';
import { Type } from './type';
import { TableSettings } from '../../lib/entityMetadataProvider/table';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = Type;

interface ITokenEntityProperties {
  token: any;
  active: any;
  corporateId: any;
  created: any;
  description: any;
  expires: any;
  source: any;
  organizationScopes: any;
  warning: any;
  scopes: any;
}

const Field: ITokenEntityProperties = {
  token: 'token',
  active: 'active',
  created: 'created',
  corporateId: 'corporateId',
  description: 'description',
  expires: 'expires',
  source: 'source',
  warning: 'warning',
  organizationScopes: 'organizationScopes',
  scopes: 'scopes',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class PersonalAccessToken
  implements IObjectWithDefinedKeys, ITokenEntityProperties {
  private _key: string;

  token: string;

  active: boolean;
  warning: string;
  corporateId: string;
  created: Date;
  description: string;
  expires: Date;
  source: string;
  organizationScopes: string;
  scopes: any;

  displayUsername?: string; // not a traditional field, just for VSTS

  constructor() {
    this.created = new Date();
  }

  static CreateFromAadAuthorization({
    appId,
    oid,
    scopes,
    organizationScopes,
  }): PersonalAccessToken {
    const pat = new PersonalAccessToken();
    pat.corporateId = null;
    pat.description = `AAD oid ${oid} app ${appId} with scopes ${scopes}`;
    pat.source = `AAD oid ${oid} app ${appId}`;
    pat.displayUsername = 'AAD Identity';
    pat.organizationScopes = organizationScopes;
    pat.scopes = scopes;
    return pat;
  }

  static CreateFromAzureDevOpsTokenAuthorization({
    corporateId,
    description,
    displayUsername,
    source,
    scopes,
  }): PersonalAccessToken {
    const pat = new PersonalAccessToken();
    pat.corporateId = corporateId;
    pat.description = description;
    pat.source = source;
    pat.displayUsername = displayUsername;
    pat.scopes = scopes;
    return pat;
  }

  static CreateNewToken(): PersonalAccessToken {
    const pat = new PersonalAccessToken();
    const { key, token } = TokenGenerator.Generate();
    pat.token = token;
    pat._key = key;
    return pat;
  }

  getObjectFieldNames(): string[] {
    return fieldNames;
  }

  getPrivateKey(): string {
    return this._key;
  }

  isRevoked(): boolean {
    return this.active === false;
  }

  getIdentifier() {
    const concat = this.created + this.token;
    return crypto
      .createHash('sha1')
      .update(concat)
      .digest('hex')
      .substring(0, 10);
  }

  isExpired(): boolean {
    // NOTE: service to service tokens do not expire, so
    // a token without an expiration is valid in this env.
    if (!this.expires) {
      return false;
    }
    const now = new Date();
    return this.expires < now;
  }

  hasScope(scope: string) {
    if (!this.scopes) {
      return false;
    }
    const apis = this.scopes.toLowerCase().split(',');
    return apis.includes(scope.toLowerCase());
  }

  hasOrganizationScope(orgName: string) {
    if (!this.organizationScopes) {
      return false;
    }
    if (this.organizationScopes === '*') {
      return true;
    }
    const orgList = this.organizationScopes.toLowerCase().split(',');
    return orgList.includes(orgName.toLowerCase());
  }
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new PersonalAccessToken();
  }
);
EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityIdColumnName,
  Field.token
);

EntityMetadataMappings.Register(
  type,
  TableSettings.TableMapping,
  new Map<string, string>([
    [Field.token, null], // RowKey
    [Field.active, 'active'],
    [Field.corporateId, 'owner'],
    [Field.created, 'entityCreated'],
    [Field.description, 'description'],
    [Field.source, 'service'],
    [Field.organizationScopes, 'orgs'],
    [Field.expires, 'expires'],
    [Field.warning, 'warning'],
    [Field.scopes, 'apis'],
  ])
);
EntityMetadataMappings.Register(type, TableSettings.TablePossibleDateColumns, [
  Field.created,
  Field.expires,
]);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultTableName,
  'settings'
);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultFixedPartitionKey,
  'apiKey'
);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultRowKeyPrefix,
  'apiKey'
);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultFixedPartitionKeyNoPrefix,
  true
);
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
    [Field.token, Field.token],
    [Field.active, Field.active],
    [Field.corporateId, Field.corporateId],
    [Field.created, Field.created],
    [Field.description, Field.description],
    [Field.source, Field.source],
    [Field.active, Field.active],
    [Field.organizationScopes, Field.organizationScopes],
    [Field.expires, Field.expires],
    [Field.warning, Field.warning],
    [Field.scopes, Field.scopes],
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
  TableSettings.TableQueries,
  (query: IEntityMetadataFixedQuery, fixedPartitionKey: string) => {
    switch (query.fixedQueryType) {
      case FixedQueryType.TokensByCorporateId:
        const { corporateId } = query as QueryTokensByCorporateID;
        if (!corporateId) {
          throw new Error('corporateId required');
        }
        return new azure.TableQuery()
          .where('PartitionKey eq ?', fixedPartitionKey)
          .and('owner eq ?string?', corporateId);
      case FixedQueryType.TokensGetAll:
        return new azure.TableQuery().where(
          'PartitionKey eq ?',
          fixedPartitionKey
        );
      default:
        throw new Error(
          `The fixed query type ${type} is not supported currently by this provider, or is of an unknown type`
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
    const columnCorporateId = translatedField(type, Field.corporateId);
    switch (query.fixedQueryType) {
      case FixedQueryType.TokensByCorporateId:
        const { corporateId } = query as QueryTokensByCorporateID;
        return allInTypeBin.filter((entity) => {
          return (
            entity[columnCorporateId] &&
            entity[columnCorporateId] === corporateId
          );
        });
      case FixedQueryType.TokensGetAll:
        return allInTypeBin;
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
