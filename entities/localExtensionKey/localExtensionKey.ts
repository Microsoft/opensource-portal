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
import { TableSettings } from '../../lib/entityMetadataProvider/table';
import { MemorySettings } from '../../lib/entityMetadataProvider/memory';

const type = new EntityMetadataType('LocalExtensionKey');

const oldestAllowedKeyExpirationMs = 1000 * 60 * 60 * 24 * 14; // 14 days

interface IExtensionKeyEntityProperties {
  corporateId: any;
  localDataKey: any;
  created: any;
}

const Field: IExtensionKeyEntityProperties = {
  corporateId: 'corporateId',
  localDataKey: 'localDataKey',
  created: 'created',
};

const fieldNames = Object.getOwnPropertyNames(Field);

export class LocalExtensionKey
  implements IObjectWithDefinedKeys, IExtensionKeyEntityProperties {
  corporateId: string;
  created: Date;
  localDataKey: string;

  constructor() {
    this.created = new Date();
  }

  static CreateNewLocalExtensionKey(corporateId: string): LocalExtensionKey {
    const localExtensionKey = new LocalExtensionKey();
    localExtensionKey.localDataKey = crypto.randomBytes(32).toString('base64');
    localExtensionKey.corporateId = corporateId;
    return localExtensionKey;
  }

  isValidNow(): boolean {
    const now = new Date();
    const expires = new Date(
      this.created.getTime() + oldestAllowedKeyExpirationMs
    );
    if (expires < now || !this.localDataKey) {
      return false;
    }
    return true;
  }

  getObjectFieldNames(): string[] {
    return fieldNames;
  }
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new LocalExtensionKey();
  }
);
EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityIdColumnName,
  Field.corporateId
);

EntityMetadataMappings.Register(
  type,
  TableSettings.TableMapping,
  new Map<string, string>([
    [Field.corporateId, null], // RowKey
    [Field.created, 'entityCreated'],
    [Field.localDataKey, Field.localDataKey],
  ])
);
EntityMetadataMappings.Register(type, TableSettings.TablePossibleDateColumns, [
  Field.created,
]);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultTableName,
  'settings'
);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultFixedPartitionKey,
  'localExtensionKey'
);
EntityMetadataMappings.Register(
  type,
  TableSettings.TableDefaultFixedPartitionKeyNoPrefix,
  true
);
EntityMetadataMappings.Register(type, TableSettings.TableEncryptedColumnNames, [
  Field.localDataKey,
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
    [Field.corporateId, Field.corporateId],
    [Field.created, Field.created],
    [Field.localDataKey, Field.localDataKey],
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
      case FixedQueryType.LocalExtensionKeysGetAll:
        return new azure.TableQuery().where(
          'PartitionKey eq ?',
          fixedPartitionKey
        );
      default:
        throw new Error(
          `The fixed query type ${query.fixedQueryType} is not supported currently by this ${type} provider`
        );
    }
  }
);

EntityMetadataMappings.Register(
  type,
  MemorySettings.MemoryQueries,
  (query: IEntityMetadataFixedQuery, allInTypeBin: IEntityMetadata[]) => {
    switch (query.fixedQueryType) {
      case FixedQueryType.LocalExtensionKeysGetAll:
        return allInTypeBin;
      default:
        throw new Error(
          `The fixed query type ${query.fixedQueryType} is not supported currently by this ${type} provider`
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
