//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Azure table storage implementation on top of the generic metadata provider
// interface. The metadata type maps to a table name.

// Opinions of note: any entities passed for storage that have a _number_ will
// be stored in Azure table as a _string_.

// Partitioning: none. For historical reasons around scale, only a fixed pkey
// is used for this product.

import azure from 'azure-storage';
import { v4 as uuidV4 } from 'uuid';

const debugShowTableOperations = false;

const emptyString = '';

import {
  IEntityMetadataProvider,
  IEntityMetadataSerializationHelper,
  IEntityMetadataDeserializationHelper,
  SerializeObjectToEntityMetadata,
  DeserializeEntityMetadataToObjectSetCollection,
  IObjectWithDefinedKeys,
  EntityFieldNames,
} from './entityMetadataProvider';
import {
  IEntityMetadata,
  EntityMetadataType,
  EntityMetadataTypes,
} from './entityMetadata';
import { IEntityMetadataFixedQuery } from './query';
import {
  MetadataMappingDefinition,
  EntityMetadataMappings,
  MetadataMappingDefinitionBase,
} from './declarations';
import {
  encryptTableEntity,
  decryptTableEntity,
  ITableEncryptionOperationOptions,
} from './tableEncryption';

export interface ITableEncryptionOptions {
  keyEncryptionKeyId: string;
  keyResolver: unknown;
}

export interface ITableEntityMetadataProviderOptions {
  metadataTypeToTableNameMapping?: any;
  metadataTypeToFixedPartitionKeyMapping?: any;
  metadataTypeToRowKeyPrefixMapping?: any;
  metadataTypeToEncryptedColumnsMapping?: any;
  account: string;
  key: string;
  prefix?: string;
  encryption?: ITableEncryptionOptions;
}

class TableMetadataDefinition extends MetadataMappingDefinitionBase {
  constructor(name: string) {
    super(name);
  }
}

export const TableSettings = {
  TableMapping: new TableMetadataDefinition('TableMapping'),
  TablePossibleDateColumns: new TableMetadataDefinition(
    'TablePossibleDateColumns'
  ),
  TableQueries: new TableMetadataDefinition('TableQueries'),
  TableNoPointQueries: new TableMetadataDefinition('TableNoPointQueries'),
  TableNoPointQueryMapping: new TableMetadataDefinition(
    'TableNoPointQueryMapping'
  ),
  TableNoPointQueryAlternateIdFieldName: new TableMetadataDefinition(
    'TableNoPointQueryAlternateIdFieldName'
  ),
  TableSpecializedDeserializationHelper: new TableMetadataDefinition(
    'TableSpecializedDeserializationHelper'
  ),
  TableSpecializedSerializationHelper: new TableMetadataDefinition(
    'TableSpecializedSerializationHelper'
  ),
  TableDefaultTableName: new TableMetadataDefinition('TableDefaultTableName'),
  TableDefaultFixedPartitionKey: new TableMetadataDefinition(
    'TableDefaultFixedPartitionKey'
  ),
  TableDefaultFixedPartitionKeyNoPrefix: new TableMetadataDefinition(
    'TableDefaultFixedPartitionKeyNoPrefix'
  ),
  TableDefaultRowKeyPrefix: new TableMetadataDefinition(
    'TableDefaultRowKeyPrefix'
  ),
  TableEncryptedColumnNames: new TableMetadataDefinition(
    'TableEncryptedColumnNames'
  ),
};

const TableClientProperties = new Set([
  'Timestamp',
  'PartitionKey',
  'RowKey',
  '.metadata',
]);

interface ITableEntity {
  RowKey: ITableStringValue;
  PartitionKey: ITableStringValue;
}

interface ITableStringValue {
  _: string;
}

export class TableEntityMetadataProvider implements IEntityMetadataProvider {
  public readonly supportsHistory: boolean = false;
  public readonly name = 'table';

  private _storageAccountName: string;

  private _table: azure.TableService;
  private _entityGenerator: any;

  private _tableNameMapping: any;
  private _fixedPartitionKeyMapping: any;
  private _rowKeyPrefixMapping: any;
  private _typeToEncryptedColumnsMapping: any;
  private _typeToEncryptionOptions: Map<
    EntityMetadataType,
    ITableEncryptionOperationOptions
  >;

  private _encryptionOptions: ITableEncryptionOptions;

  private _prefix: string;

  private _initialized: boolean;
  private _initializedTables: Map<EntityMetadataType, string>;

  constructor(options: ITableEntityMetadataProviderOptions) {
    if (!options) {
      throw new Error('ITableEntityMetadataProviderOptions required');
    }
    this._tableNameMapping = Object.assign(
      defaultTableNames(),
      options.metadataTypeToTableNameMapping || {}
    );

    this._storageAccountName = options.account;
    if (!this._storageAccountName) {
      throw new Error('Storage account name required');
    }
    const storageAccountKey = options.key;
    if (!storageAccountKey) {
      throw new Error('Storage account key required');
    }
    this._prefix = options.prefix || '';
    this._fixedPartitionKeyMapping = Object.assign(
      defaultFixedPartitionKeys(this._prefix),
      options.metadataTypeToFixedPartitionKeyMapping || {}
    );
    this._rowKeyPrefixMapping = Object.assign(
      defaultRowKeyPrefixes(),
      options.metadataTypeToRowKeyPrefixMapping || {}
    );
    this._typeToEncryptedColumnsMapping = Object.assign(
      defaultEncryptionColumns(),
      options.metadataTypeToEncryptedColumnsMapping || {}
    );
    this._typeToEncryptionOptions = new Map();
    this._encryptionOptions = options.encryption;
    try {
      this._table = azure.createTableService(
        this._storageAccountName,
        storageAccountKey
      );
      this._entityGenerator = azure.TableUtilities.entityGenerator;
    } catch (storageAccountError) {
      throw storageAccountError;
    }
    this._initializedTables = new Map();
  }

  supportsPointQueryForType(type: EntityMetadataType): boolean {
    const noPointQueries = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableNoPointQueries,
      false
    );
    if (noPointQueries) {
      // The original implementation of storing repository create
      // metadata in table stored the data in a new row without any
      // ID. Then, a few years later, 'repoId' was set to the new
      // repo ID, which is more durable than repository names which
      // can change. By not supporting point queries for the repository
      // type, a query is used for get operations instead.
      return false;
    }
    return true;
  }

  async initialize(): Promise<void> {
    this._initialized = true;
  }

  // ---

  async getMetadata(
    type: EntityMetadataType,
    id: string
  ): Promise<IEntityMetadata> {
    this.throwIfNotInitialized();
    const tableName = await this.initializeEntityType(type);
    const tableEntity = await this.tableRetrieveEntity(
      type,
      tableName,
      this.getFixedPartitionKey(type),
      this.getRowKey(type, id)
    );
    return this.tableEntityToMetadataObject(type, tableEntity);
  }

  async setMetadata(metadata: IEntityMetadata): Promise<void> {
    this.throwIfNotInitialized();
    const tableName = await this.initializeEntityType(metadata.entityType);
    const tableEntity = this.metadataToTableEntity(metadata);
    const result = await this.tableInsertEntity(
      metadata.entityType,
      tableName,
      tableEntity
    );
  }

  async updateMetadata(metadata: IEntityMetadata): Promise<void> {
    this.throwIfNotInitialized();
    const tableName = await this.initializeEntityType(metadata.entityType);
    const tableEntity = this.metadataToTableEntity(metadata);
    const result = await this.tableReplaceEntity(
      metadata.entityType,
      tableName,
      tableEntity
    );
  }

  async deleteMetadata(metadata: IEntityMetadata): Promise<void> {
    this.throwIfNotInitialized();
    const tableName = await this.initializeEntityType(metadata.entityType);
    await this.tableDeleteEntity(
      tableName,
      this.getFixedPartitionKey(metadata.entityType),
      this.getRowKey(metadata.entityType, metadata.entityId)
    );
  }

  async clearMetadataStore(type: EntityMetadataType): Promise<void> {
    throw new Error('The table provider does not support clearMetadataStore');
  }

  async fixedQueryMetadata(
    type: EntityMetadataType,
    query: IEntityMetadataFixedQuery
  ): Promise<IEntityMetadata[]> {
    this.throwIfNotInitialized();
    const tableName = await this.initializeEntityType(type);
    const azureTableQuery = this.createQueryFromFixedQueryEnum(type, query);
    return await this.tableQueryToMetadataArray(
      type,
      tableName,
      azureTableQuery
    );
  }

  // ---

  private throwIfNotInitialized() {
    if (!this._initialized) {
      throw new Error('TableEntityMetadataProvider is not initialized');
    }
  }

  private getFixedPartitionKey(type: EntityMetadataType) {
    const pk = this._fixedPartitionKeyMapping[type.typeName];
    if (pk) {
      return pk;
    }
    throw new Error(`No fixed partition key is defined for the type: ${type}`);
  }

  private isTypeEncrypted(type: EntityMetadataType): Set<string> {
    if (!this._encryptionOptions) {
      return null;
    }
    const set = this._typeToEncryptedColumnsMapping[
      type.typeName
    ] as Set<string>;
    return set && set.size > 0 ? set : null;
  }

  private getEncryptionOptionsForType(
    type: EntityMetadataType
  ): ITableEncryptionOperationOptions {
    let options = this._typeToEncryptionOptions.get(type);
    if (options) {
      return options;
    }
    if (!this._encryptionOptions) {
      throw new Error('Encryption options required');
    }
    if (!this._encryptionOptions.keyEncryptionKeyId) {
      throw new Error('keyEncryptionKeyId encryption option required');
    }
    if (!this._encryptionOptions.keyResolver) {
      throw new Error('keyResolver encryption option required');
    }
    options = {
      keyEncryptionKeyId: this._encryptionOptions.keyEncryptionKeyId,
      keyResolver: this._encryptionOptions.keyResolver,
      encryptedPropertyNames: this.isTypeEncrypted(type),
      binaryProperties: 'buffer',
    };
    this._typeToEncryptionOptions.set(type, options);
    return options;
  }

  private getRowKey(type: EntityMetadataType, entityId: string) {
    const prefix = this._rowKeyPrefixMapping[type.typeName] || emptyString;
    return `${prefix}${entityId}`;
  }

  private rowKeyToEntityId(type: EntityMetadataType, rowKey: string) {
    const prefix = this._rowKeyPrefixMapping[type.typeName] || emptyString;
    if (!prefix) {
      return rowKey;
    }
    if (rowKey.startsWith(prefix)) {
      return rowKey.substr(prefix.length);
    }
    throw new Error(
      `The entity type ${type} table provider has a defined and expected prefix of ${prefix} that was not present for the entity at row key ${rowKey}`
    );
  }

  private async initializeEntityType(
    type: EntityMetadataType
  ): Promise<string> {
    let tableName = this._initializedTables.get(type);
    if (tableName) {
      return tableName;
    }
    const tableSuffix = this._tableNameMapping[type.typeName];
    if (!tableSuffix) {
      throw new Error(
        `No storage table name mapping provided for value ${type}`
      );
    }
    tableName = `${this._prefix}${tableSuffix}`;
    await this.tableCreateIfNotExists(type, tableName);
    return tableName;
  }

  private metadataToTableEntity(metadata: IEntityMetadata): ITableEntity {
    if (!metadata.entityId) {
      throw new Error(`metadata.entityId must be provided`);
    }
    const rowKey = this.getRowKey(metadata.entityType, metadata.entityId);
    const partitionKey = this.getFixedPartitionKey(metadata.entityType);
    const te: any = {};
    const fieldNames = metadata.entityFieldNames;
    if (!fieldNames || !Array.isArray(fieldNames)) {
      throw new Error('metadata.entityFieldNames is not an array');
    }
    for (let i = 0; i < fieldNames.length; i++) {
      const key = fieldNames[i];
      const value = this.createEntityDescriptorForValue(key, metadata[key]);
      if (value) {
        te[key] = value;
      }
    }
    const entity: ITableEntity = Object.assign(
      te,
      this.createRowEntity(partitionKey, rowKey)
    );
    return entity;
  }

  private createEntityDescriptorForValue(key: string, value: any) {
    if (value === undefined || value === null) {
      // Table provider should not include null/undefined values
      return value;
    }
    if (typeof value === 'string') {
      return this._entityGenerator.String(value);
    } else if (value === true || value === false) {
      return this._entityGenerator.Boolean(value);
    } else if (Buffer.isBuffer(value)) {
      return this._entityGenerator.Binary(value);
    } else if (value instanceof Date) {
      if (!isFinite(value as any)) {
        return undefined;
      }
      return this._entityGenerator.DateTime(value);
    } else if (typeof value === 'number') {
      // NOTE: OPINIONATED: we store all numbers are strings in Azure Table...
      return this._entityGenerator.String(value.toString());
    }
    throw new Error(
      `The key ${key} in the entity is of an unsupported type: ${typeof value}`
    );
  }

  private tableEntityToMetadataObject(
    type: EntityMetadataType,
    tableEntity: any
  ): IEntityMetadata {
    const id = this.rowKeyToEntityId(type, tableEntity.RowKey._);
    const created = tableEntity.Timestamp._;
    const reducedObject = this.reduceTableEntityToObject(tableEntity);
    const keys = Object.getOwnPropertyNames(reducedObject);
    const newMetadataObject: IEntityMetadata = Object.assign(reducedObject, {
      entityType: type,
      entityId: id,
      entityFieldNames: keys,
      entityCreated: created,
    });
    return newMetadataObject;
  }

  private reduceTableEntityToObject(tableEntity: any): any {
    if (tableEntity === undefined || tableEntity === null) {
      return tableEntity;
    }
    const newObject = {};
    for (let column in tableEntity) {
      if (TableClientProperties.has(column)) {
        // Timestamp, PartitionKey, RowKey, .metadata
        continue;
      }
      if (tableEntity[column] && tableEntity[column]._ !== undefined) {
        newObject[column] = tableEntity[column]._;
      }
    }
    return newObject;
  }

  private createRowEntity(partitionKey: string, rowKey: string): ITableEntity {
    return {
      PartitionKey: this._entityGenerator.String(partitionKey),
      RowKey: this._entityGenerator.String(rowKey),
    };
  }

  private async tableQueryToMetadataArray(
    type: EntityMetadataType,
    tableName: string,
    azureTableQuery: any
  ): Promise<IEntityMetadata[]> {
    const tableEntities = await this.tableQueryAllEntities(
      type,
      tableName,
      azureTableQuery
    );
    return tableEntities.map((tableEntity) => {
      return this.tableEntityToMetadataObject(type, tableEntity);
    });
  }

  private async tableQueryAllEntities(
    type: EntityMetadataType,
    tableName: string,
    azureTableQuery: azure.TableQuery
  ): Promise<any[]> {
    const rows = [];
    let continuationToken = null;
    if (debugShowTableOperations) {
      displayTableQuery(
        'TABLE QUERY',
        this._storageAccountName,
        tableName,
        azureTableQuery,
        null,
        debugShowTableOperations
      );
    }
    do {
      const resultSet = await this.tableQueryEntities(
        type,
        tableName,
        azureTableQuery,
        continuationToken
      );
      if (resultSet.entries) {
        rows.push(...resultSet.entries);
      }
      continuationToken = resultSet.continuationToken;
    } while (continuationToken !== null);
    if (debugShowTableOperations) {
      console.log(`Total rows returned: ${rows.length}`);
    }
    return rows;
  }

  private async tableQueryEntities(
    type: EntityMetadataType,
    tableName: string,
    azureTableQuery: any,
    continuationToken?: string
  ): Promise<azure.TableService.QueryEntitiesResult<ITableEntity>> {
    const results = await this.tableQueryEntitiesAsync(
      tableName,
      azureTableQuery,
      continuationToken
    );
    const encryptedColumnNames = this.isTypeEncrypted(type);
    if (encryptedColumnNames) {
      const encryptionOptions = this.getEncryptionOptionsForType(type);
      const decryptedEntries = await this.decryptQueryResults(
        type,
        encryptionOptions,
        results.entries
      );
      results.entries = decryptedEntries;
    }
    return results;
  }

  private async decryptQueryResults(
    type: EntityMetadataType,
    encryptionOptions: ITableEncryptionOperationOptions,
    results: ITableEntity[]
  ): Promise<ITableEntity[]> {
    const decrypted = [];
    // sync; throat may be a better call if encrypted queries are common
    for (let i = 0; i < results.length; i++) {
      const entity = results[i];
      const decryptedEntity = await decryptTableEntity(
        entity.PartitionKey._,
        entity.RowKey._,
        entity,
        encryptionOptions
      );
      decrypted.push(decryptedEntity);
    }
    return decrypted;
  }

  private async tableReplaceEntity(
    type: EntityMetadataType,
    tableName: string,
    entity: ITableEntity
  ): Promise<any> {
    const encryptedColumnNames = this.isTypeEncrypted(type);
    if (encryptedColumnNames) {
      entity = await encryptTableEntity(
        entity.PartitionKey._,
        entity.RowKey._,
        entity,
        this.getEncryptionOptionsForType(type)
      );
    }
    const result = await this.tableReplaceEntityAsync(tableName, entity);
    if (debugShowTableOperations) {
      displayTableRow(
        'TABLE REPLACE',
        this._storageAccountName,
        tableName,
        entity,
        result['.metadata'].etag,
        debugShowTableOperations
      );
    }
    return result;
  }

  private async tableInsertEntity(
    type: EntityMetadataType,
    tableName: string,
    entity: ITableEntity
  ): Promise<any> {
    const encryptedColumnNames = this.isTypeEncrypted(type);
    if (encryptedColumnNames) {
      entity = await encryptTableEntity(
        entity.PartitionKey._,
        entity.RowKey._,
        entity,
        this.getEncryptionOptionsForType(type)
      );
    }
    const result = await this.tableInsertEntityAsync(tableName, entity);
    displayTableRow(
      'TABLE INSERT',
      this._storageAccountName,
      tableName,
      entity,
      result['.metadata'].etag,
      debugShowTableOperations
    );
    return result;
  }

  private async tableRetrieveEntity(
    type: EntityMetadataType,
    tableName: string,
    partitionKey: string,
    rowKey: string
  ): Promise<any> {
    let tableEntity = await this.tableRetrieveEntityAsync(
      type,
      tableName,
      partitionKey,
      rowKey
    );
    const encryptedColumnNames = this.isTypeEncrypted(type);
    if (encryptedColumnNames) {
      tableEntity = await decryptTableEntity(
        partitionKey,
        rowKey,
        tableEntity,
        this.getEncryptionOptionsForType(type)
      );
    }
    displayTableRow(
      'TABLE GET',
      this._storageAccountName,
      tableName,
      tableEntity,
      encryptedColumnNames ? '(decrypted)' : '',
      debugShowTableOperations
    );
    return tableEntity;
  }

  private async tableDeleteEntity(
    tableName: string,
    partitionKey: string,
    rowKey: string
  ): Promise<any> {
    const entity = this.createRowEntity(partitionKey, rowKey);
    return new Promise((resolve, reject) => {
      this._table.deleteEntity(tableName, entity, (error, result) => {
        return error ? reject(error) : resolve(result);
      });
    });
  }

  private async tableCreateIfNotExists(
    type: EntityMetadataType,
    tableName: string
  ): Promise<boolean> {
    return new Promise<boolean>((resolve, reject) => {
      this._table.createTableIfNotExists(tableName, (error, result) => {
        if (!error) {
          this._initializedTables.set(type, tableName);
        }
        return error ? reject(error) : resolve(result.created);
      });
    });
  }

  private async tableDeleteEntireTable(
    type: EntityMetadataType,
    tableName: string
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      // WARNING: a table can technically support multiple types of metadata.
      // This call may be more destructive than you would like as it can
      // delete other types.
      this._table.deleteTable(tableName, (error, result) => {
        if (!error) {
          this._initializedTables.delete(type);
        }
        return error ? reject(error) : resolve(result);
      });
    });
  }

  private createQueryFromFixedQueryEnum(
    type: EntityMetadataType,
    query: IEntityMetadataFixedQuery
  ): any {
    let get = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableQueries,
      true
    );
    return get(query, this.getFixedPartitionKey(type));
  }

  getSerializationHelper(
    type: EntityMetadataType
  ): IEntityMetadataSerializationHelper {
    const tableMapping = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableMapping,
      true
    );
    if (!tableMapping) {
      return null;
    }
    const alternateIdFieldName = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableNoPointQueryAlternateIdFieldName,
      false
    );
    const supportsPointQuery = this.supportsPointQueryForType(type);
    const idFieldName = supportsPointQuery
      ? EntityMetadataMappings.GetDefinition(
          type,
          MetadataMappingDefinition.EntityIdColumnName,
          false
        )
      : alternateIdFieldName;
    const noPointQueryMapObjectToTableFields = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableNoPointQueryMapping,
      false
    );
    const mapObjectToTableFields =
      !supportsPointQuery && noPointQueryMapObjectToTableFields
        ? mergeMaps(tableMapping, noPointQueryMapObjectToTableFields)
        : tableMapping;
    const specializedSerializer = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableSpecializedSerializationHelper,
      false
    );
    return function objectToTable(object: any): IEntityMetadata {
      if (!supportsPointQuery && !object[idFieldName]) {
        // CONSIDER: it might be best if the serialization helpers took options, and so only when
        // a new object is being created through the appropriate code path would this identify
        // get created.
        object[idFieldName] = uuidV4(); // new entity to insert
      }
      const entity = SerializeObjectToEntityMetadata(
        type,
        idFieldName,
        object,
        mapObjectToTableFields,
        true /* numbers to strings */,
        false /* throw if missing translations */,
        true /* ignore private variables */
      );
      if (specializedSerializer) {
        specializedSerialize(entity, object, specializedSerializer);
      }
      return entity;
    };
  }

  getDeserializationHelper(
    type: EntityMetadataType
  ): IEntityMetadataDeserializationHelper {
    const tableMapping = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableMapping,
      true
    );
    if (!tableMapping) {
      return null;
    }
    const idFieldName = EntityMetadataMappings.GetDefinition(
      type,
      MetadataMappingDefinition.EntityIdColumnName,
      true
    );
    const alternateIdFieldName = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableNoPointQueryAlternateIdFieldName,
      false
    );
    const possibleDateColumnNames = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TablePossibleDateColumns,
      false
    );
    const noPointQueryMapObjectToTableFields = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableNoPointQueryMapping,
      false
    );
    const supportsPointQuery = this.supportsPointQueryForType(type);
    const specializedDeserializer = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableSpecializedDeserializationHelper,
      false
    );
    const mapObjectToTableFields =
      !supportsPointQuery && noPointQueryMapObjectToTableFields
        ? mergeMaps(tableMapping, noPointQueryMapObjectToTableFields)
        : tableMapping;
    return function tableEntityToObject(entity: IEntityMetadata): any {
      const object = EntityMetadataMappings.InstantiateObject(
        type
      ) as IObjectWithDefinedKeys;
      const toSet = DeserializeEntityMetadataToObjectSetCollection(
        entity,
        idFieldName,
        mapObjectToTableFields
      );
      const fieldSet = new Set(
        object.getObjectFieldNames
          ? object.getObjectFieldNames()
          : Object.getOwnPropertyNames(toSet)
      ); // IObjectWithDefinedKeys is not required
      for (const property in toSet) {
        let value = toSet[property];
        if (possibleDateColumnNames.includes(property)) {
          if (typeof value === 'string' || value! instanceof Date) {
            let date = tryGetDate(value);
            if (date && !isFinite((date as any) as number)) {
              date = null;
            }
            if (date) {
              value = date;
            } else {
              value = undefined;
            }
          }
        }
        object[property] = value;
        if (!fieldSet.has(property) && property !== alternateIdFieldName) {
          object.getObjectFieldNames().push(property);
        }
      }
      if (specializedDeserializer) {
        specializedDeserialize(entity, object, specializedDeserializer);
      }
      return object;
    };
  }

  private async tableReplaceEntityAsync(
    tableName: string,
    entity: any
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      this._table.replaceEntity(tableName, entity, (error, result) => {
        return error ? reject(error) : resolve(result);
      });
    });
  }

  private async tableInsertEntityAsync(
    tableName: string,
    entity: any
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      this._table.insertEntity(tableName, entity, (error, result) => {
        return error ? reject(error) : resolve(result);
      });
    });
  }

  private async tableRetrieveEntityAsync(
    type: EntityMetadataType,
    tableName: string,
    partitionKey: string,
    rowKey: string
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      this._table.retrieveEntity(
        tableName,
        partitionKey,
        rowKey,
        (error, tableEntity) => {
          return error ? reject(error) : resolve(tableEntity);
        }
      );
    });
  }

  private async tableQueryEntitiesAsync(
    tableName: string,
    azureTableQuery: any,
    continuationToken?: string
  ): Promise<azure.TableService.QueryEntitiesResult<ITableEntity>> {
    return new Promise((resolve, reject) => {
      this._table.queryEntities(
        tableName,
        azureTableQuery,
        ((continuationToken ||
          null) as unknown) as azure.TableService.TableContinuationToken,
        (
          error,
          results: azure.TableService.QueryEntitiesResult<ITableEntity>
        ) => {
          return error ? reject(error) : resolve(results);
        }
      );
    });
  }
}

function specializedDeserialize(
  entity: IEntityMetadata,
  object: any,
  specializedDeserializer
) {
  try {
    specializedDeserializer(entity, object);
  } catch (deserializationError) {
    const error = new Error(
      `Specialized deserializer defined for type ${entity.entityType} failed deserializing an entity id=${entity.entityId}: ${deserializationError.message}`
    );
    error['inner'] = deserializationError;
    error['entity'] = entity;
    throw error;
  }
}

function specializedSerialize(
  entity: IEntityMetadata,
  object: any,
  specializedSerializer
) {
  try {
    specializedSerializer(entity, object);
    const currentKeys = new Set(entity.entityFieldNames);
    const keys = Object.getOwnPropertyNames(entity);
    keys.forEach((key) => {
      if (
        !currentKeys.has(key) &&
        !EntityFieldNames.has(key) &&
        !key.startsWith('_') /* no private vars */
      ) {
        entity.entityFieldNames.push(key); // net new field added
      }
    });
  } catch (serializationError) {
    const error = new Error(
      `Specialized serializer defined for type ${entity.entityType} failed serializing to an entity id=${entity.entityId}: ${serializationError.message}`
    );
    error['inner'] = serializationError;
    error['entity'] = entity;
    throw error;
  }
}

function mergeMaps(mapA: Map<string, string>, mapB: Map<string, string>) {
  return new Map([...mapA, ...mapB]);
}

function tryGetDate(value) {
  if (!value) {
    return;
  }
  if (typeof value === 'string' && !isNaN(value as any)) {
    value = Number(value);
  }
  try {
    const date = new Date(value);
    if (isFinite((date as any) as number)) {
      return date;
    }
  } catch (dateError) {
    console.dir(dateError);
  }
  console.warn(`value did not produce a valid date ${value}`);
  // undefined
}

function displayTableQuery(
  header: string,
  accountName: string,
  tableName: string,
  query: azure.TableQuery,
  footer?: string,
  output: boolean = true
) {
  const entity = {};
  if (query['_where'] && Array.isArray(query['_where'])) {
    const w = query['_where'] as string[];
    entity['where'] = { _: w.join('') };
  } else {
    entity['query'] = { _: query.toString() };
  }
  displayTableRow(header, accountName, tableName, entity, footer, output);
}

function displayTableRow(
  header: string,
  accountName: string,
  tableName: string,
  entity: any,
  footer?: string,
  output: boolean = true
) {
  if (!output) {
    return;
  }
  const log = [];
  const keys = Object.getOwnPropertyNames(entity);
  const list = [];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (key === '.metadata') {
      continue;
    }
    const value = entity[key]['_']; // type would be entity['$']
    list.push(`${key}=${value}`);
  }
  if (header) {
    log.push(header);
    log.push('-'.repeat(header.length));
  }
  const tableInfo = `${accountName}/${tableName}`;
  log.push(tableInfo);
  log.push('='.repeat(tableInfo.length));
  log.push(list.join(', '));
  if (footer) {
    log.push('-'.repeat(footer.length));
    log.push(footer);
  }
  console.log(log.join('\n'));
  console.log();
}

function defaultRowKeyPrefixes() {
  const defaults = {};
  EntityMetadataTypes.forEach((type) => {
    if (
      !EntityMetadataMappings.GetDefinition(
        type,
        TableSettings.TableMapping,
        false
      )
    ) {
      return;
    }
    const rowKeyPrefix = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableDefaultRowKeyPrefix,
      false
    );
    defaults[type.typeName] = rowKeyPrefix;
  });
  return defaults;
}

function defaultFixedPartitionKeys(prefix: string) {
  const defaults = {};
  EntityMetadataTypes.forEach((type) => {
    try {
      if (
        !EntityMetadataMappings.GetDefinition(
          type,
          TableSettings.TableMapping,
          false
        )
      ) {
        return;
      }
      const partitionKey = EntityMetadataMappings.GetDefinition(
        type,
        TableSettings.TableDefaultFixedPartitionKey,
        true
      );
      const skipPrefix = EntityMetadataMappings.GetDefinition(
        type,
        TableSettings.TableDefaultFixedPartitionKeyNoPrefix,
        false
      );
      defaults[type.typeName] = skipPrefix
        ? partitionKey
        : `${prefix || ''}${partitionKey}`;
    } catch (error) {
      throw new Error(
        `No default Azure table fixed partition key defined for type ${type}`
      );
    }
  });
  return defaults;
}

function defaultTableNames() {
  const defaults = {};
  EntityMetadataTypes.forEach((type) => {
    try {
      if (
        !EntityMetadataMappings.GetDefinition(
          type,
          TableSettings.TableMapping,
          false
        )
      ) {
        return;
      }
      const tableName = EntityMetadataMappings.GetDefinition(
        type,
        TableSettings.TableDefaultTableName,
        true
      );
      defaults[type.typeName] = tableName;
    } catch (error) {
      throw new Error(
        `No default Azure table name defined for type ${type} (${error})`
      );
    }
  });
  return defaults;
}

function defaultEncryptionColumns() {
  const defaults = {};
  EntityMetadataTypes.forEach((type) => {
    if (
      !EntityMetadataMappings.GetDefinition(
        type,
        TableSettings.TableMapping,
        false
      )
    ) {
      return;
    }
    const encryptedColumnNames = EntityMetadataMappings.GetDefinition(
      type,
      TableSettings.TableEncryptedColumnNames,
      false
    );
    if (encryptedColumnNames) {
      defaults[type.typeName] = new Set(encryptedColumnNames);
    }
  });
  return defaults;
}
