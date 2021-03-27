//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Note: numbers are cast to strings

import { Pool as PostgresPool } from 'pg';

import {
  IEntityMetadataProvider,
  IEntityMetadataSerializationHelper,
  IEntityMetadataDeserializationHelper,
  SerializeObjectToEntityMetadata,
  DeserializeEntityMetadataToObjectSetCollection,
} from './entityMetadataProvider';
import {
  IEntityMetadata,
  EntityMetadataType,
  EntityMetadataTypes,
} from './entityMetadata';
import {
  PostgresPoolQuerySingleRowAsync,
  PostgresPoolQueryAsync,
} from '../postgresHelpers';
import { IEntityMetadataFixedQuery } from './query';
import {
  EntityMetadataMappings,
  MetadataMappingDefinition,
  MetadataMappingDefinitionBase,
} from './declarations';
import { CreateError, IDictionary } from '../../transitional';

const MetadataColumnName = 'metadata';

const MapMetadataPropertiesToFields: any = {
  entityCreated: null,
  entityFieldNames: null,
  entityId: 'entityid',
  entityType: 'entitytype',
};

class PostgresMetadataDefinition extends MetadataMappingDefinitionBase {
  constructor(name: string) {
    super(name);
  }
}

export const PostgresSettings = {
  PostgresQueries: new PostgresMetadataDefinition('PostgresQueries'),
  PostgresDefaultTypeColumnName: new PostgresMetadataDefinition(
    'PostgresDefaultTypeColumnName'
  ),
  PostgresDateColumns: new PostgresMetadataDefinition('PostgresDateColumns'),
};

interface IExtractedFieldObjects {
  native: IDictionary<any>;
  metadata: IDictionary<any>;
  leftovers: IDictionary<any>;
}

interface IExtractedRowObjects extends IExtractedFieldObjects {}

class PostgresInternals {
  private static _byType: Map<
    EntityMetadataType,
    PostgresInternals
  > = new Map();

  static instance(type: EntityMetadataType) {
    let instance = PostgresInternals._byType.get(type);
    if (!instance) {
      instance = new PostgresInternals(type);
      PostgresInternals._byType.set(type, instance);
    }
    return instance;
  }

  #type: EntityMetadataType;
  nativeFieldNames: string[];
  defaultTableName: string;
  mapFieldsToColumnNames: Map<string, string>;

  #nativeColumnsSet: Set<string>;
  #metadataToFieldNamesSet: Set<string>;
  #mapNativeFieldsSet: Set<string>; // CUT:

  constructor(type: EntityMetadataType) {
    this.#type = type;
    this.nativeFieldNames = [];
    this.mapFieldsToColumnNames = new Map();
  }

  hasNativeProperties() {
    return this.getNativeColumnsAsSet().size > 0;
  }

  getNativePropertyNamesAsSet() {
    // CUT:
    if (!this.#mapNativeFieldsSet) {
      this.#mapNativeFieldsSet = new Set<string>(this.nativeFieldNames);
    }
    return this.#mapNativeFieldsSet;
  }

  getNativeColumnsAsSet() {
    if (!this.#nativeColumnsSet) {
      const mapping = this.mapFieldsToColumnNames;
      this.#nativeColumnsSet = new Set(
        this.nativeFieldNames.map((fieldName) => mapping.get(fieldName))
      );
    }
    return this.#nativeColumnsSet;
  }

  getMetadataColumnsAsSet() {
    if (!this.#metadataToFieldNamesSet) {
      const nativeColumns = this.getNativeColumnsAsSet();
      const metadataColumns = Array.from(
        this.mapFieldsToColumnNames.values()
      ).filter((column) => !nativeColumns.has(column));
      this.#metadataToFieldNamesSet = new Set(metadataColumns);
    }
    return this.#metadataToFieldNamesSet;
  }

  getIdFieldName(throwIfMissing: boolean) {
    const idFieldName = EntityMetadataMappings.GetDefinition(
      this.#type,
      MetadataMappingDefinition.EntityIdColumnName,
      throwIfMissing
    );
    return idFieldName;
  }

  slicePropertiesBySet(
    obj: IDictionary<any>,
    leftovers: IDictionary<any>,
    recognizedPropertyNames: Set<string>,
    ignoredLeftovers: Set<string>
  ) {
    const properties = Object.getOwnPropertyNames(obj);
    for (const property of properties) {
      if (!recognizedPropertyNames.has(property)) {
        if (!ignoredLeftovers.has(property)) {
          // console.log(`Additional property: ${property}`);
          leftovers[property] = obj[property];
        }
        delete obj[property];
      }
    }
    return obj;
  }

  extractEntityToRowObjects(
    serializedAlready: IEntityMetadata,
    throwIfLeftovers: boolean
  ): IExtractedRowObjects {
    const internals = PostgresInternals.instance(serializedAlready.entityType);
    const natives = internals.getNativeColumnsAsSet();
    //const nativeFieldSet = internals.getNativePropertyNamesAsSet();
    const obj = { ...serializedAlready };
    delete obj.entityCreated;
    delete obj.entityFieldNames;
    delete obj.entityId;
    delete obj.entityType;
    const metadata = {};
    const native = this.slicePropertiesBySet(obj, metadata, natives, new Set());
    const leftovers = {};
    this.slicePropertiesBySet(
      metadata,
      leftovers,
      internals.getMetadataColumnsAsSet(),
      new Set()
    );
    if (throwIfLeftovers) {
      const names = Object.getOwnPropertyNames(leftovers);
      if (names.length) {
        throw new Error(
          `Entity to Postgres native column extraction identified left-over fields: ${names.join(
            ', '
          )}`
        );
      }
    }
    return {
      native,
      metadata,
      leftovers,
    };
  }

  extractRowToFieldObjects(
    obj: any,
    metadataColumnName: string,
    throwIfLeftovers: boolean
  ): IExtractedFieldObjects {
    const nativeSet = this.getNativeColumnsAsSet();
    const metadataSet = this.getMetadataColumnsAsSet();
    const leftovers = {};
    let metadata = { ...obj[metadataColumnName] };
    const nativeClone = { ...obj };
    delete nativeClone[metadataColumnName];
    metadata = this.slicePropertiesBySet(
      metadata,
      leftovers,
      metadataSet,
      nativeSet
    );
    const native = this.slicePropertiesBySet(
      nativeClone,
      leftovers,
      nativeSet,
      metadataSet
    );
    if (throwIfLeftovers) {
      const names = Object.getOwnPropertyNames(leftovers);
      if (names.length) {
        throw new Error(
          `Postgres field extraction identified left-over fields: ${names.join(
            ', '
          )}`
        );
      }
    }
    return { native, metadata, leftovers };
  }
}

export class PostgresConfiguration {
  static IdentifyNativeFields(type: EntityMetadataType, fieldNames: string[]) {
    PostgresInternals.instance(type).nativeFieldNames = fieldNames;
  }

  static SetDefaultTableName(type: EntityMetadataType, tableName: string) {
    PostgresInternals.instance(type).defaultTableName = tableName;
  }

  static MapFieldsToColumnNames(
    type: EntityMetadataType,
    map: Map<string, string>,
    lowercaseColumnNamesAutomatically?: boolean
  ) {
    const dest = PostgresInternals.instance(type).mapFieldsToColumnNames;
    for (const [key, value] of map.entries()) {
      dest.set(
        key,
        lowercaseColumnNamesAutomatically ? value.toLowerCase() : value
      );
    }
  }

  static MapFieldsToColumnNamesFromListLowercased(
    type: EntityMetadataType,
    fieldNames: string[]
  ) {
    PostgresConfiguration.MapFieldsToColumnNames(
      type,
      new Map(
        fieldNames.map((fieldName) => {
          return [fieldName, fieldName];
        })
      ),
      true
    );
  }

  static ValidateMappings(
    type: EntityMetadataType,
    fieldNames: string[],
    permittedAdditionalUnvisitedMappings: string[]
  ) {
    const { mapFieldsToColumnNames } = PostgresInternals.instance(type);
    const dynamicType = new PostgresMetadataDefinition('PostgresMapping');
    dynamicType.alternateRuntimeValidateMapping = mapFieldsToColumnNames;
    EntityMetadataMappings.RuntimeValidateMappings(
      type,
      dynamicType,
      fieldNames,
      permittedAdditionalUnvisitedMappings
    );
  }
}

interface IPostgresGetQueries {
  (
    query: IEntityMetadataFixedQuery,
    mapMetadataPropertiesToFields: string[],
    metadataColumnName: string,
    tableName: string,
    getEntityTypeColumnValue: any
  ): any;
}

export interface IPostgresQuery {
  sql: string;
  values: any;
}

export function PostgresGetAllEntities(
  tableName: string,
  entityTypeColumn: string,
  entityTypeValue: string
): IPostgresQuery {
  const sql = `
    SELECT * FROM ${tableName} WHERE
      ${entityTypeColumn} = $1`;
  const values = [entityTypeValue];
  return { sql, values };
}

export function PostgresGetByID(
  tableName: string,
  entityTypeColumn: string,
  entityTypeValue: string,
  entityIdColumn: string,
  idValue: string
): IPostgresQuery {
  const sql = `
    SELECT * FROM ${tableName} WHERE
      ${entityTypeColumn} = $1 AND
      ${entityIdColumn} = $2`;
  const values = [entityTypeValue, idValue];
  return { sql, values };
}

export function PostgresJsonEntityQuery(
  tableName: string,
  entityTypeColumn: string,
  entityTypeValue: string,
  metadataColumnName: string,
  jsonQueryObject: any,
  optionalOrderFieldName?: string,
  isDescending?: boolean
): IPostgresQuery {
  const orderBy = optionalOrderBy(
    metadataColumnName,
    optionalOrderFieldName,
    3,
    isDescending
  );
  const sql = `SELECT * FROM ${tableName} WHERE ${entityTypeColumn} = $1 AND
      ${metadataColumnName} @> $2 ${orderBy.sql}`;
  const values = [entityTypeValue, jsonQueryObject, ...orderBy.values];
  return { sql, values };
}

function optionalOrderBy(
  metadataColumnName,
  optionalOrderFieldName: string,
  variableStartNumber: number,
  isDescending: boolean
) {
  const r = {
    sql: '',
    values: [],
  };
  if (optionalOrderFieldName) {
    const ascend = isDescending ? ' DESC' : 'ASC';
    r.sql = ` ORDER BY ${metadataColumnName}->$${variableStartNumber} ${ascend}`;
    r.values = [optionalOrderFieldName];
  }
  return r;
}

export function PostgresJsonEntityQueryMultiple(
  tableName: string,
  entityTypeColumn: string,
  entityTypeValue: string,
  metadataColumnName: string,
  jsonQueryObjects: any[]
): IPostgresQuery {
  if (jsonQueryObjects.length <= 0) {
    throw new Error(
      'Multi-entity value queries in Postgres require at least 1 query object'
    );
  }
  const values = [entityTypeValue];
  const sqlSet = [];
  for (let i = 0; i < jsonQueryObjects.length; i++) {
    sqlSet.push(`${metadataColumnName} @> $${i + 2}`);
    values.push(jsonQueryObjects[i]);
  }
  const sqlGroup = sqlSet.join(' OR ');
  let sql = `SELECT * FROM ${tableName} WHERE ${entityTypeColumn} = $1 AND ( ${sqlGroup} )`;
  return { sql, values };
}

export interface IPostgresEntityMetadataProviderOptions {
  entityTypeToTableNamesMapping?: any;
  entityTypeToColumnValuesMapping?: any;
  pool: PostgresPool;
}

export class PostgresEntityMetadataProvider implements IEntityMetadataProvider {
  public readonly supportsHistory: boolean = false;
  public readonly name = 'postgres';

  private _pool: PostgresPool;
  private _entityTypeToTableNamesMapping: any;
  private _entityTypeToColumnValuesMapping: any;

  constructor(options: IPostgresEntityMetadataProviderOptions) {
    if (!options) {
      throw new Error('IPostgresEntityMetadataProviderOptions required');
    }
    this._pool = options.pool;
    if (!this._pool) {
      throw new Error(
        'PostgresEntityMetadataProvider requires a Postgres pool'
      );
    }
    this._entityTypeToTableNamesMapping = Object.assign(
      defaultTableNames(),
      options.entityTypeToTableNamesMapping || {}
    );
    this._entityTypeToColumnValuesMapping = Object.assign(
      defaultTypeColumnNames(),
      options.entityTypeToColumnValuesMapping || {}
    );
  }

  async initialize(): Promise<void> {}

  supportsPointQueryForType(type: EntityMetadataType): boolean {
    return true;
  }

  async getMetadata(
    type: EntityMetadataType,
    id: string
  ): Promise<IEntityMetadata> {
    const tableName = this.getTableName(type);
    const row = await PostgresPoolQuerySingleRowAsync(
      this._pool,
      `
      SELECT *
      FROM ${tableName}
      WHERE
        entitytype = $1 AND
        entityid = $2
    `,
      [this.getEntityTypeColumnValue(type), id]
    );
    return this.rowToMetadataObject(type, row);
  }

  async setMetadata(serializedEntity: IEntityMetadata): Promise<void> {
    const entityType = serializedEntity.entityType;
    const tableName = this.getTableName(entityType);
    const internals = PostgresInternals.instance(entityType);
    const hasNatives = internals.hasNativeProperties();
    const { native, metadata } = internals.extractEntityToRowObjects(
      serializedEntity,
      true
    );
    const jsonValue = metadata; // formerly: this.metadataToRowMetadata(metadata);
    let nativeSqlInsert = '',
      nativeSqlNumbers = '',
      nativeSqlValues = [];
    const existingValues = 3;
    if (hasNatives) {
      const nativeProperties = Object.getOwnPropertyNames(native);
      if (nativeProperties.length > 0) {
        nativeSqlValues = nativeProperties.map(
          (columnName) => native[columnName]
        );
        nativeSqlInsert = nativeProperties
          .map((columnName) => `\n        ${columnName}`)
          .join(',');
        nativeSqlNumbers = nativeProperties
          .map((columnName, index) => '$' + String(index + existingValues + 1))
          .join(', ');
      }
    }
    try {
      const sql = `
      INSERT INTO ${tableName}(
        ${MapMetadataPropertiesToFields['entityType']},
        ${MapMetadataPropertiesToFields['entityId']},
        ${MetadataColumnName}${
        nativeSqlValues.length ? ',' : ''
      }${nativeSqlInsert}
      )
      VALUES (
        $1, $2, $3${nativeSqlValues.length ? ', ' : ''}${nativeSqlNumbers}
      )
    `;
      const values = [
        this.getEntityTypeColumnValue(entityType),
        serializedEntity.entityId,
        jsonValue,
        ...nativeSqlValues,
      ];
      const result = await PostgresPoolQueryAsync(this._pool, sql, values);
    } catch (insertError) {
      // insertError: message includes "duplicate key value violates"
      throw insertError;
    }
  }

  async updateMetadata(serializedEntity: IEntityMetadata): Promise<void> {
    const entityType = serializedEntity.entityType;
    const entityId = serializedEntity.entityId;
    const tableName = this.getTableName(entityType);
    const internals = PostgresInternals.instance(entityType);
    const hasNatives = internals.hasNativeProperties();
    const { nativeFieldNames, mapFieldsToColumnNames } = internals;
    const { native, metadata } = internals.extractEntityToRowObjects(
      serializedEntity,
      true
    );
    const jsonValue = metadata; // formerly: this.metadataToRowMetadata(metadata);
    let nativeSqlUpdates = '',
      nativeSqlValues = [];
    let updatedValuesCount = 1; // metadata is $1
    if (hasNatives) {
      nativeSqlUpdates = nativeFieldNames
        .map((nativeFieldName) => {
          const columnName = mapFieldsToColumnNames.get(nativeFieldName);
          return `\n           ${columnName} = $${++updatedValuesCount}`;
        })
        .join(',');
      nativeSqlValues = nativeFieldNames.map((nativeFieldName) => {
        const columnName = mapFieldsToColumnNames.get(nativeFieldName);
        return native[columnName] === undefined ? null : native[columnName];
      });
    }
    try {
      const sql = `
        UPDATE ${tableName}
        SET ${MetadataColumnName} = $1${
        nativeSqlValues.length ? ',' : ''
      }${nativeSqlUpdates}
        WHERE
          ${
            MapMetadataPropertiesToFields['entityType']
          } = $${++updatedValuesCount} AND
          ${
            MapMetadataPropertiesToFields['entityId']
          } = $${++updatedValuesCount}
      `;
      const values = [
        jsonValue,
        ...nativeSqlValues,
        this.getEntityTypeColumnValue(entityType),
        entityId,
      ];
      await PostgresPoolQueryAsync(this._pool, sql, values);
    } catch (updateError) {
      throw updateError;
    }
  }

  async deleteMetadata(metadata: IEntityMetadata): Promise<void> {
    const tableName = this.getTableName(metadata.entityType);
    await PostgresPoolQueryAsync(
      this._pool,
      `
      DELETE FROM ${tableName}
      WHERE
        ${MapMetadataPropertiesToFields['entityType']} = $1 AND
        ${MapMetadataPropertiesToFields['entityId']} = $2
    `,
      [this.getEntityTypeColumnValue(metadata.entityType), metadata.entityId]
    );
  }

  async clearMetadataStore(type: EntityMetadataType): Promise<void> {
    const tableName = this.getTableName(type);
    await PostgresPoolQueryAsync(
      this._pool,
      `
      DELETE FROM ${tableName}
      WHERE
        ${MapMetadataPropertiesToFields['entityType']} = $1
    `,
      [this.getEntityTypeColumnValue(type)]
    );
  }

  async fixedQueryMetadata(
    type: EntityMetadataType,
    query: IEntityMetadataFixedQuery
  ): Promise<IEntityMetadata[]> {
    const tableName = this.getTableName(type);
    const {
      sql,
      values,
      skipEntityMapping,
    } = this.createQueryFromFixedQueryEnum(tableName, type, query);
    return await this.sqlQueryToMetadataArray(
      type,
      sql,
      values,
      skipEntityMapping
    );
  }

  getSerializationHelper(
    type: EntityMetadataType
  ): IEntityMetadataSerializationHelper {
    const {
      mapFieldsToColumnNames: mapObjectToPostgresFields,
    } = PostgresInternals.instance(type);
    if (!mapObjectToPostgresFields) {
      return null;
    }
    const idFieldName = PostgresInternals.instance(type).getIdFieldName(true);
    return function objectToPostgresEntity(obj: any): IEntityMetadata {
      const metadata = SerializeObjectToEntityMetadata(
        type,
        idFieldName,
        obj,
        mapObjectToPostgresFields,
        true /* numbers to strings */,
        true /* throw if missing translations */,
        true
      );
      return metadata;
    };
  }

  getDeserializationHelper(
    type: EntityMetadataType
  ): IEntityMetadataDeserializationHelper {
    const {
      mapFieldsToColumnNames: mapObjectToPostgresFields,
    } = PostgresInternals.instance(type);
    if (!mapObjectToPostgresFields || mapObjectToPostgresFields.size === 0) {
      throw CreateError.ParameterRequired('mapFieldsToColumnNames');
    }
    if (!mapObjectToPostgresFields) {
      return null;
    }
    const idFieldName = PostgresInternals.instance(type).getIdFieldName(true);
    const dateColumnNames = EntityMetadataMappings.GetDefinition(
      type,
      PostgresSettings.PostgresDateColumns,
      false
    ) as string[];
    const dateColumns = new Set(dateColumnNames || []);
    return function postgresEntityToObject(entity: IEntityMetadata): any {
      const approval = EntityMetadataMappings.InstantiateObject(type);
      const toSet = DeserializeEntityMetadataToObjectSetCollection(
        entity,
        idFieldName,
        mapObjectToPostgresFields
      );
      for (const property in toSet) {
        approval[property] = toSet[property];
        if (
          dateColumns.has(property) &&
          approval[property] &&
          typeof approval[property] === 'string'
        ) {
          try {
            const dateParsed = new Date(approval[property]);
            approval[property] = dateParsed;
          } catch (ignored) {
            /* ignored */
          }
        }
      }
      return approval;
    };
  }

  private getEntityTypeColumnValue(type: EntityMetadataType): string {
    const value = this._entityTypeToColumnValuesMapping[type.typeName];
    if (!value) {
      throw new Error(
        `No Postgres column value mapping provider for EntityMetadataType value ${type}`
      );
    }
    return value;
  }

  private getTableName(type: EntityMetadataType): string {
    const tableName = this._entityTypeToTableNamesMapping[type.typeName];
    if (tableName) {
      return tableName;
    }
    if (!tableName) {
      throw new Error(
        `No Postgres table name mapping provided for EntityMetadataType value ${type}`
      );
    }
  }

  private metadataToRowMetadata(metadata: IEntityMetadata): any {
    const shallowClone = Object.assign({}, metadata);
    delete shallowClone.entityCreated;
    delete shallowClone.entityFieldNames;
    delete shallowClone.entityId;
    delete shallowClone.entityType;
    return shallowClone;
  }

  private stripEntityIdentities(type: EntityMetadataType, entity: any) {
    let entityTypeString = null;
    let entityCreated = null;
    let entityId = null;
    if (MapMetadataPropertiesToFields.entityType) {
      entityTypeString = entity[MapMetadataPropertiesToFields.entityType];
      delete entity[MapMetadataPropertiesToFields.entityType];
    }
    if (MapMetadataPropertiesToFields.entityId) {
      entityId = entity[MapMetadataPropertiesToFields.entityId];
      delete entity[MapMetadataPropertiesToFields.entityId];
    }
    if (MapMetadataPropertiesToFields.entityCreated) {
      entityCreated = entity[MapMetadataPropertiesToFields.entityCreated];
      delete entity[MapMetadataPropertiesToFields.entityCreated];
    }
    const internals = PostgresInternals.instance(type);
    const { metadata, native, leftovers } = internals.extractRowToFieldObjects(
      entity,
      MetadataColumnName,
      false
    );
    const combined = { ...metadata, ...native };
    const entityFieldNames = Object.getOwnPropertyNames(combined);
    return {
      entity: combined,
      entityTypeString,
      entityId,
      entityCreated,
      entityFieldNames,
      leftovers,
    };
  }

  private rowToMetadataObject(
    type: EntityMetadataType,
    row: any
  ): IEntityMetadata {
    const {
      entity,
      entityId,
      entityCreated,
      entityFieldNames,
    } = this.stripEntityIdentities(type, row);
    const entityIdentity: IEntityMetadata = {
      entityType: type,
      entityId,
      entityFieldNames,
      entityCreated,
    };
    const newMetadataObject: IEntityMetadata = Object.assign(
      entity,
      entityIdentity
    );
    return newMetadataObject;
  }

  private async sqlQueryToMetadataArray(
    type,
    sql,
    values,
    skipEntityMapping
  ): Promise<IEntityMetadata[]> {
    try {
      const result = await PostgresPoolQueryAsync(this._pool, sql, values);
      const rows = result['rows'];
      if (!rows) {
        throw new Error('No rows or empty rows returned');
      }
      return skipEntityMapping
        ? rows
        : rows.map((row) => this.rowToMetadataObject(type, row));
    } catch (errorish) {
      console.dir(errorish);
      throw errorish;
    }
  }

  private createQueryFromFixedQueryEnum(
    tableName: string,
    type: EntityMetadataType,
    query: IEntityMetadataFixedQuery
  ): any {
    let get = EntityMetadataMappings.GetDefinition(
      type,
      PostgresSettings.PostgresQueries,
      true
    ) as IPostgresGetQueries;
    const self = this;
    const getEntityTypeColumnValue = function (t) {
      return self.getEntityTypeColumnValue(t);
    };
    return get(
      query,
      MapMetadataPropertiesToFields,
      MetadataColumnName,
      tableName,
      getEntityTypeColumnValue
    );
  }
}

function defaultTableNames() {
  const defaults = {};
  EntityMetadataTypes.forEach((type) => {
    try {
      if (
        !PostgresInternals.instance(type).mapFieldsToColumnNames ||
        PostgresInternals.instance(type).mapFieldsToColumnNames.size === 0
      ) {
        return;
      }
      const tableName = PostgresInternals.instance(type).defaultTableName;
      if (!tableName) {
        throw CreateError.ParameterRequired('defaultTableName');
      }
      defaults[type.typeName] = tableName;
    } catch (noDefaultTableNameError) {
      throw new Error(
        `No default Postgres table name is defined for the type ${type}`
      );
    }
  });
  return defaults;
}

function defaultTypeColumnNames() {
  const defaults = {};
  EntityMetadataTypes.forEach((type) => {
    try {
      const map = PostgresInternals.instance(type).mapFieldsToColumnNames;
      if (!map || map.size === 0) {
        return;
      }
      const column = EntityMetadataMappings.GetDefinition(
        type,
        PostgresSettings.PostgresDefaultTypeColumnName,
        true
      );
      defaults[type.typeName] = column;
    } catch (noDefaultTableNameError) {
      throw new Error(
        `No default Postgres type column name is defined for the type ${type}`
      );
    }
  });
  return defaults;
}
