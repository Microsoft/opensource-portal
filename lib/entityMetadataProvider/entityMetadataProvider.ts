//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IEntityMetadata, EntityMetadataType } from './entityMetadata';
import { IEntityMetadataFixedQuery } from './query';
import { swapMap } from '../../utils';

export enum EntityField {
  Type = 'entityType',
  ID = 'entityId',
  FieldNames = 'entityFieldNames',
  Created = 'entityCreated',
}

export const EntityFieldNames = new Set<string>([
  EntityField.Type as string,
  EntityField.ID as string,
  EntityField.FieldNames as string,
  EntityField.Created as string,
]);

export interface IEntityMetadataSerializationHelper {
  // serialize an object to defined metadata
  (obj: any): IEntityMetadata;
}

export interface IEntityMetadataDeserializationHelper {
  // deserialize defined metadata to an object
  (entity: IEntityMetadata): any;
}

export interface IEntityMetadataProvider {
  initialize(): Promise<void>;

  getMetadata(type: EntityMetadataType, id: string): Promise<IEntityMetadata>;
  setMetadata(metadata: IEntityMetadata): Promise<void>;
  updateMetadata(metadata: IEntityMetadata): Promise<void>;
  deleteMetadata(metadata: IEntityMetadata): Promise<void>;
  clearMetadataStore(type: EntityMetadataType): Promise<void>;

  supportsHistory: boolean;
  name: string;

  fixedQueryMetadata(
    type: EntityMetadataType,
    query: IEntityMetadataFixedQuery
  ): Promise<IEntityMetadata[]>;

  supportsPointQueryForType(type: EntityMetadataType): boolean;

  getSerializationHelper(
    type: EntityMetadataType
  ): IEntityMetadataSerializationHelper;
  getDeserializationHelper(
    type: EntityMetadataType
  ): IEntityMetadataDeserializationHelper;
}

export interface IObjectWithDefinedKeys {
  getObjectFieldNames(): string[];
}

export function SerializeObjectToEntityMetadata(
  type: EntityMetadataType,
  idFieldName: string,
  obj: IObjectWithDefinedKeys,
  translationMap: Map<string, string>,
  castNumbersToStrings: boolean,
  throwIfMissingTranslations: boolean,
  ignorePrivateMembers: boolean
): IEntityMetadata {
  const id = obj[idFieldName];
  if (!id) {
    throw new Error(
      `No identity for entity object to serialize found in key: ${idFieldName}`
    );
  }
  const em: IEntityMetadata = {
    entityType: type,
    entityId: id,
    entityFieldNames: [],
  };
  let allKeys = new Set(Array.from(Object.getOwnPropertyNames(obj)));
  allKeys.delete(idFieldName);
  const objectKeys = obj.getObjectFieldNames
    ? obj.getObjectFieldNames()
    : Object.getOwnPropertyNames(obj);
  const setKeys = [];
  const missingKeys = [];
  for (let i = 0; i < objectKeys.length; i++) {
    const key = objectKeys[i];
    let value = obj[key];
    if (castNumbersToStrings && typeof value === 'number') {
      value = value.toString();
    }
    allKeys.delete(key);
    const translatesTo = translationMap.get(key);
    if (!translatesTo && key === idFieldName) {
      // since the entity id is this key, a mapping is not required; ignore.
    } else if (!translatesTo && key !== idFieldName) {
      missingKeys.push(key);
    } else {
      if (value !== undefined && value !== null) {
        em[translatesTo] = value;
        if (!EntityFieldNames.has(translatesTo)) {
          // do not map internal properties
          setKeys.push(translatesTo);
        }
      }
    }
  }
  if (ignorePrivateMembers) {
    Array.from(allKeys.keys()).map((key: string) => {
      if (key.startsWith('_')) {
        allKeys.delete(key);
      }
    });
  }
  if (allKeys.size > 0) {
    const asList = Array.from(allKeys.keys()).join(', ');
    const message = `Additional ${type} entity properties were present on the object: ${asList}`;
    if (throwIfMissingTranslations) {
      throw new Error(`${message} [throwIfMissingTranslations=true]`);
    } else {
      console.warn(message);
    }
  }
  em.entityFieldNames.push(...setKeys);
  if (missingKeys.length && throwIfMissingTranslations) {
    throw new Error(
      `Missing translation entries for keys: ${missingKeys.join(', ')}`
    );
  }
  return em;
}

export function DeserializeEntityMetadataToObjectSetCollection(
  entity: IEntityMetadata,
  destinationIdFieldName: string,
  serializationTranslationMap: Map<
    string,
    string
  > /*, allowOverridingIdFieldname: boolean*/
): any {
  const setCollection = {};
  setCollection[destinationIdFieldName] = entity.entityId;
  const reverseMap = swapMap(serializationTranslationMap);
  for (const [key, objectKey] of reverseMap.entries()) {
    const value = entity[key];
    if (value !== undefined && value !== null) {
      setCollection[objectKey] = value;
    }
  }
  if (!setCollection[destinationIdFieldName]) {
    throw new Error(
      `The destination field ${destinationIdFieldName} was overwritten deserializing the metadata for ${entity.entityType}`
    );
  }
  return setCollection;
}
