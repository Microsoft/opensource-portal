//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Loose in-memory implementation. Note that it does not translate or clone
// objects, rather storing them as-is.

import {
  IEntityMetadataProvider,
  IEntityMetadataSerializationHelper,
  IEntityMetadataDeserializationHelper,
  SerializeObjectToEntityMetadata,
  DeserializeEntityMetadataToObjectSetCollection,
} from './entityMetadataProvider';
import { IEntityMetadata, EntityMetadataType } from './entityMetadata';
import { IEntityMetadataFixedQuery } from './query';
import {
  EntityMetadataMappings,
  MetadataMappingDefinition,
  MetadataMappingDefinitionBase,
} from './declarations';

class MemoryMetadataDefinition extends MetadataMappingDefinitionBase {
  constructor(name: string) {
    super(name);
  }
}

export const MemorySettings = {
  MemoryMapping: new MemoryMetadataDefinition('MemoryMapping'),
  MemoryQueries: new MemoryMetadataDefinition('MemoryQueries'),
};

interface IMemoryGetQueries {
  (query: IEntityMetadataFixedQuery, directMemory: IEntityMetadata[]): any;
}

export class MemoryEntityMetadataProvider implements IEntityMetadataProvider {
  private _entitiesByType: Map<
    EntityMetadataType,
    Map<string, IEntityMetadata[]>
  >;
  public readonly name = 'memory';
  public readonly supportsHistory: boolean = true;

  constructor() {}

  supportsPointQueryForType(type: EntityMetadataType): boolean {
    return true;
  }

  async initialize(): Promise<void> {
    this._entitiesByType = new Map();
  }

  private entityBin(type: EntityMetadataType): Map<string, IEntityMetadata[]> {
    let map: Map<string, IEntityMetadata[]> = this._entitiesByType.get(type);
    if (!map) {
      map = new Map<string, IEntityMetadata[]>();
      this._entitiesByType.set(type, map);
    }
    return map;
  }

  async getMetadata(
    type: EntityMetadataType,
    id: string
  ): Promise<IEntityMetadata> {
    const bin = this.entityBin(type).get(id);
    return bin ? bin[bin.length - 1] : null;
  }

  async setMetadata(metadata: IEntityMetadata): Promise<void> {
    // CONSIDER: should throw if the ID already exists to mimic other common data stores
    const type = metadata.entityType;
    const id = metadata.entityId;
    let bin = this.entityBin(type).get(id);
    if (!bin) {
      bin = [];
      this.entityBin(type).set(id, bin);
    }
    bin.push(metadata);
  }

  async updateMetadata(metadata: IEntityMetadata): Promise<void> {
    return await this.setMetadata(metadata);
  }

  async deleteMetadata(metadata: IEntityMetadata): Promise<void> {
    const type = metadata.entityType;
    const id = metadata.entityId;
    this.entityBin(type).delete(id);
  }

  async clearMetadataStore(type: EntityMetadataType): Promise<void> {
    this.entityBin(type).clear();
  }

  async fixedQueryMetadata(
    type: EntityMetadataType,
    query: IEntityMetadataFixedQuery
  ): Promise<IEntityMetadata[]> {
    const allInTypeBin = this.getAllInTypeBin(type);
    let get = EntityMetadataMappings.GetDefinition(
      type,
      MemorySettings.MemoryQueries,
      true
    ) as IMemoryGetQueries;
    return get(query, allInTypeBin);
  }

  private getAllInTypeBin(type: EntityMetadataType): IEntityMetadata[] {
    const allValuesInBin = Array.from(this.entityBin(type).values());
    return allValuesInBin.map((eachInnerBin) => {
      return eachInnerBin[eachInnerBin.length - 1];
    });
  }

  getSerializationHelper(
    type: EntityMetadataType
  ): IEntityMetadataSerializationHelper {
    const mapObjectToMemoryFields = EntityMetadataMappings.GetDefinition(
      type,
      MemorySettings.MemoryMapping,
      true
    );
    if (!mapObjectToMemoryFields) {
      return null;
    }
    const idFieldName = EntityMetadataMappings.GetDefinition(
      type,
      MetadataMappingDefinition.EntityIdColumnName,
      true
    );
    return function objectToMemoryEntity(obj: any): IEntityMetadata {
      const metadata = SerializeObjectToEntityMetadata(
        type,
        idFieldName,
        obj,
        mapObjectToMemoryFields,
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
    const mapObjectToMemoryFields = EntityMetadataMappings.GetDefinition(
      type,
      MemorySettings.MemoryMapping,
      true
    );
    if (!mapObjectToMemoryFields) {
      return null;
    }
    const idFieldName = EntityMetadataMappings.GetDefinition(
      type,
      MetadataMappingDefinition.EntityIdColumnName,
      true
    );
    return function memoryEntityToObject(entity: IEntityMetadata): any {
      const approval = EntityMetadataMappings.InstantiateObject(type);
      const toSet = DeserializeEntityMetadataToObjectSetCollection(
        entity,
        idFieldName,
        mapObjectToMemoryFields
      );
      for (const property in toSet) {
        approval[property] = toSet[property];
      }
      return approval;
    };
  }
}
