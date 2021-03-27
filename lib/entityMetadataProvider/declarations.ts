//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { EntityMetadataType } from './entityMetadata';

export abstract class MetadataMappingDefinitionBase {
  alternateRuntimeValidateMapping: Map<string, string>;

  constructor(public definitionName: string) {}

  toString(): string {
    return this.definitionName;
  }
}

class LegacyMappingDefinition extends MetadataMappingDefinitionBase {
  constructor(definitionName: string) {
    super(definitionName);
  }
}

export const MetadataMappingDefinition = {
  EntityIdColumnName: new LegacyMappingDefinition('EntityIdColumnName'),
  EntityInstantiate: new LegacyMappingDefinition('EntityInstantiate'),
};

export class EntityMetadataMappings {
  private static _values = new Map<
    EntityMetadataType,
    Map<MetadataMappingDefinitionBase, any>
  >();

  public static Register(
    type: EntityMetadataType,
    definitionType: MetadataMappingDefinitionBase,
    definition: any
  ) {
    if (!EntityMetadataMappings._values.has(type)) {
      EntityMetadataMappings._values.set(type, new Map());
    }
    const typeMap = EntityMetadataMappings._values.get(type);
    if (typeMap.has(definitionType)) {
      throw new Error(`Entity type ${type} already registered`);
    }
    typeMap.set(definitionType, definition);
  }

  public static GetDefinition(
    type: EntityMetadataType,
    definitionType: MetadataMappingDefinitionBase,
    throwIfMissing: boolean
  ): any {
    if (!EntityMetadataMappings._values.has(type)) {
      throw new Error(
        `Type definitions not initialized or set to ${type} (${definitionType})`
      );
    }
    const typeMap = EntityMetadataMappings._values.get(type);
    const d = typeMap.get(definitionType);
    if (d) {
      return d;
    }
    if (throwIfMissing) {
      throw new Error(
        `Entity type definitions (${definitionType}) are not available for ${type} in the configured entity metadata provider`
      );
    }
  }

  public static InstantiateObject(type: EntityMetadataType) {
    const ctor = EntityMetadataMappings.GetDefinition(
      type,
      MetadataMappingDefinition.EntityInstantiate,
      true
    );
    return ctor();
  }

  public static RuntimeValidateMappings(
    type: EntityMetadataType,
    definitionType: MetadataMappingDefinitionBase,
    fieldNames: string[],
    permittedAdditionalUnvisitedMappings: string[]
  ) {
    try {
      const mapping =
        definitionType.alternateRuntimeValidateMapping ||
        (EntityMetadataMappings.GetDefinition(
          type,
          definitionType,
          true
        ) as Map<string, string>);
      if (!mapping || !mapping.keys) {
        throw new Error(
          `RuntimeValidateMappings: type ${type} definition ${definitionType} does not have a map`
        );
      }
      const unvisitedMappings = new Set(mapping.keys());
      const fields = new Set(fieldNames);
      for (let i = 0; i < fieldNames.length; i++) {
        const fn = fieldNames[i];
        if (!mapping.has(fn)) {
          throw new Error(
            `RuntimeValidateMappings: type ${type} definition ${definitionType} does not have a defined mapping for the column named: ${fn}`
          );
        }
        unvisitedMappings.delete(fn);
        fields.delete(fn);
      }
      if (fields.size) {
        const list = Array.from(fields.keys());
        throw new Error(
          `RuntimeValidateMappings: type ${type} definition ${definitionType} has no mapping for fields: ${list.join(
            ', '
          )}`
        );
      }
      if (
        permittedAdditionalUnvisitedMappings &&
        permittedAdditionalUnvisitedMappings.length
      ) {
        permittedAdditionalUnvisitedMappings.map((each) => {
          unvisitedMappings.delete(each);
        });
      }
      if (unvisitedMappings.size) {
        const list = Array.from(unvisitedMappings.keys());
        throw new Error(
          `RuntimeValidateMappings: type ${type} definition ${definitionType} has unvisited mappings: ${list.join(
            ', '
          )}`
        );
      }
    } catch (error) {
      console.log('RuntimeValidateMappings error:');
      console.log(error);
      throw error;
    }
  }
}
