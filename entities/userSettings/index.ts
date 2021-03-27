//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  EntityMetadataType,
  IEntityMetadataBaseOptions,
  EntityMetadataBase,
  IEntityMetadata,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  QueryBase,
  IEntityMetadataFixedQuery,
} from '../../lib/entityMetadataProvider/query';
import {
  EntityMetadataMappings,
  MetadataMappingDefinition,
} from '../../lib/entityMetadataProvider/declarations';
import {
  PostgresConfiguration,
  PostgresSettings,
} from '../../lib/entityMetadataProvider/postgres';

const type = new EntityMetadataType('UserSettings');
const thisProviderType = type;

export interface IUserSettingsProperties {
  // PRIMARY: corporateId: any;
  contributionShareOptIn: any;
}

const corporateId = 'corporateId';

class UserSettingsQueryBase extends QueryBase<UserSettings> {
  constructor(public query: Query) {
    super();
  }
}

class UserSettingsQuery<T> extends UserSettingsQueryBase {
  constructor(query: Query, public parameters: T) {
    super(query);
    if (!this.parameters) {
      this.parameters = {} as T;
    }
  }
}

enum Query {
  UsersOptedInToShareData = 'UsersOptedInToShareData',
}

interface NoParameters {}

const Field: IUserSettingsProperties = {
  // corporateId: 'corporateId',
  contributionShareOptIn: 'contributionShareOptIn',
};

const fieldNames = Object.getOwnPropertyNames(Field);
const nativeFieldNames = fieldNames;

export class UserSettings implements IUserSettingsProperties {
  corporateId: string;

  contributionShareOptIn: boolean;

  constructor() {}
}

EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityInstantiate,
  () => {
    return new UserSettings();
  }
);
EntityMetadataMappings.Register(
  type,
  MetadataMappingDefinition.EntityIdColumnName,
  corporateId
);

PostgresConfiguration.SetDefaultTableName(type, 'usersettings');
EntityMetadataMappings.Register(
  type,
  PostgresSettings.PostgresDefaultTypeColumnName,
  'usersettings'
);
PostgresConfiguration.MapFieldsToColumnNamesFromListLowercased(
  type,
  fieldNames
);
PostgresConfiguration.IdentifyNativeFields(type, nativeFieldNames);
PostgresConfiguration.ValidateMappings(type, fieldNames, [corporateId]);

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
    // const entityTypeColumn = mapMetadataPropertiesToFields[EntityField.Type];
    // const entityTypeValue = getEntityTypeColumnValue(type);
    const base = query as UserSettingsQueryBase;
    switch (base.query) {
      case Query.UsersOptedInToShareData: {
        return {
          sql: `
          SELECT
            *
          FROM
            ${tableName}
          WHERE
            contributionshareoptin = $1
          `,
          values: [true],
        };
      }
      default:
        throw new Error(
          `The query ${base.query} is not implemented by this provider for the type ${type}`
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

export interface IUserSettingsProviderCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IUserSettingsProvider {
  initialize(): Promise<void>;
  getUserSettings(corporateId: string): Promise<UserSettings>;
  insertUserSettings(record: UserSettings): Promise<string>;
  updateUserSettings(record: UserSettings): Promise<void>;

  queryContributionOptInUsers(): Promise<UserSettings[]>;
}

export class UserSettingsProvider
  extends EntityMetadataBase
  implements IUserSettingsProvider {
  constructor(options: IUserSettingsProviderCreateOptions) {
    super(options);
  }

  async initialize() {}

  async updateUserSettings(record: UserSettings): Promise<void> {
    const entity = this.serialize(thisProviderType, record);
    return await this._entities.updateMetadata(entity);
  }

  async getUserSettings(eventId: string): Promise<UserSettings> {
    this.ensureHelpers(thisProviderType);
    let metadata: IEntityMetadata = null;
    metadata = await this._entities.getMetadata(thisProviderType, eventId);
    return this.deserialize<UserSettings>(thisProviderType, metadata);
  }

  async insertUserSettings(record: UserSettings): Promise<string> {
    const entity = this.serialize(thisProviderType, record);
    await this._entities.setMetadata(entity);
    return entity.entityId;
  }

  async queryContributionOptInUsers(): Promise<UserSettings[]> {
    const query = new UserSettingsQuery<NoParameters>(
      Query.UsersOptedInToShareData,
      {}
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<UserSettings>(
      thisProviderType,
      metadatas
    );
    return results;
  }
}
