//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// OUT-OF-DATE: not as up-to-date as the postgres provider that is completely functional now

// This wrapper also implements table row encryption at rest.

import _ from 'lodash';
import azure from 'azure-storage';
import { v4 as uuidV4 } from 'uuid';

import { IReposError } from '../../../transitional';
import {
  ICorporateLinkProperties,
  ICorporateLink,
  ICorporateLinkExtended,
  CorporatePropertyNames,
} from '../../../business/corporateLink';
import { CorporateTableLink } from './tableLink';
import { ILinkProvider } from '..';

import tableEntity from '../../tableEntity';
import tableEncryption from '../../tableEncryption';

const defaultThirdPartyType = 'github';
const defaultPageSize = 500;
const defaultTableName = 'links';

const linkProviderInstantiationTypeProperty = '_i';
const dehydratedIdentityKey = '_lpi';
const dehydratedTableProviderName = 'xtable';
const dehydratedTableProviderVersion = '0';
const dehydratedTableProviderIdentitySeperator = '_';
const dehydratedTableProviderIdentity = `${dehydratedTableProviderName}${dehydratedTableProviderIdentitySeperator}${dehydratedTableProviderVersion}`;

enum LinkInstantiatedType {
  AzureTableEntity,
  Rehydrated,
}

interface IAlreadyLinkedError extends IReposError {
  alreadyLinked?: boolean;
}

interface IMultipleResultsError extends IReposError {
  multipleResults?: boolean;
}

const defaultEncryptedPropertyNames = [
  'githubToken',
  'githubTokenIncreasedScope',
  'localDataKey',
];

export interface ITableLinkProperties extends ICorporateLinkProperties {
  linkId: string;
  created: string;
}

const linkInterfacePropertyMapping: ITableLinkProperties = {
  linkId: 'linkid',

  isServiceAccount: 'serviceAccount',
  serviceAccountMail: 'serviceAccountMail',

  corporateId: 'aadoid',
  corporateUsername: 'aadupn',
  corporateDisplayName: 'aadname',
  corporateMailAddress: 'corporateMailAddress', // NOTE: this was not part of the original table entity
  corporateAlias: 'corporateAlias', // NOTE: this was not part of the original table entity

  thirdPartyId: 'ghid',
  thirdPartyUsername: 'ghu',
  thirdPartyAvatar: 'ghavatar',

  created: 'Timestamp',
};

const coreColumns = [
  'ghid',
  'ghu',
  'ghavatar',
  'aadoid',
  'aadupn',
  'aadname',
  'corporateMailAddress',
  'serviceAccount',
  'serviceAccountMail',
  'linkid',
];
// const coreColumnsList = coreColumns.join(', ');

export class TableLinkProvider implements ILinkProvider {
  private _tableName: string;
  private _tableNamePrefix: string;
  private _table: azure.TableService;
  private _providers: any;
  private _entityGenerator: any;
  private _options: any;
  private _thirdPartyType: string;

  public readonly propertyMapping: ITableLinkProperties = linkInterfacePropertyMapping;

  public readonly serializationIdentifierVersion: string = dehydratedTableProviderIdentity;

  constructor(providers, options) {
    if (!providers) {
      throw new Error(
        'The TableLinkProvider requires that available providers are passed into the constructor'
      );
    }

    options = options || {};

    const thirdPartyType = options.thirdPartyType || defaultThirdPartyType;
    if (thirdPartyType !== 'github') {
      throw new Error(
        'At this time only "github" is a supported third-party type.'
      );
    }

    const storageAccountName = options.account;
    if (!storageAccountName) {
      throw new Error(
        'Must provide options.account with an Azure Table storage account name'
      );
    }
    const storageAccountKey = options.key;
    if (!storageAccountKey) {
      throw new Error(
        'Must provide options.key with an Azure Table storage account key'
      );
    }

    this._providers = providers;
    this._options = options;
  }

  async initialize(): Promise<ILinkProvider> {
    const options = this._options || {};
    let table = azure.createTableService(options.account, options.key);
    this._table = table;
    this._entityGenerator = azure.TableUtilities.entityGenerator;
    this._tableNamePrefix = options.prefix || '';
    this._tableName =
      options.tableName || `${this._tableNamePrefix}${defaultTableName}`;
    if (options.encryption) {
      configureTableEncryption(this, options);
    }
    if (options.throwIfTableMissing) {
      let tableInfo = null;
      try {
        tableInfo = await tableDoesNotExist(table, this._tableName);
      } catch (tableError) {
        throw tableError;
      }
      if (!tableInfo.exists) {
        throw new Error(
          `The table named "${this._tableName}" does not exist. With options.throwIfTableMissing set, this error is thrown.`
        );
      }
    } else {
      await tableCreateIfNotExists(table, this._tableName);
    }
    return (this as any) as ILinkProvider;
  }

  get thirdPartyType() {
    return this._thirdPartyType;
  }

  getByThirdPartyUsername(username: string): Promise<CorporateTableLink> {
    username = username.toLowerCase();
    // NOTE: this is not normalized in the current data set!!!!!!!!
    // TODO: NOT NORMALIZED
    // TODO: VALUES in the current table have usernames that are MIXED CASE!!!!!!
    return this.getSingleLinkByProperty(
      this.propertyMapping.thirdPartyUsername,
      username
    ) as Promise<CorporateTableLink>;
  }

  async getByThirdPartyId(id: string): Promise<CorporateTableLink> {
    if (typeof id !== 'string') {
      id = (id as any).toString();
    }
    // Legacy table design: this call actually can go direct; in the
    // original implementation, the partition key is fixed and the
    // row key is the string value of the GitHub user ID.
    // SLOW query equivalent: return getUserEntityByProperty(this, 'ghid', id, callback);
    const partitionKey = this._options.partitionKey;
    if (!partitionKey) {
      throw new Error(
        'No table options.partitionKey provided with a fixed partition key at this time'
      );
    }
    const tableName = this._tableName;
    const fullEntity = await tableRetrieveEntity(
      this._table,
      tableName,
      partitionKey,
      id
    );
    if (fullEntity === false) {
      return (false as any) as CorporateTableLink;
    }
    const row = tableEntity.reduce(fullEntity);
    const link = createLinkInstanceFromAzureTableEntity(this, row);
    return link;
  }

  queryByCorporateId(id: string): Promise<CorporateTableLink[]> {
    return this.getLinksByProperty('aadoid', id);
  }

  async getAll(): Promise<CorporateTableLink[]> {
    const queryOptions = {
      columns: [
        'aadoid',
        'aadupn',
        'aadname',
        'ghu',
        'ghid',
        'ghavatar',
        'serviceAccount',
        'serviceAccountMail',
        'PartitionKey',
        'RowKey',
        'Timestamp',
      ],
    };
    const unsorted = await queryLinksTable(this, queryOptions);
    const sorted = _.sortBy(unsorted, ['aadupn', 'ghu']);
    const links = createLinkInstancesFromAzureTableEntityArray(this, sorted);
    return links;
  }

  async getAllCorporateIds(): Promise<string[]> {
    const queryOptions = {
      columns: ['aadoid', 'PartitionKey', 'RowKey', 'Timestamp'],
    };
    const results = await queryLinksTable(this, queryOptions);
    return results.map((row) => String(row.aadoid)) as string[];
  }

  queryByCorporateUsername(username: string): Promise<CorporateTableLink[]> {
    // ?? username = username.toLowerCase();
    // TODO: not sure if this one is normalized or not...
    return this.getLinksByProperty('aadupn', username);
  }

  async createLink(link: ICorporateLink): Promise<string> {
    const generatedLinkId = uuidV4();
    const tableName = this._tableName;
    let entity = null;
    try {
      const initialEntity = {};
      initialEntity[linkInterfacePropertyMapping.linkId] = generatedLinkId;
      for (let linkPropertyName of CorporatePropertyNames) {
        // linkInterfacePropertyMapping
        const tableColumnName = linkInterfacePropertyMapping[linkPropertyName];
        if (!tableColumnName) {
          throw new Error(
            `Missing mapping from property ${linkPropertyName} to equivalent table column`
          );
        }
        initialEntity[tableColumnName] = link[linkPropertyName];
      }
      const partitionKey = this._options.partitionKey;
      if (!partitionKey) {
        throw new Error(
          'No table options.partitionKey provided with a fixed partition key at this time'
        );
      }
      entity = tableEntity.create(
        partitionKey,
        link.thirdPartyId,
        initialEntity
      );
    } catch (processingError) {
      throw processingError;
    }
    await tableInsertEntity(
      this._table,
      tableName,
      entity,
      'This user is already linked'
    );
    return generatedLinkId;
  }

  updateLink(linkInstance: ICorporateLink): Promise<any> {
    const tl = linkInstance as CorporateTableLink;
    const replacementEntity = tl.internal().getDirectEntity();
    if (linkInstance.thirdPartyId) {
      return this.updateLinkByThirdPartyIdLegacy(
        linkInstance.thirdPartyId,
        replacementEntity
      );
    }
    throw new Error(
      'updateLink is not yet updated for linkId without a given thirdPartyId (ghid)'
    );
  }

  async deleteLink(linkInstance: ICorporateLink): Promise<any> {
    // This is inefficient at this time; with the newer design centering
    // around a link ID, this has to query first.
    const tl = linkInstance as CorporateTableLink;
    const linkId = tl.id;
    if (!linkId && linkInstance.thirdPartyId) {
      return this.deleteLinkByThirdPartyIdLegacy(linkInstance.thirdPartyId);
    }
    const link = (this.getSingleLinkByProperty(
      this.propertyMapping.linkId,
      linkId
    ) as any) as CorporateTableLink;
    if (!link) {
      throw new Error(`No link found with ID ${linkId}`);
    }
    if (!link.thirdPartyId) {
      throw new Error(`Link ${linkId} is missing a valid thirdPartyId`);
    }
    return this.deleteLinkByThirdPartyIdLegacy(link.thirdPartyId);
  }

  async updateLinkByThirdPartyIdLegacy(
    thirdPartyId: string,
    replaceEntity: any
  ): Promise<void> {
    const partitionKey = this._options.partitionKey;
    if (!partitionKey) {
      throw new Error(
        'No table options.partitionKey provided with a fixed partition key at this time'
      );
    }
    if (typeof thirdPartyId !== 'string') {
      thirdPartyId = (thirdPartyId as any).toString();
    }
    const tableName = this._tableName;
    if (!replaceEntity.linkId) {
      console.log('Generated a new linkId as part of an update operation');
      const newLinkId = uuidV4();
      replaceEntity.linkId = newLinkId;
    }
    const entity = tableEntity.create(
      partitionKey,
      thirdPartyId,
      replaceEntity
    );
    await tableReplaceEntity(this._table, tableName, entity);
  }

  deleteLinkByThirdPartyIdLegacy(thirdPartyId: string): Promise<void> {
    const partitionKey = this._options.partitionKey;
    if (!partitionKey) {
      throw new Error(
        'No table options.partitionKey provided with a fixed partition key at this time'
      );
    }
    if (typeof thirdPartyId !== 'string') {
      thirdPartyId = (thirdPartyId as any).toString();
    }
    const tableName = this._tableName;
    return tableDeleteEntity(
      this._table,
      tableName,
      partitionKey,
      thirdPartyId
    );
  }

  dehydrateLink(linkInstance: ICorporateLinkExtended): any {
    // CONSIDER: check whether the current link type feels appropriate to us (PGSQL)
    const tlink = linkInstance as CorporateTableLink;
    const entity = tlink.internal().getDirectEntity();
    const shriveled = Object.assign({}, entity);
    shriveled[dehydratedIdentityKey] = dehydratedTableProviderIdentity;
    return shriveled;
  }

  rehydrateLink(jsonObject: any): ICorporateLink {
    if (!jsonObject) {
      throw new Error('No object provided to rehydrate');
    }
    const identity = jsonObject[dehydratedIdentityKey] as string;
    if (!identity) {
      throw new Error('No stored link provider identity to validate');
    }
    if (identity !== dehydratedTableProviderIdentity) {
      const sameProviderType = identity.startsWith(
        `${dehydratedTableProviderName}${dehydratedTableProviderIdentitySeperator}`
      );
      if (sameProviderType) {
        // Cross-version rehydration not supported
        throw new Error(
          `The hydrated link was created by the same ${dehydratedTableProviderName} provider, but a different version: ${identity}`
        );
      } else {
        throw new Error(
          `The hydrated link is incompatible with this runtime environment: ${identity}`
        );
      }
    }
    const clonedObject = Object.assign({}, jsonObject);
    delete clonedObject[dehydratedIdentityKey];
    const pglink = this.createLinkInstanceFromHydratedEntity(clonedObject);
    return pglink;
  }

  dehydrateLinks(linkInstances: ICorporateLink[]): any[] {
    if (!Array.isArray(linkInstances)) {
      throw new Error('linkInstances must be an array');
    }
    if (linkInstances.length > 0) {
      const first = linkInstances[0];
      if (first[linkProviderInstantiationTypeProperty] === undefined) {
        throw new Error(
          'linkInstances[0] does not appear to be a link instantiated by a provider'
        );
      }
    }
    //
    const arr: any[] = linkInstances.map(this.dehydrateLink.bind(this));
    return arr;
  }

  rehydrateLinks(jsonArray: any): ICorporateLink[] {
    if (!Array.isArray(jsonArray)) {
      throw new Error('jsonArray must be an array');
    }
    //
    const arr = jsonArray.map(this.rehydrateLink.bind(this));
    return (arr as any[]) as ICorporateLink[];
  }

  private createLinkInstanceFromHydratedEntity(jsonObject) {
    const linkInternalOptions = {
      provider: this,
    };
    const newLink = new CorporateTableLink(linkInternalOptions, jsonObject);
    newLink[linkProviderInstantiationTypeProperty] =
      LinkInstantiatedType.Rehydrated; // in case this helps while debugging
    return newLink;
  }

  private async getTableEntitiesByProperty(
    propertyName: string,
    value
  ): Promise<any[]> {
    const queryOptions = {
      wherePropertyName: propertyName,
      whereValue: value,
    };
    return queryLinksTable(this, queryOptions);
  }

  private async getLinksByProperty(
    propertyName,
    value
  ): Promise<CorporateTableLink[]> {
    const rows = await this.getTableEntitiesByProperty(propertyName, value);
    const links = createLinkInstancesFromAzureTableEntityArray(this, rows);
    return links;
  }

  private async getSingleLinkByProperty(
    propertyName: string,
    value
  ): Promise<CorporateTableLink | boolean> {
    const rows = await this.getTableEntitiesByProperty(propertyName, value);
    if (rows.length <= 0) {
      return false;
    }
    if (rows.length > 1) {
      const error: IMultipleResultsError = new Error(
        `More than a single result were returned by the query (${rows.length})`
      );
      error.multipleResults = rows.length > 0;
      throw error;
    }
    const entityRow = rows[0];
    const link = createLinkInstanceFromAzureTableEntity(this, entityRow);
    return link;
  }
}

function createLinkInstancesFromAzureTableEntityArray(
  provider: TableLinkProvider,
  rows: any[]
): CorporateTableLink[] {
  return rows.map(createLinkInstanceFromAzureTableEntity.bind(null, provider));
}

function createLinkInstanceFromAzureTableEntity(
  provider: TableLinkProvider,
  row: any
): CorporateTableLink {
  const linkInternalOptions = {
    provider,
  };
  const newLink = new CorporateTableLink(linkInternalOptions, row);
  newLink[linkProviderInstantiationTypeProperty] =
    LinkInstantiatedType.AzureTableEntity; // in case this helps while debugging
  return newLink;
}

function getTableName(self) {
  // setup during init
  return self._tableName;
}

function getTableService(self): azure.TableService {
  // setup during init
  return self._table;
}

async function queryLinksTable(
  self: TableLinkProvider,
  options
): Promise<any[]> {
  const rows = [];
  function buildQuery(options) {
    let q = new azure.TableQuery();
    if (options.columns) {
      q = q.select(options.columns);
    }
    if (options.wherePropertyName && options.whereValue) {
      q = q.where(`${options.wherePropertyName} eq ?`, options.whereValue);
    }
    const pageSize = options.pageSize || defaultPageSize;
    q = q.top(pageSize);
    return q;
  }
  const tableName = options.tableName || getTableName(self);
  let continuationToken = null;
  let done = false;
  function pushQueryEntities(
    tableService: azure.TableService,
    options: any
  ): Promise<void> {
    return new Promise((resolve, reject) => {
      tableService.queryEntities(
        tableName,
        buildQuery(options),
        continuationToken,
        (queryError, results: azure.TableService.QueryEntitiesResult<any>) => {
          if (queryError) {
            done = true;
            return reject(queryError);
          }
          // Is there another page of query results?
          if (results.continuationToken) {
            continuationToken = results.continuationToken;
          } else {
            done = true;
          }
          // NOTE: This function does not use continuation tokens or anything like that!
          const entries = results && results.entries ? results.entries : null;
          for (let i = 0; entries && i < entries.length; i++) {
            rows.push(tableEntity.reduce(entries[i]));
          }
          return resolve();
        }
      );
    });
  }
  const tableService = getTableService(self);
  while (!done) {
    await pushQueryEntities(tableService, buildQuery(options));
  }
  return rows;
}

function configureTableEncryption(self, options) {
  const encryptionOptions = options.encryption;

  const encryptionKeyId = encryptionOptions.encryptionKeyId;
  if (!encryptionKeyId) {
    throw new Error('Encryption requires options.encryptionKeyId');
  }

  const keyResolver = encryptionOptions.keyEncryptionKeyResolver;
  if (!keyResolver) {
    throw new Error('Encryption requires options.keyResolver');
  }

  const encryptedPropertyNames = new Set(
    encryptionOptions.encryptedPropertyNames || defaultEncryptedPropertyNames
  );

  const opts = {
    keyEncryptionKeyId: encryptionKeyId,
    keyResolver: keyResolver,
    encryptedPropertyNames: encryptedPropertyNames,
    binaryProperties: 'buffer',
    tableDehydrator: encryptionOptions.tableDehydrator || tableEntity.reduce,
    tableRehydrator: encryptionOptions.tableRehydrator || tableEntity.create,
  };

  // re-wrap the table client with the encrypting variant
  const standardTableClient = self._table;
  self._table = tableEncryption(standardTableClient, opts);
}

function tableDoesNotExist(
  tableService: azure.TableService,
  tableName: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    tableService.doesTableExist(tableName, (tableError, tableInfo) => {
      if (!tableError && !tableInfo.exists) {
        tableError = new Error(
          `The table named "${this._tableName}" does not exist. With options.throwIfTableMissing set, this error is thrown.`
        );
      }
      if (tableError) {
        return reject(tableError);
      }
      return resolve();
    });
  });
}

function tableCreateIfNotExists(
  tableService: azure.TableService,
  tableName: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    tableService.createTableIfNotExists(this._tableName, (error) => {
      return error ? reject(error) : resolve();
    });
  });
}

function tableRetrieveEntity(
  tableService: azure.TableService,
  tableName: string,
  partitionKey: string,
  rowKey: string
): Promise<any> {
  return new Promise((resolve, reject) => {
    tableService.retrieveEntity(
      tableName,
      partitionKey,
      rowKey,
      (getError: any, fullEntity) => {
        if (getError && getError.statusCode === 404) {
          return resolve(false);
        }
        if (getError) {
          return reject(getError);
        }
        return resolve(fullEntity);
      }
    );
  });
}

function tableInsertEntity(
  tableService: azure.TableService,
  tableName: string,
  entity: any,
  entityAlreadyExistsErrorMessage: string
): Promise<any> {
  return new Promise((resolve, reject) => {
    return tableService.insertEntity(
      tableName,
      entity,
      (insertError: any, inserted) => {
        if (insertError && insertError.code === 'EntityAlreadyExists') {
          const error: IAlreadyLinkedError = new Error(
            entityAlreadyExistsErrorMessage
          );
          error.alreadyLinked = true;
          error.innerError = insertError;
          return reject(error);
        }
        return insertError ? reject(insertError) : resolve(inserted);
      }
    );
  });
}

function tableReplaceEntity(
  tableService: azure.TableService,
  tableName: string,
  entity: any
): Promise<any> {
  return new Promise((resolve, reject) => {
    return tableService.replaceEntity(tableName, entity, (error, ok) => {
      return error ? reject(error) : resolve(ok);
    });
  });
}

function tableDeleteEntity(
  tableService: azure.TableService,
  tableName: string,
  partitionKey: string,
  rowKey: string
): Promise<any> {
  return new Promise((resolve, reject) => {
    return this._table.deleteEntity(
      tableName,
      tableEntity.create(partitionKey, rowKey),
      (error, ok) => {
        return error ? reject(error) : resolve(ok);
      }
    );
  });
}
