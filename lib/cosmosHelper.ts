//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { CosmosClient, Database, Container } from '@azure/cosmos';

const debug = require('debug')('cosmosdb');

export interface ICosmosHelperOptions {
  endpoint: string;
  key: string;
  database?: string;
  collection?: string;
}

export default class CosmosHelper {
  // Simplified subset
  private _options: ICosmosHelperOptions;
  private _client: CosmosClient;
  private _initialized: boolean;
  private _database: Database;
  private _collection: Container;

  constructor(options: ICosmosHelperOptions) {
    this._options = options;
  }

  async initialize() {
    if (this._initialized) {
      return;
    }
    const { endpoint, key } = this._options;
    if (!endpoint) {
      throw new Error('options.endpoint required');
    }
    if (!key) {
      throw new Error('options.key required');
    }
    if (!this._options.collection) {
      throw new Error('options.collection required');
    }
    if (!this._options.database) {
      throw new Error('options.database required');
    }
    this._client = new CosmosClient({ endpoint, key });
    this._database = (
      await this._client.databases.createIfNotExists({
        id: this._options.database,
      })
    ).database;
    this._collection = (
      await this._database.containers.createIfNotExists({
        id: this._options.collection,
      })
    ).container;
    this._initialized = true;
  }

  async getObject(partitionKey: string, documentId: string): Promise<any> {
    this.throwIfNotInitialized();
    debug(`COSMOS GET OBJECT: ${documentId} from PARTITION: ${partitionKey}`);
    const response = await this._collection
      .item(documentId, partitionKey)
      .read();
    if (response.resource) {
      const clone = Object.assign({}, response.resource);
      delete clone._attachments;
      delete clone._etag;
      delete clone._rid;
      delete clone._self;
      delete clone._ts;
      delete clone.ttl;
      return clone;
    }
  }

  async setObject(object: any): Promise<void> {
    this.throwIfNotInitialized();
    debug(`COSMOS SET OBJECT: ${object.id}`);
    let approxSize: number = 0;
    try {
      const item = Object.assign({}, object);
      await this._collection.items.upsert(item);
    } catch (upsertError) {
      console.dir(upsertError);
      console.log(approxSize);
      throw upsertError;
    }
  }

  async delete(partitionKey, documentId: string): Promise<void> {
    await this._collection.item(documentId, partitionKey).delete();
  }

  private throwIfNotInitialized() {
    if (!this._initialized) {
      throw new Error(
        'This provider must be initialized before it can be used'
      );
    }
  }
}
