//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ICacheHelper } from '.';
import { CosmosClient, Database, Container } from '@azure/cosmos';
import BlobCache, { IBlobCacheOptions } from './blob';
import { sleep } from '../../utils';
import { ErrorHelper, sha256 } from '../../transitional';

const debug = require('debug')('cache');

export interface ICosmosCacheOptions {
  endpoint: string;
  key: string;
  database?: string;
  collection?: string;
  prefix?: string;

  blobFallback?: IBlobCacheOptions;
}

const objectIdKeyReplacement = 'cached_id';
const largess = 0.5 * 1000000;
const cut = 1.8 * 1000000;

export default class CosmosCache implements ICacheHelper {
  private _options: ICosmosCacheOptions;
  private _client: CosmosClient;
  private _initialized: boolean;
  private _database: Database;
  private _collection: Container;
  private _blobCache: BlobCache;
  private _keyPrefix: string;

  constructor(options: ICosmosCacheOptions) {
    this._options = options;
    this._keyPrefix = options.prefix || '';
    if (options.blobFallback && options.blobFallback.key) {
      this._blobCache = new BlobCache(options.blobFallback);
    }
  }

  cloneWithPrefix(newPrefix: string) {
    const options = Object.assign({}, this._options, {
      prefix: newPrefix,
    });
    return new CosmosCache(options);
  }

  get expiringBlobCache() {
    return this._blobCache;
  }

  private safetyKey(str: string) {
    return str.replace(/[%:\\\/\?#]/g, '');
  }

  private key(key: string) {
    // https://graph.microsoft.com/v1.0/groups/ec78f98a-6e8c-4685-87f6-6fa18fd86183/transitiveMembers?%24select=id%2CuserPrincipalName
    const value = this.safetyKey(this._keyPrefix + key);
    if (value.length > 255) {
      const hash = sha256(value);
      debug(
        `hash ${hash} overrides long key length=${value.length} key=${value}`
      );
      const newKey =
        value.substr(0, 20) + '_longkeyhash_' + hash.replace('=', '');
      return this.safetyKey(newKey);
    }
    return value;
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
    if (this._blobCache) {
      await this._blobCache.initialize();
    }
    this._initialized = true;
  }

  async get(key: string): Promise<string> {
    this.throwIfNotInitialized();
    key = this.key(key);
    debug(`COSMOS GET: ${key}`);
    let response = null;
    try {
      response = await this._collection.item(key, key).read();
    } catch (cosmosError) {
      console.dir(cosmosError);
      throw cosmosError;
    }
    if (response.resource) {
      if (!response.resource.value) {
        throw new Error('The retrieved resource does not have a single value');
      }
      return response.resource.value as string;
    }
  }

  getCompressed(key: string): Promise<string> {
    return this.get(key); // NOTE: CosmosDB does not support compression
  }

  async getObject(key: string): Promise<any> {
    key = this.key(key);
    this.throwIfNotInitialized();
    debug(`COSMOS GET OBJECT: ${key}`);
    let response = null;
    try {
      response = await this._collection.item(key, key).read();
    } catch (cosmosError) {
      console.dir(cosmosError);
      throw cosmosError;
    }
    if (response.resource?.blobKey) {
      if (!this._blobCache) {
        return null; // or throw
      }
      const compressed =
        response.resource.compress && response.resource.compress === true;
      return compressed
        ? await this._blobCache.getObjectCompressed(response.resource.blobKey)
        : await this._blobCache.getObject(response.resource.blobKey);
    } else if (response.resource?.chunks && response.resource.chunk) {
      const num = response.resource.chunks as number;
      debug('LARGE retrieval: ' + num + ' chunks found');
      const chunky = response.resource;
      let combined = response.resource.chunk;
      for (let i = 1; i < num; i++) {
        const ck = this.key(`${key}_c${i}`);
        const innerChunk = await this._collection.item(ck, ck).read();
        combined += innerChunk.resource.chunk;
      }
      const clone = JSON.parse(combined);
      if (clone[objectIdKeyReplacement]) {
        clone.id = clone[objectIdKeyReplacement];
        delete clone[objectIdKeyReplacement];
      }
      return clone;
    }
    const clone = Object.assign({}, response.resource);
    delete clone._attachments;
    delete clone._etag;
    delete clone._rid;
    delete clone._self;
    delete clone._ts;
    delete clone.ttl;
    if (clone[objectIdKeyReplacement]) {
      clone.id = clone[objectIdKeyReplacement];
      delete clone[objectIdKeyReplacement];
    }
    return clone;
  }

  getObjectCompressed(key: string): Promise<any> {
    return this.getObject(key); // NOTE: CosmosDB does not support compression
  }

  async set(key: string, value: string): Promise<void> {
    key = this.key(key);
    this.throwIfNotInitialized();
    debug(`COSMOS SET: ${key}`);
    const item = Object.assign({
      id: key,
      value,
    });
    try {
      await this._collection.items.upsert(item);
    } catch (upsertError) {
      console.dir(upsertError);
      throw upsertError;
    }
  }

  async setObject(key: string, object: any): Promise<void> {
    this.throwIfNotInitialized();
    const originalKey = this._keyPrefix + key;
    key = this.key(key);
    debug(`COSMOS SET OBJECT: ${key}`);
    let approxSize: number = 0;
    try {
      if (object.id && object.id !== key) {
        // SAVE ORIGINAL ID PROPERTY
        object[objectIdKeyReplacement] = object.id;
        delete object.id;
      }
      let item = Object.assign({}, object, { id: key });
      const asJsonText = JSON.stringify(item);
      approxSize = bytes(asJsonText);
      if (approxSize > largess) {
        if (this._blobCache) {
          debug(
            `storing in blob instead... key=${originalKey}, dataSize=${approxSize}`
          );
          const ttlSeconds = object.ttl;
          delete object.ttl;
          (await ttlSeconds)
            ? this._blobCache.setObjectCompressedWithExpire(
                originalKey,
                object,
                ttlSeconds / 60
              )
            : this._blobCache.setObject(originalKey, object);
          item = {
            blobKey: originalKey,
            id: key,
          };
          if (ttlSeconds) {
            item['compress'] = true;
          }
          approxSize = 500;
        }
      }
      if (approxSize > cut) {
        const chunks = this.intoChunks(asJsonText, cut);
        debug(
          `LARGE Cosmos save, would blob be better? chunks: ${chunks.length}`
        );
        for (let i = 0; i < chunks.length; i++) {
          const id = i === 0 ? key : `${key}_c${i}`;
          const chunkDoc = {
            i,
            id,
            chunk: chunks[i],
            chunks: chunks.length,
          };
          if (item.ttl) {
            chunkDoc['ttl'] = item.ttl;
          }
          let chunkSize = 0;
          try {
            chunkSize = bytes(JSON.stringify(chunkDoc));
            await this._collection.items.upsert(chunkDoc);
            debug(
              `chunk saved for key ${key}, id=${chunkDoc.id}, sizeApprox=${chunkDoc.chunk.length}`
            );
          } catch (chunkSave) {
            console.dir(chunkSave);
            debug(`may be too large at ${chunkSize}... ${chunkSave}`);
          }
        }
      } else {
        let tries = 0;
        let success = false;
        while (success === false && tries < 10) {
          ++tries;
          try {
            await this._collection.items.upsert(item);
            success = true;
          } catch (cosmosError) {
            if (cosmosError && cosmosError.code && cosmosError.code === 429) {
              const time =
                cosmosError && cosmosError.retryAfterInMs
                  ? cosmosError.retryAfterInMs
                  : 500;
              debug(`pausing, will retry on 429 in ${time}ms`);
              await sleep(time);
            } else {
              tries += 10;
              throw cosmosError;
            }
          }
        }
      }
    } catch (upsertError) {
      console.dir(upsertError);
      console.log(approxSize);
      throw upsertError;
    }
  }

  private intoChunks(longString: string, chunkLength: number): string[] {
    const chunks = [];
    const slightlyLess = Math.floor(chunkLength);
    let remainder = longString;
    while (remainder.length) {
      const chunk = remainder.substr(0, slightlyLess);
      const cs = bytes(chunk);
      chunks.push(chunk);
      remainder = remainder.substr(slightlyLess);
    }
    return chunks;
  }

  setObjectWithExpire(
    key: string,
    object: any,
    minutesToExpire: number
  ): Promise<void> {
    if (object.ttl) {
      console.warn(
        "Warning: the object has an existing 'ttl' property before caching."
      );
    }
    object.ttl = minutesToExpire * 60;
    return this.setObject(key, object);
  }

  setObjectCompressedWithExpire(
    key: string,
    object: any,
    minutesToExpire: number
  ): Promise<void> {
    return this.setObjectWithExpire(key, object, minutesToExpire);
  }

  setCompressed(key: string, value: string): Promise<void> {
    return this.set(key, value); // NOTE: CosmosDB does not support compression
  }

  setCompressedWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void> {
    return this.setObjectWithExpire(
      key,
      {
        value,
      },
      minutesToExpire
    );
  }

  setWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void> {
    return this.setCompressedWithExpire(key, value, minutesToExpire); // NOTE: CosmosDB does not support compression
  }

  async expire(key: string, minutesToExpire: number): Promise<void> {
    this.throwIfNotInitialized();
    // this may not be the best implementation but this should work...
    const current = await this.getObject(key);
    if (current) {
      delete current.ttl;
      await this.setObjectWithExpire(key, current, minutesToExpire);
    }
  }

  async delete(key: string): Promise<void> {
    try {
      key = this.key(key);
      const item = this._collection.item(key, key);
      await item.delete();
    } catch (cosmosError) {
      if (!ErrorHelper.IsNotFound(cosmosError)) {
        console.dir(cosmosError);
      }
      throw cosmosError;
    }
  }

  private throwIfNotInitialized() {
    if (!this._initialized) {
      throw new Error(
        'This provider must be initialized before it can be used'
      );
    }
  }
}

const bytes = (s) => {
  return ~-encodeURI(s).split(/%..|./).length;
};
