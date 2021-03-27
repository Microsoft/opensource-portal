//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  BlobServiceClient,
  BlobItem,
  StorageSharedKeyCredential,
  ContainerClient,
  BlockBlobUploadResponse,
  BlobClient,
} from '@azure/storage-blob';

import { ICacheHelper } from '.';
import { gunzipBuffer, gzipString } from '../../utils';

const debug = require('debug')('cache');

export interface IBlobCacheOptions {
  account: string;
  key: string;
  container: string;
}

interface IExpiredBlobsStats {
  processedBlobs: number;
  processedPages: number;
  expired: number;
  errors: Error[];
}

interface ISetOptions {
  minutesToExpire?: number;
  extension?: string;
  compress?: boolean;
}

const ttlAttributeName = 'expires';
const compressedAttributeName = 'compressed';
const compressedGzip = 'gzip';

export default class BlobCache implements ICacheHelper {
  private _options: IBlobCacheOptions;
  private _client: BlobServiceClient;
  private _container: ContainerClient;
  private _initialized: boolean;

  constructor(options: IBlobCacheOptions) {
    this._options = options;
  }

  cloneForNewContainer(containerName: string) {
    const optionsClone = Object.assign({}, this._options, {
      container: containerName,
    });
    return new BlobCache(optionsClone);
  }

  readonly expiringBlobCache = this;

  private getBlobName(key: string, extension: string) {
    key = key.replace(/W\//g, ''); // stript the W/ from e-tags
    key = key.replace(/[@,\(\)\\\?#:=]/g, '/');
    key = key.replace(/\/\//g, '/');
    return `${key}${extension}`;
  }

  async initialize() {
    if (this._initialized) {
      return;
    }
    const { account, key } = this._options;
    if (!account) {
      throw new Error('options.account required');
    }
    if (!key) {
      throw new Error('options.key required');
    }
    if (!this._options.container) {
      throw new Error('options.container required');
    }
    const sharedKeyCredential = new StorageSharedKeyCredential(account, key);
    this._client = new BlobServiceClient(
      `https://${account}.blob.core.windows.net`,
      sharedKeyCredential
    );
    try {
      this._container = this._client.getContainerClient(
        this._options.container
      );
      if (!(await this._container.exists())) {
        await this._client.createContainer(this._options.container);
      }
    } catch (containerError) {
      console.dir(containerError);
    }
    this._initialized = true;
  }

  async get(key: string, extension?: string): Promise<string> {
    this.throwIfNotInitialized();
    const blobName = this.getBlobName(key, extension || '.txt');
    debug(`BLOB GET: ${blobName}`);
    const now = new Date();
    const blobClient = this._container.getBlobClient(blobName);
    try {
      const { metadata } = await blobClient.getProperties();
      if (metadata && metadata[ttlAttributeName]) {
        const asDate = new Date(metadata[ttlAttributeName]);
        if (asDate < now) {
          debug('EXPIRED on  get: ' + key);
          blobClient.delete(); // no need to wait for the promise to return
          return null;
        }
      }
      const buffer = await blobClient.downloadToBuffer();
      let value: string = null;
      if (metadata && metadata[compressedAttributeName] === compressedGzip) {
        value = await gunzipBuffer(buffer);
      } else if (metadata && metadata[compressedAttributeName]) {
        throw new Error(
          `Unsupported compression type: ${metadata[compressedAttributeName]}`
        );
      } else {
        value = buffer.toString('utf8');
      }
      return value;
    } catch (error) {
      if (error && error.statusCode && error.statusCode === 404) {
        return null;
      }
      console.dir(error);
      throw error;
    }
  }

  getCompressed(key: string): Promise<string> {
    return this.get(key, '.txt.gz');
  }

  async getObject(key: string): Promise<any> {
    const value = await this.get(key, '.json');
    const object = JSON.parse(value);
    return object;
  }

  async getObjectCompressed(key: string): Promise<any> {
    const value = await this.get(key, '.json.gz');
    const object = JSON.parse(value);
    return object;
  }

  async set(key: string, value: string, options?: ISetOptions): Promise<void> {
    options = options || {};
    const ext = options.extension || '.txt';
    const blobName = this.getBlobName(key, ext);
    this.throwIfNotInitialized();
    debug(`BLOB SET: ${blobName}`);
    const blockBlobClient = this._container.getBlockBlobClient(blobName);
    const uploadStarted = new Date();
    let metadata = {};
    let contentType =
      ext === '.json'
        ? 'application/json'
        : ext === '.txt'
        ? 'text/plain'
        : 'application/octet-stream';
    const blobHTTPHeaders = { blobContentType: contentType };
    if (options.minutesToExpire) {
      const expires = new Date(
        uploadStarted.getTime() + 1000 * 60 * options.minutesToExpire
      );
      const iso8601 = expires.toISOString();
      debug(
        `blob will expire in ${options.minutesToExpire}m; expires=${iso8601}, blob=${blobName}, key=${key}`
      );
      metadata[ttlAttributeName] = iso8601;
    }
    let response: BlockBlobUploadResponse = null;
    if (ext.endsWith('.gz') && !options.compress) {
      console.warn(
        `Warning, extension ${ext} for blobName ${blobName} appears compressed but options have not set the value to be compressed`
      );
    }
    if (options.compress) {
      const compressed = await gzipString(value);
      metadata[compressedAttributeName] = compressedGzip;
      response = await blockBlobClient.upload(
        compressed,
        compressed.byteLength,
        { blobHTTPHeaders, metadata }
      );
    } else {
      response = await blockBlobClient.upload(value, Buffer.byteLength(value), {
        blobHTTPHeaders,
        metadata,
      });
    }
  }

  async setObject(key: string, object: any): Promise<void> {
    const asJsonText = JSON.stringify(object);
    return this.set(key, asJsonText, { extension: '.json' });
  }

  setObjectWithExpire(
    key: string,
    object: any,
    minutesToExpire: number
  ): Promise<void> {
    if (object.ttl) {
      console.warn(
        "The object should not have an existing 'ttl' property before caching."
      );
    }
    const asJsonText = JSON.stringify(object);
    return this.set(key, asJsonText, { minutesToExpire, extension: '.json' });
  }

  setObjectCompressedWithExpire(
    key: string,
    object: any,
    minutesToExpire: number
  ): Promise<void> {
    const asJsonText =
      object.valueOnly === true ? object.value : JSON.stringify(object);
    const extension = object.valueOnly === true ? '.txt.gz' : '.json.gz';
    return this.set(key, asJsonText, {
      minutesToExpire,
      extension,
      compress: true,
    });
  }

  setCompressed(key: string, value: string): Promise<void> {
    return this.set(key, value, { extension: '.txt.gz', compress: true });
  }

  setCompressedWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void> {
    return this.setObjectCompressedWithExpire(
      key,
      {
        value,
        valueOnly: true,
      },
      minutesToExpire
    );
  }

  setWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void> {
    return this.setCompressedWithExpire(key, value, minutesToExpire);
  }

  async expire(key: string, minutesToExpire: number): Promise<void> {
    this.throwIfNotInitialized();
    const candidateExtensions = ['.txt', '.txt.gz', '.json', '.json.gz'];
    const now = new Date();
    while (candidateExtensions.length) {
      const extension = candidateExtensions.pop();
      let metadata = null;
      let blobClient: BlobClient = null;
      try {
        blobClient = this._container.getBlobClient(
          this.getBlobName(key, extension)
        );
        const resp = (await blobClient.getProperties()) || {};
        metadata = resp['metadata'];
      } catch (notFoundError) {
        continue;
      }
      const expires = new Date(now.getTime() + 1000 * 60 * minutesToExpire);
      const iso8601 = expires.toISOString();
      metadata[ttlAttributeName] = iso8601;
      await blobClient.setMetadata(metadata);
      return;
    }
  }

  async delete(key: string): Promise<void> {
    const candidateExtensions = ['.txt', '.txt.gz', '.json', '.json.gz'];
    while (candidateExtensions.length) {
      const extension = candidateExtensions.pop();
      try {
        const blobClient = this._container.getBlobClient(
          this.getBlobName(key, extension)
        );
        await blobClient.delete();
        return;
      } catch (ignoredError) {}
    }
  }

  async deleteExpiredBlobs(): Promise<IExpiredBlobsStats> {
    let iterator = this._container
      .listBlobsFlat({ includeMetadata: true })
      .byPage({ maxPageSize: 100 /* 25 */ });
    let response = await iterator.next();
    const stats: IExpiredBlobsStats = {
      processedBlobs: 0,
      processedPages: 0,
      expired: 0,
      errors: [],
    };
    let x = 0;
    while (!response.done) {
      const segment = response.value.segment;
      for (const b of segment.blobItems) {
        const blob = b as BlobItem;
        try {
          if (!blob.metadata || !blob.metadata.expires) {
            debug(
              `${++x}. FYI: blob ${
                blob.name
              } does not have an expiration to review and was skipped`
            );
            ++stats.processedBlobs;
            continue;
          }
          const expires = new Date(blob.metadata.expires);
          const now = new Date();
          if (now > expires) {
            debug(blob.name);
            const blobClient = this._container.getBlobClient(blob.name);
            await blobClient.delete();
            debug(`expired ${blob.name}`);
            ++stats.expired;
          }
        } catch (processBlobError) {
          console.dir(processBlobError);
          stats.errors.push(processBlobError);
        }
        ++stats.processedBlobs;
        debug(
          `processed=${stats.processedBlobs}, pages=${stats.processedPages}, expired=${stats.expired}`
        );
      }
      ++stats.processedPages;
      response = await iterator.next();
    }
    return stats;
  }

  private throwIfNotInitialized() {
    if (!this._initialized) {
      throw new Error(
        'This provider must be initialized before it can be used'
      );
    }
  }
}
