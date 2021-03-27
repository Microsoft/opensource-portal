//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["dir", "warn", "log"] }] */

import _ from 'lodash';
const debug = require('debug')('restapi');
import { v4 as uuidV4 } from 'uuid';
import moment from 'moment';
import { RestLibrary } from '.';
import {
  IAuthorizationHeaderValue,
  ErrorHelper,
  IDictionary,
} from '../../transitional';
import { sleep } from '../../utils';

import cost from './cost';

const delayBeforeRefreshMilliseconds = 1000;

const delayBeforeRequestMilliseconds = false; // useful during cache fill: 2 * 1000;

// --- Intelligent cache object interfaces ---

export enum ApiContextType {
  Composite,
  GitHubRestApi,
}

export interface IRestResponseHeaders {
  etag?: string;
  link?: unknown;
}

const headerKeysWanted = [
  'etag',
  'last-modified',
  'x-github-request-id',
  'x-ratelimit-limit',
  'x-ratelimit-remaining',
  'x-ratelimit-reset',
  'x-ratelimit-used',
];

export interface IInterestingHeaders {
  etag?: string;
  ['last-modified']?: string;
  ['x-github-request-id']?: string;
  ['x-ratelimit-limit']?: string;
  ['x-ratelimit-remaining']?: string;
  ['x-ratelimit-reset']?: string;
  ['x-ratelimit-used']?: string;
}

export interface ISpecializedCollectionHeaders {
  // really, these are the metadata fields, no headers at all...
  dirty?: boolean;
  pages?: string[];
  etag?: string;
  link?: string;
  av?: string;
  updated?: string;
  changed?: string;
  ['last-modified']?: string;
}

export interface IRestMetadata {
  etag?: string;
  av?: string;
  updated?: string;
  changed?: string;
  refreshing?: string;
  headers?: ISpecializedCollectionHeaders; // IDictionary<string>;
  status?: number;
  pages?: string[];
}

export interface IRestResponse {
  headers?: ISpecializedCollectionHeaders;
  status?: number;
  data: unknown;
  cost?: unknown;
  notModified?: boolean;
}

export interface IIntelligentCacheResponseArray
  extends Array<any>,
    IRestResponse {}

export interface IShouldServeCache {
  cache?: boolean;
  remaining?: string;
  refresh?: boolean;
}

export abstract class ApiContext {
  private _log: string[];
  private _calledTime: Date;
  private _cost: any;

  libraryContext: RestLibrary;
  etag?: string;
  tokenSource: IAuthorizationHeaderValue;

  abstract get apiTypePrefix(): string;
  abstract get cacheValues(): IApiContextCacheValues;
  abstract get redisKey(): IApiContextRedisKeys;
  abstract get contextType(): ApiContextType;

  get calledTime(): Date {
    return this._calledTime;
  }

  get delayBeforeRefreshMilliseconds(): number {
    return delayBeforeRefreshMilliseconds;
  }

  get log(): string[] {
    return this._log;
  }

  get cost(): any {
    return this._cost;
  }

  constructor(public api: any, public options: any) {
    this._log = [];
    if (!this._calledTime) {
      this._calledTime = new Date();
    }
    if (!this._cost) {
      this._cost = cost.create();
    }
  }

  maxAgeSeconds?: number;
  backgroundRefresh?: boolean;

  metadata?: IRestMetadata;

  generatedRefreshId?: string;
}

export interface IApiContextRedisKeys {
  metadata: string;
  root: string;
}

export interface IApiContextCacheValues {
  longtermMetadata: number;
  longtermResponse: number;
  acceleratedExpiration: number;
}

export abstract class IntelligentEngine {
  // in hindsight, "intelligent" is not what any of this is
  public static redisKeyAspectSuffix(aspect: string): string {
    return aspect ? `:${aspect}` : '';
  }

  public static redisKeyForApi(
    apiPrefix: string,
    api: string,
    apiOptions,
    aspect?: string
  ) {
    const normalizedOptions = normalizedOptionsString(apiOptions);
    const aspectSuffix = IntelligentEngine.redisKeyAspectSuffix(aspect);
    return `${apiPrefix}${api}${normalizedOptions}${aspectSuffix}`;
  }

  // was in api context:
  abstract processMetadataBeforeCall(
    apiContext: ApiContext,
    metadata: IRestMetadata
  ): IRestMetadata;
  abstract callApi(
    apiContext: ApiContext,
    optionalMessage?: string
  ): Promise<IRestResponse>;
  abstract withResponseUpdateMetadata(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestResponse;

  abstract withResponseShouldCacheBeServed(
    apiContext: ApiContext,
    response: IRestResponse
  ): boolean | IShouldServeCache;
  abstract withMetadataShouldCacheBeServed(
    apiContext: ApiContext,
    metadata: IRestMetadata
  ): boolean | IShouldServeCache;
  abstract reduceMetadataToCacheFromResponse(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestMetadata;
  abstract getResponseMetadata(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestMetadata;
  abstract optionalStripResponse(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestResponse;

  protected async cacheResponseAsync(
    apiContext: ApiContext,
    response: IRestResponse
  ) {
    const backgroundAsyncWork = async () => {
      try {
        await this.storeResult(apiContext, response);
        await this.storeMetadata(apiContext, response);
        await this.reduceObjectExpirationWindow(apiContext, response);
      } catch (err) {
        console.log(`Background async work (cacheResponseAsync): ${err}`);
        console.dir(err);
      }
    };
    backgroundAsyncWork()
      .then((ok) => {})
      .catch(() => {});
    return this.finalizeResult(apiContext, response);
  }

  protected finalizeResult(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestResponse {
    if (!response || !response.data) {
      // This was a warning in the past, but to try and improve the underlying library, this should be an error
      if (response.headers.av) {
        console.log('cached version was from: v' + response.headers.av);
      }
      const reason = !response ? 'no response' : 'no response.data';
      const warning = `${apiContext.redisKey.root} : ${reason}`;
      console.warn(warning);
      // this.evict(apiContext)
      //   .then(ok => {}, rejected => {});
      throw new Error(warning);
    }
    // If there are situations where you do not want the cost shared
    // back the API context could be customized here.
    if (response) {
      response.cost = apiContext.cost;
    }
    return response;
  }

  protected async tryGetCachedResult(
    apiContext: ApiContext
  ): Promise<IRestResponse> {
    const key = this.redisKeyBodyVersion(apiContext);
    let response = (await apiContext.libraryContext.cacheProvider.getObjectCompressed(
      key
    )) as IRestResponse;
    this.recordRedisCost(apiContext, 'get', response);
    return response;
  }

  protected async getCachedResult(
    apiContext: ApiContext,
    optionalCacheDecisions?,
    notModifiedHeaders?: IInterestingHeaders
  ): Promise<IRestResponse> {
    let result = await this.tryGetCachedResult(apiContext);
    if (result && result.data) {
      // use the context metadata over any headers in the stored response, + any headers from 304
      result.headers = Object.assign(
        {},
        notModifiedHeaders || {},
        apiContext.metadata
      );
      if (optionalCacheDecisions && optionalCacheDecisions.refresh === true) {
        // NOTE: this kicks off the refresh and so does not await
        debug('Starting a background refresh');
        this.backgroundRefreshAsync(apiContext, apiContext.metadata)
          .then((ok) => {})
          .catch(() => {});
      } else {
        this.slideObjectExpirationWindow(apiContext)
          .then((ok) => {})
          .catch(() => {});
      }
      debug('Finalizing result');
      return this.finalizeResult(apiContext, result);
    }
    debug('Cache miss.');
    if (result) {
      this.evict(apiContext)
        .then(() => {
          console.log('(evicted)');
        })
        .catch((err) => {
          console.warn(`(eviction error: ${err})`);
        });
    }
    ++apiContext.cost.redis.cacheMisses;
    delete apiContext.etag;

    let response = await this.callApi(apiContext);
    response = await this.processResponse(apiContext, response);
    return response;
  }

  protected async backgroundRefreshAsync(
    apiContext: ApiContext,
    currentMetadata
  ): Promise<void> {
    // Potential data loss/consistency problem: upsert/overwrite
    try {
      let refreshing = moment().utc().format();
      let refreshId = uuidV4();
      currentMetadata.refreshing = refreshing;
      currentMetadata.refreshId = refreshId;
      apiContext.generatedRefreshId = refreshId;
      debug(
        `refresh in the background starting for ${apiContext.redisKey.metadata} was updated ${apiContext.metadata.updated} and seconds of ${apiContext.maxAgeSeconds}`
      );
      // TODO: use proper next tick to kick this off?
      const setReturnValue = await apiContext.libraryContext.cacheProvider.setObjectWithExpire(
        apiContext.redisKey.metadata,
        currentMetadata,
        apiContext.cacheValues.longtermMetadata
      );
      // Remove the values in case the refresh uses the metadata
      delete currentMetadata.refreshing;
      delete currentMetadata.refreshId;
      this.recordRedisCost(apiContext, 'set', setReturnValue);
      await sleep(delayBeforeRefreshMilliseconds);
      let response = await this.callApi(apiContext);
      response = await this.processResponse(apiContext, response);
      // ? anything to return... I don't think so
    } catch (backgroundError) {
      if (backgroundError.status === 404) {
        // gone entity
      } else if (backgroundError.status === 304) {
        // did not change
      } else {
        console.dir(backgroundError);
      }
    }
  }

  // --- Caching ---

  protected async reduceObjectExpirationWindow(
    apiContext: ApiContext,
    response: IRestResponse
  ): Promise<void> {
    if (
      !apiContext.etag ||
      (apiContext.etag && apiContext.etag === response.headers.etag)
    ) {
      return;
    }
    debug('Expiring older cached response');
    const cost = await apiContext.libraryContext.cacheProvider.expire(
      this.redisKeyBodyVersion(apiContext, apiContext.etag),
      apiContext.cacheValues.acceleratedExpiration
    );
    this.recordRedisCost(apiContext, 'expire', cost);
  }

  protected async evict(apiContext: ApiContext): Promise<void> {
    if (!apiContext.etag) {
      debug('No etag to evict.');
      return;
    }
    const key = this.redisKeyBodyVersion(apiContext, apiContext.etag);
    await apiContext.libraryContext.cacheProvider.delete(key);
  }

  protected async slideObjectExpirationWindow(
    apiContext: ApiContext
  ): Promise<void> {
    if (!apiContext.etag) {
      debug('Could not slide the window, no etag stored in the context.');
      return;
    }
    debug(
      `Sliding expiration window for ${this.redisKeyBodyVersion(
        apiContext,
        apiContext.etag
      )}`
    );
    const cost = await apiContext.libraryContext.cacheProvider.expire(
      this.redisKeyBodyVersion(apiContext, apiContext.etag),
      apiContext.cacheValues.longtermResponse
    );
    this.recordRedisCost(apiContext, 'expire', cost);
  }

  protected async storeMetadata(
    apiContext: ApiContext,
    response: IRestResponse
  ): Promise<void> {
    const reducedMetadata = this.reduceMetadataToCacheFromResponse(
      apiContext,
      response
    );
    const cost = await apiContext.libraryContext.cacheProvider.setObjectWithExpire(
      apiContext.redisKey.metadata,
      reducedMetadata,
      apiContext.cacheValues.longtermMetadata
    );
    this.recordRedisCost(apiContext, 'set', cost);
  }

  protected async storeResult(
    apiContext: ApiContext,
    response: IRestResponse
  ): Promise<void> {
    let key = null;
    try {
      key = this.redisKeyBodyVersion(apiContext, response.headers.etag);
    } catch (noKey) {
      return;
    }
    const cost = await apiContext.libraryContext.cacheProvider.setObjectCompressedWithExpire(
      key,
      response,
      apiContext.cacheValues.longtermResponse
    );
    this.recordRedisCost(apiContext, 'set', cost);
  }

  protected redisKeyBodyVersion(apiContext: ApiContext, etag?: string): string {
    const tag = etag || apiContext.etag;
    if (!tag) {
      throw new Error(
        'A cached result cannot be retrieved without an etag value.'
      );
    }
    const strippedTag = tag.replace(/"/g, '');
    const root = apiContext.redisKey.root;
    if (!root) {
      throw new Error(
        'No Redis key root provided in API context apiContext.redisKey.root'
      );
    }
    return root + IntelligentEngine.redisKeyAspectSuffix(`body@${strippedTag}`);
  }

  protected recordRedisCost(
    apiContext: ApiContext,
    type: string,
    object: any
  ): any {
    if (!type) {
      throw new Error('No type defined for recordRedisCost.');
    }
    let hit = object !== undefined;
    if (type === 'get') {
      apiContext.cost.redis.cacheHit += hit ? 1 : 0;
      apiContext.cost.redis.cacheMisses += hit ? 0 : 1;
    }
    if (type !== 'get' && type !== 'set' && type !== 'expire') {
      throw new Error(
        `The Redis type of ${type} is not configured for storing API costs.`
      );
    }
    apiContext.cost.redis[`${type}Calls`] += 1;
    return object;
  }

  public async execute(apiContext: ApiContext): Promise<IRestResponse> {
    let metadata = await this.getCachedMetadata(apiContext);
    metadata = this.processMetadataBeforeCall(apiContext, metadata);
    const shouldCacheBeServedImmediately:
      | boolean
      | IShouldServeCache = this.withMetadataShouldCacheBeServed(
      apiContext,
      metadata
    );
    const displayKey = apiContext.redisKey
      ? apiContext.redisKey.root + ' '
      : '';
    if (
      shouldCacheBeServedImmediately === true ||
      (shouldCacheBeServedImmediately as IShouldServeCache).cache === true
    ) {
      debug('Cache should be served immediately.');
      if (metadata) {
        const innerMessage =
          shouldCacheBeServedImmediately &&
          (shouldCacheBeServedImmediately as IShouldServeCache).remaining
            ? (shouldCacheBeServedImmediately as IShouldServeCache).remaining
            : '';
        debug(`Cache ${displayKey}data: ${innerMessage}`);
      }
      ++apiContext.cost.github.cacheHits;
      const cachedResponse = await this.getCachedResult(
        apiContext,
        shouldCacheBeServedImmediately
      );
      debug('Retrieved a cached response.');
      return cachedResponse;
    }
    if (delayBeforeRequestMilliseconds) {
      debug(`DELAY...: ${displayKey} ${delayBeforeRefreshMilliseconds}`);
      await sleep(delayBeforeRefreshMilliseconds);
    }
    debug('Directly calling the function or REST API');
    let response: IRestResponse = undefined;
    try {
      response = (await this.callApi(
        apiContext,
        `GET:               ${displayKey}`
      )) as IRestResponse;
    } catch (error) {
      if (error && error.status && error.status === 304) {
        const liveHeaders = error.headers || {};
        const headers = {};
        for (let i = 0; i < headerKeysWanted.length; i++) {
          const key = headerKeysWanted[i];
          if (liveHeaders[key]) {
            headers[key] = liveHeaders[key];
          }
        }
        const notModifiedResponse = {
          data: undefined,
          headers,
          notModified: true,
        };
        response = notModifiedResponse;
      } else {
        throw error;
      }
    }
    response = await this.processResponse(apiContext, response);
    return response;
  }

  private async processResponse(
    apiContext: ApiContext,
    response: IRestResponse
  ): Promise<IRestResponse> {
    this.withResponseUpdateMetadata(apiContext, response);
    const isCacheOk = this.withResponseShouldCacheBeServed(
      apiContext,
      response
    );
    if (isCacheOk === true) {
      ++apiContext.cost.github.cacheHits;
      debug('Cache is OK to retrieve and serve');
      const notModifiedHeaders = response.headers;
      return await this.getCachedResult(apiContext, null, notModifiedHeaders);
    }
    debug('Cache should not be served. Reading the response metadata');
    ++apiContext.cost.github.usedApiTokens;
    const metadata = this.getResponseMetadata(apiContext, response);
    if (metadata) {
      const responseToCache = this.optionalStripResponse(apiContext, response);
      debug('Caching the live response and metadata');
      return await this.cacheResponseAsync(apiContext, responseToCache); // finalizeResult will happen after caching
    } else {
      debug('Finalizing result. There was no metadata with the response');
      return this.finalizeResult(apiContext, response);
    }
  }

  private async getCachedMetadata(
    apiContext: ApiContext
  ): Promise<IRestMetadata> {
    if (apiContext.metadata || apiContext.etag) {
      debug('Shortcut: apiContext.metadata or apiContext.etag are set');
      return;
    }
    const redisKey = apiContext.redisKey.metadata;
    if (!redisKey) {
      throw new Error('No Redis key provided in apiContext.redisKey.metadata');
    }
    const cachedMetadata: IRestMetadata = (await apiContext.libraryContext.cacheProvider.getObject(
      redisKey
    )) as IRestMetadata;
    // debug('Cached metadata retrieved');
    this.recordRedisCost(apiContext, 'get', cachedMetadata);
    return cachedMetadata;
  }

  // end class after
}

// --- Caching integration with metadata/responses ---

function normalizedOptionsString(options) {
  if (!options) {
    return '';
  }
  let additional = null;
  if (options.additionalDifferentiationParameters) {
    additional = options.additionalDifferentiationParameters;
  }
  let opts = { ...options, ...additional };
  if (opts.additionalDifferentiationParameters) {
    delete opts.additionalDifferentiationParameters;
  }
  const sortedkeys = _.keys(opts).sort();
  let normalized = [];
  sortedkeys.forEach((key) => {
    let value = opts[key];
    const typeOf = typeof value;
    if (typeOf === 'undefined') {
      return;
    }
    if (typeOf === 'object') {
      value = normalizedOptionsString(value);
    } else if (
      typeOf !== 'string' &&
      typeOf !== 'number' &&
      typeOf !== 'boolean'
    ) {
      throw new Error(`Normalized option ${key} is not a string`);
    }
    if (typeOf === 'boolean') {
      value = value === true ? '1' : '0';
    }
    normalized.push(`${key}=${value}`);
  });
  return `(${normalized.join(',')})`;
}

function projectFlatObjectWithData(entity) {
  const flat = entity.data;
  if (entity.cost) {
    flat.cost = entity.cost;
  }
  if (flat && entity.headers) {
    flat.headers = entity.headers;
  }
  return flat;
}

export function flattenData(entity: any): any {
  if (!entity) {
    return entity;
  }
  if (entity.data !== undefined && !entity.data) {
    // If it's an empty string, or false, etc., return the value directly
    return entity.data;
  }
  const flat = projectFlatObjectWithData(entity);
  return flat;
}

export function createCallbackFlattenDataOptionally(callback) {
  return function callbackFlatEntityFromDataOptionally(error, entity) {
    if (error) {
      return callback(error);
    }
    if (!entity) {
      return callback(null, entity);
    }
    if (entity.data !== undefined && !entity.data) {
      // If it's an empty string, or false, etc., return the value directly
      return callback(null, entity.data);
    }
    const flat = projectFlatObjectWithData(entity);
    return callback(null, flat);
  };
}

export function createCallbackFlattenData(callback) {
  return function callbackFlatEntityFromData(error, entity) {
    if (!error && entity && !entity.data) {
      error = new Error(
        'No entity.data present in the result, cannot flatten the object'
      );
    }
    if (!error) {
      try {
        const flat = projectFlatObjectWithData(entity);
        return callback(null, flat);
      } catch (flattenError) {
        error = flattenError;
      }
    }
    return callback(error);
  };
}
