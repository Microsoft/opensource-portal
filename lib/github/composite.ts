//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["log"] }] */

// Composite method/filter/call memoization

import { v4 as uuidV4 } from 'uuid';
import moment from 'moment';
import semver from 'semver';

const debug = require('debug')('restapi');

import {
  IShouldServeCache,
  ApiContext,
  IntelligentEngine,
  IApiContextRedisKeys,
  IApiContextCacheValues,
  ApiContextType,
  IRestMetadata,
  IRestResponse,
} from './core';
import { IGetAuthorizationHeader } from '../../transitional';

import appPackage from '../../package.json';

const appVersion = appPackage.version;

const longtermMinutes = 60 * 24 * 7; // 7 days
const acceleratedExpirationMinutes = 60; // 1 hour

export class CompositeApiContext extends ApiContext {
  private _apiMethod: any;
  private _apiTypePrefix: string;
  private _token: string | IGetAuthorizationHeader;
  private _cacheValues: IApiContextCacheValues;
  private _redisKeys: IApiContextRedisKeys;

  previouslyChanged: any;

  constructor(api: any, apiMethod: any, options: any) {
    super(api, options);

    const customApiTypePrefix = this.options.apiTypePrefix;
    if (customApiTypePrefix) {
      delete this.options.apiTypePrefix;
    }

    this._apiMethod = apiMethod;
    this._apiTypePrefix = customApiTypePrefix || 'github.col#';

    const root = IntelligentEngine.redisKeyForApi(
      this.apiTypePrefix,
      api,
      options
    );
    this._redisKeys = {
      root: root,
      metadata: root
        ? root + IntelligentEngine.redisKeyAspectSuffix('headers')
        : IntelligentEngine.redisKeyForApi(
            this.apiTypePrefix,
            api,
            options,
            'headers'
          ),
    };

    this._cacheValues = {
      longtermMetadata: longtermMinutes,
      longtermResponse: longtermMinutes,
      acceleratedExpiration: acceleratedExpirationMinutes,
    };
  }

  get redisKey(): IApiContextRedisKeys {
    return this._redisKeys;
  }

  get cacheValues(): IApiContextCacheValues {
    return this._cacheValues;
  }

  get token(): string | IGetAuthorizationHeader {
    return this._token;
  }

  overrideToken(token: string | IGetAuthorizationHeader) {
    this._token = token;
  }

  get apiMethod(): any {
    return this._apiMethod;
  }

  get apiTypePrefix(): string {
    return this._apiTypePrefix;
  }

  get contextType(): ApiContextType {
    return ApiContextType.Composite;
  }
}

export class CompositeIntelligentEngine extends IntelligentEngine {
  withMetadataShouldCacheBeServed(
    apiContext: ApiContext,
    metadata: IRestMetadata
  ): boolean | IShouldServeCache {
    // result can be falsy OR an object; { cache: true, refresh: true }
    // cache: whether to use the cache, if available
    // refresh: whether to refresh in the background for a newer value
    let shouldServeCache: IShouldServeCache | boolean = false;
    const maxAgeSeconds = apiContext.maxAgeSeconds;
    const updatedIso = metadata ? metadata.updated : null;
    const refreshingIso = metadata ? metadata.refreshing : null;
    if (metadata && !updatedIso) {
      debug(
        `${apiContext.redisKey.metadata} entity without updated date found`
      );
    }
    if (apiContext.generatedRefreshId) {
      debug(
        `${apiContext.redisKey.metadata} this is technically a refresh operation right now behind the scenes`
      );
    }
    if (maxAgeSeconds && updatedIso) {
      const updated = moment(updatedIso);
      const calledTime = apiContext.calledTime;
      if (updated.add(maxAgeSeconds, 'seconds').isAfter(calledTime)) {
        shouldServeCache = true;
        shouldServeCache = {
          cache: true,
          remaining:
            'expires ' +
            moment(updatedIso).add(maxAgeSeconds, 'seconds').fromNow(),
        };
        // debug('cache OK to serve as last updated was ' + updated);
      } else if (apiContext.backgroundRefresh) {
        let shouldRefresh = true;
        debug(
          apiContext.redisKey.metadata +
            ' need to go live as last updated ' +
            updated.format() +
            ' and our max seconds value is ' +
            maxAgeSeconds
        );
        if (refreshingIso) {
          let secondsToAllowForRefresh =
            2 + apiContext.delayBeforeRefreshMilliseconds / 1000;
          if (Array.isArray(metadata.pages)) {
            secondsToAllowForRefresh += metadata.pages.length * 1.25;
          }
          secondsToAllowForRefresh = Math.round(secondsToAllowForRefresh);
          const refreshWindow = moment(refreshingIso).add(
            secondsToAllowForRefresh,
            'seconds'
          );
          if (moment().utc().isAfter(refreshWindow)) {
            debug(
              `Another worker\'s refresh did not complete. Refreshing in this instance. ${apiContext.redisKey.metadata}`
            );
          } else {
            shouldRefresh = false;
            debug(
              `A refresh is already being processed by another worker. Allowing a window of ${secondsToAllowForRefresh}s before retry. ${apiContext.redisKey.metadata}`
            );
          }
        }
        shouldServeCache = {
          cache: true,
          refresh: shouldRefresh,
        };
      }
    } else {
      if (!metadata) {
        let reason = metadata === undefined ? 'undefined' : 'unknown';
        if (metadata === false) {
          reason = 'false value';
        } else if (metadata === null) {
          reason = 'null value';
        }
        debug(
          `composite: no metadata for key ${apiContext.redisKey.metadata} (${reason})`
        );
      } else {
        debug(
          `composite: no updated for key ${apiContext.redisKey.metadata} (but metadata present)`
        );
      }
    }
    return shouldServeCache;
  }

  withResponseShouldCacheBeServed(
    apiContext: ApiContext,
    response: IRestResponse
  ) {
    if (typeof response === 'function') {
      throw new Error('The response must not be a function');
    }
    if (response === undefined) {
      throw new Error(
        `${apiContext.redisKey.metadata}: the response was undefined and unable to process`
      );
    }
    if (!response.headers) {
      throw new Error(
        `${apiContext.redisKey.metadata}: no metadata was provided alongside the API response`
      );
    }
    let shouldUseCache = false;
    apiContext.etag = response.headers.etag;

    // Probably should check; if original data has not changed at all, then return true.
    // XXX: cannot find updated ever being set to false but...?
    debug(
      'composite.withResponseShouldCacheBeServed: not checking for updated = false so never serving cache'
    );
    // if (response && response.headers && response.headers.updated === false) {
    //   shouldUseCache = true;
    // }

    return shouldUseCache;
  }

  optionalStripResponse(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestResponse {
    // Composite does not strip any results further before caching
    return response;
  }

  withResponseUpdateMetadata(apiContext: ApiContext, response: IRestResponse) {
    return response;
  }

  reduceMetadataToCacheFromResponse(
    apiContext: ApiContext,
    response: IRestResponse
  ) {
    // No reduction for object type metadata.
    // Store the app version in case it is needed for a future
    // schema update or cache invalidation
    if (response.headers) {
      response.headers.av = appVersion;
      return response.headers;
    }
  }

  async callApi(apiContext: CompositeApiContext): Promise<IRestResponse> {
    const args = [];
    const apiMethod = apiContext.apiMethod;
    if (apiContext.token) {
      args.push(apiContext.token);
    }
    const argOptions = Object.assign({}, apiContext.options);
    args.push(argOptions);
    const thisArgument = apiMethod.thisInstance || null;
    let unknown = undefined;
    try {
      unknown = await apiMethod.apply(thisArgument, args);
    } catch (applyError) {
      throw applyError;
    }
    return unknown as IRestResponse;
  }

  getResponseMetadata(
    apiContext: CompositeApiContext,
    response: IRestResponse
  ): IRestMetadata {
    const headers = response.headers || {};
    let calledTime = apiContext.calledTime
      ? apiContext.calledTime.toISOString()
      : new Date().toISOString();
    headers.updated = calledTime;
    let changed = calledTime;
    if (headers.dirty === true) {
      changed = calledTime;
    } else if (headers.dirty === false && apiContext.previouslyChanged) {
      changed = apiContext.previouslyChanged;
    }
    headers.changed = changed;
    headers.etag = apiContext.generatedRefreshId || uuidV4();
    delete headers.dirty;
    return headers;
  }

  processMetadataBeforeCall(
    apiContext: CompositeApiContext,
    metadata: IRestMetadata
  ) {
    if (metadata && !metadata.av) {
      // Old version of metadata, no package version, which is required for all composite metadata now
      metadata = undefined;
    } else if (
      metadata &&
      metadata.av &&
      apiContext.libraryContext.breakingChangeGitHubPackageVersion &&
      !semver.gte(
        metadata.av,
        apiContext.libraryContext.breakingChangeGitHubPackageVersion
      )
    ) {
      console.log(
        `${apiContext.redisKey.metadata} was using ${metadata.av}, which is < to ${apiContext.libraryContext.breakingChangeGitHubPackageVersion}. This is a schema break, discarding cache.`
      );
      metadata = undefined;
    }
    if (metadata && metadata.etag) {
      apiContext.etag = metadata.etag;
      apiContext.metadata = metadata;
    }
    if (metadata && metadata.changed) {
      apiContext.previouslyChanged = metadata.changed;
    }
    return metadata;
  }
}
