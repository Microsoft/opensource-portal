//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["log", "warn"] }] */

import _ from 'lodash';
import moment from 'moment';
import semver from 'semver';

const debug = require('debug')('restapi');
const debugCacheOptimization = require('debug')('oss-cache-optimization');

const debugShowStandardBehavior = false;
const debugOutputUnregisteredEntityApis = true;

import {
  IShouldServeCache,
  IntelligentEngine,
  ApiContext,
  IApiContextCacheValues,
  IApiContextRedisKeys,
  ApiContextType,
  IRestResponse,
  IRestMetadata,
} from './core';
import {
  getEntityDefinitions,
  GitHubResponseType,
  ResponseBodyType,
} from './endpointEntities';

import appPackage from '../../package.json';
import {
  IGetAuthorizationHeader,
  IAuthorizationHeaderValue,
} from '../../transitional';

const appVersion = appPackage.version;

const longtermMetadataMinutes = 60 * 24 * 14; // assumed to be a long time
const longtermResponseMinutes = 60 * 24 * 7; // a week, sliding
const acceleratedExpirationMinutes = 10; // quick cleanup

const entityData = getEntityDefinitions();
const emptySet = new Set<string>();

interface IReducedGitHubMetadata {
  etag: string;
  av: string;
  link?: any;
  updated?: any;
}

interface IGitHubLink {
  link: string;
}

export class IntelligentGitHubEngine extends IntelligentEngine {
  public static findLibraryMethod(libraryInstance, apiName) {
    const instance = libraryInstance;
    const combined = apiName;
    const i = combined.indexOf('.');
    let apiGroup = null;
    let apiMethodName = combined;
    if (i >= 0) {
      apiGroup = combined.substr(0, i);
      apiMethodName = combined.substr(i + 1);
    }
    const group = apiGroup ? instance[apiGroup] : instance;
    if (!group) {
      throw new Error(
        `The GitHub REST API library does not support the API group of type "${apiGroup}".`
      );
    }
    const method = group[apiMethodName];
    if (!method) {
      throw new Error(
        `The GitHub REST API library does not support the API "${apiMethodName}" within the API group of type "${apiGroup}".`
      );
    }
    return method;
  }

  async callApi(
    apiContext: GitHubApiContext,
    optionalMessage?: string
  ): Promise<IRestResponse> {
    const token = apiContext.token;
    // CONSIDER: rename apiContext.token *to* something like apiContext.authorization
    if (
      typeof token === 'string' &&
      !(token as string).startsWith('token ') &&
      !(token as string).startsWith('bearer ')
    ) {
      if (optionalMessage) {
        debug(optionalMessage);
      }
      const warning = `API context api=${apiContext.api} does not have a token that starts with 'token [REDACTED]' or 'bearer [REDACTED], investigate this breakpoint`;
      throw new Error(warning);
    }
    let authorizationHeaderValue =
      typeof token === 'string' ? (token as string) : null;
    if (!authorizationHeaderValue) {
      if (typeof token === 'function') {
        const response = await token();
        if (typeof response === 'string') {
          // happens when it isn't a more modern GitHub app response
          authorizationHeaderValue = response;
        } else {
          const value = response['value'];
          if (!value) {
            throw new Error('No value');
          }
          authorizationHeaderValue = value;
          apiContext.tokenSource = response;
        }
      }
    }
    if (optionalMessage) {
      let apiTypeSuffix =
        apiContext.tokenSource && apiContext.tokenSource.purpose
          ? ' [' + apiContext.tokenSource.purpose + ']'
          : '';
      if (
        !apiTypeSuffix &&
        apiContext.tokenSource &&
        apiContext.tokenSource.source
      ) {
        apiTypeSuffix = ` [token source=${apiContext.tokenSource.source}]`;
      }
      debug(`${optionalMessage}${apiTypeSuffix}`);
    }
    const headers = {
      Authorization: authorizationHeaderValue,
    };
    if (apiContext.options.headers) {
      apiContext.headers = apiContext.options.headers;
      Object.assign(headers, apiContext.headers);
    }
    if (apiContext.etag) {
      headers['If-None-Match'] = apiContext.etag;
    }
    ++apiContext.cost.github.restApiCalls;
    const args = [];
    const apiMethod = apiContext.apiMethod;
    if (apiContext.fakeLink) {
      args.push(apiContext.fakeLink, headers);
    } else {
      const argOptions = Object.assign({}, apiContext.options);
      if (argOptions.octokitRequest) {
        args.push(argOptions.octokitRequest);
        delete argOptions.octokitRequest;
      }
      if (argOptions.additionalDifferentiationParameters) {
        delete argOptions.additionalDifferentiationParameters;
      }
      argOptions.headers = headers;
      args.push(argOptions);
    }
    const thisArgument = apiMethod.thisInstance || null;
    const response = await apiMethod.apply(thisArgument, args);
    return response;
  }

  processMetadataBeforeCall(apiContext: ApiContext, metadata: IRestMetadata) {
    if (
      metadata &&
      metadata.av &&
      apiContext.libraryContext.breakingChangeGitHubPackageVersion &&
      !semver.gte(
        metadata.av,
        apiContext.libraryContext.breakingChangeGitHubPackageVersion
      )
    ) {
      console.warn(
        `${apiContext.redisKey.metadata} was using ${metadata.av}, which is < to ${apiContext.libraryContext.breakingChangeGitHubPackageVersion}. This is a schema break, discarding cache.`
      );
      metadata = undefined;
    } else if (metadata && !metadata.av) {
      // Old version of metadata, no package version, which is required for all GitHub REST API metadata now
      metadata = undefined;
    }
    if (metadata && metadata.etag) {
      apiContext.etag = metadata.etag;
      apiContext.metadata = metadata;
    }
    return metadata;
  }

  withResponseUpdateMetadata(apiContext: ApiContext, response: IRestResponse) {
    return response;
  }

  optionalStripResponse(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestResponse {
    const clonedResponse = Object.assign({}, response);
    if (response.headers) {
      let clonedHeaders = StripGitHubEntity(
        GitHubResponseType.Headers,
        response.headers,
        'response.headers'
      );
      if (clonedHeaders) {
        clonedResponse.headers = clonedHeaders;
        if (debugShowStandardBehavior) {
          debugCacheOptimization('using stripped headers');
        }
      }
    }
    if (response.data) {
      let apiCall = apiContext.api as string;
      if ((apiContext as GitHubApiContext).pageAwareTypeInformation) {
        const pageAwareTypeInformation = (apiContext as GitHubApiContext)
          .pageAwareTypeInformation;
        if (pageAwareTypeInformation && pageAwareTypeInformation.methodName) {
          apiCall = pageAwareTypeInformation.methodName;
        }
      }
      const knownEntityType = entityData.apiToEntityType.get(apiCall);
      const knownResponseBodyType = entityData.apiToEntityResponseType.get(
        apiCall
      );
      if (!knownEntityType) {
        if (debugOutputUnregisteredEntityApis) {
          debugCacheOptimization(apiCall);
          debugCacheOptimization(JSON.stringify(response.data, undefined, 2));
        }
        debugCacheOptimization(
          `Cache Optimization WARNING: the API call ${apiCall} has no known entity response type, so data will not be optimized for caching`
        );
      } else if (
        Array.isArray(response.data) &&
        knownResponseBodyType !== ResponseBodyType.Array
      ) {
        if (debugOutputUnregisteredEntityApis) {
          debugCacheOptimization(apiCall);
          debugCacheOptimization(JSON.stringify(response.data, undefined, 2));
        }
        debugCacheOptimization(
          `Cache Optimization WARNING: the API call ${apiCall} is not registered to return an array, but it did.. NO optimization being performed.`
        );
      } else if (
        knownResponseBodyType === ResponseBodyType.Array &&
        Array.isArray(response.data)
      ) {
        let arrayClone = [];
        const remainingKeys = new Set(
          Object.getOwnPropertyNames(response.data)
        );
        remainingKeys.delete('length');
        for (let i = 0; i < response.data.length; i++) {
          const entity = response.data[i];
          const entityClone = StripGitHubEntity(
            knownEntityType,
            entity,
            'response.data[' + i + ']'
          );
          arrayClone.push(entityClone ? entityClone : entity);
          remainingKeys.delete(i.toString());
        }
        if (remainingKeys.size) {
          const names = Array.from(remainingKeys.keys()).join(', ');
          throw new Error(
            `This entity simplification function assumes that there are no additional keys appended to the response data array. The following keys remain: ${names}`
          );
        }
        if (arrayClone.length) {
          clonedResponse.data = arrayClone;
          if (debugShowStandardBehavior) {
            debugCacheOptimization(
              `using reduced response array body for ${arrayClone.length} entities`
            );
          }
        }
      } else if (knownResponseBodyType === ResponseBodyType.Array) {
        if (debugOutputUnregisteredEntityApis) {
          debugCacheOptimization(apiCall);
          debugCacheOptimization(JSON.stringify(response.data, undefined, 2));
        }
        debugCacheOptimization(
          `Cache Optimization WARNING: the API call ${apiCall} is registered to return an array, but it did not.. NO optimization being performed.`
        );
      } else {
        const strippedBody = StripGitHubEntity(
          knownEntityType,
          response.data,
          'response.data'
        );
        if (strippedBody) {
          clonedResponse.data = strippedBody;
          if (debugShowStandardBehavior) {
            debugCacheOptimization(
              `reduced response body for entity ${knownEntityType} used`
            );
          }
        } else {
          if (debugShowStandardBehavior) {
            debugCacheOptimization(
              `nothing could be reduced from the response.data for ${knownEntityType}`
            );
          }
        }
      }
    }
    return clonedResponse;
  }

  reduceMetadataToCacheFromResponse(
    apiContext: ApiContext,
    response: IRestResponse
  ): any {
    const headers = response ? response.headers : null;
    if (headers?.etag) {
      let reduced: IReducedGitHubMetadata = {
        etag: headers.etag,
        av: appVersion,
      };
      if (headers.link) {
        reduced.link = headers.link;
      }
      // Updated for 2021: parse last-modified to use as a more accurate 'changed' value
      const lastModifiedTime = headers?.['last-modified'];
      let updated = lastModifiedTime ? new Date(lastModifiedTime) : null;
      if (!updated) {
        const calledTime = apiContext.calledTime
          ? apiContext.calledTime
          : new Date();
        updated = calledTime;
      }
      reduced.updated = updated.toISOString();
      return reduced;
    }
    return headers;
  }

  withResponseShouldCacheBeServed(
    apiContext: ApiContext,
    response: IRestResponse
  ): boolean | IShouldServeCache {
    if (response === undefined) {
      throw new Error('The response was undefined and unable to process.');
    }
    if (!response.headers) {
      throw new Error(
        'As of Octokit 15.8.0, responses must have headers on the response'
      );
    }
    const headers = response.headers;
    let retryAfter = headers['retry-after'];
    if (retryAfter) {
      debug(`Retry-After header was present: ${retryAfter}`);
    }
    const rateLimitRemaining = headers['x-ratelimit-remaining'];
    if (rateLimitRemaining) {
      apiContext.cost.github.remainingApiTokens = rateLimitRemaining;
    }
    const { status } = response;
    let cacheOk = false;
    const displayInfo = apiContext.redisKey ? apiContext.redisKey.root : '';
    if (status === 304 || response.notModified) {
      let appPurposeSuffix =
        apiContext.tokenSource && apiContext.tokenSource.purpose
          ? ` [${apiContext.tokenSource.purpose}]`
          : '';
      if (
        apiContext.tokenSource &&
        !apiContext.tokenSource.purpose &&
        apiContext.tokenSource.source
      ) {
        appPurposeSuffix = ` [token source=${apiContext.tokenSource.source}]`;
      }
      debug(`304:               ${displayInfo} ${appPurposeSuffix}`);
      ++apiContext.cost.github.cacheHits;
      cacheOk = true;
    } else if (status !== undefined && (status < 200 || status >= 300)) {
      // The underlying library I believe actually processes these conditions as errors anyway
      throw new Error(
        `Response code of ${status} is not currently supported in this system.`
      );
    }
    return cacheOk;
  }

  getResponseMetadata(
    apiContext: ApiContext,
    response: IRestResponse
  ): IRestMetadata {
    const md: IRestMetadata = {
      headers: response.headers,
      status: response.status,
    };
    return md;
  }

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
        debug(`NO_METADATA:       ${apiContext.redisKey.metadata} [empty]`);
      } else {
        debug(
          `NO_CHANGE:         ${apiContext.redisKey.metadata} ${
            metadata.etag ? '[etag: ' + metadata.etag + ']' : ''
          }`
        );
      }
    }
    return shouldServeCache;
  }
}

export class GitHubApiContext extends ApiContext {
  private _apiMethod: any;
  private _redisKeys: IApiContextRedisKeys;
  private _cacheValues: IApiContextCacheValues;
  private _token: string | IGetAuthorizationHeader | IAuthorizationHeaderValue;

  public fakeLink?: IGitHubLink;

  public headers?: any;

  public pageAwareTypeInformation?: any; // used for extended cache options

  constructor(api: any, options: any) {
    super(api, options);

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
      longtermMetadata: longtermMetadataMinutes,
      longtermResponse: longtermResponseMinutes,
      acceleratedExpiration: acceleratedExpirationMinutes,
    };
  }

  get token(): string | IGetAuthorizationHeader | IAuthorizationHeaderValue {
    return this._token;
  }

  get apiMethod(): any {
    return this._apiMethod;
  }

  get apiTypePrefix(): string {
    return 'github#';
  }

  get redisKey(): IApiContextRedisKeys {
    return this._redisKeys;
  }

  get cacheValues(): IApiContextCacheValues {
    return this._cacheValues;
  }

  get contextType(): ApiContextType {
    return ApiContextType.GitHubRestApi;
  }

  attachToApiImplementation(implementationLibrary: any) {
    if (this._apiMethod) {
      // NOTE: this restriction was not in place in the original implementation
      // and is probably not needed
      throw new Error('API has already been attached to');
    }
    const method = IntelligentGitHubEngine.findLibraryMethod(
      implementationLibrary,
      this.api
    );
    method['thisInstance'] = implementationLibrary; // // HACK, is there a better way?
    this._apiMethod = method;
  }

  setLibraryContext(libraryContext: any) {
    this.libraryContext = libraryContext;
  }

  overrideToken(
    token: string | IGetAuthorizationHeader | IAuthorizationHeaderValue
  ) {
    if (token && token['value']) {
      const asPair = token as IAuthorizationHeaderValue;
      this._token = asPair.value;
      this.tokenSource = asPair;
    } else if (typeof token === 'string') {
      this._token = token as string;
    } else {
      this._token = token;
    }
  }

  overrideApiMethod(method: any) {
    this._apiMethod = method;
  }
}

function prepareApiContextForGithub(
  apiContext: GitHubApiContext,
  github: any
): GitHubApiContext {
  if (!apiContext.apiMethod) {
    apiContext.attachToApiImplementation(github);
  }
  return apiContext;
}

export function createFullContext(
  api: any,
  options: any,
  github: any,
  libraryContext: any
): GitHubApiContext {
  const apiContext = prepareApiContextForGithub(
    createApiContextForGithub(api, options),
    github
  );
  apiContext.setLibraryContext(libraryContext);
  return apiContext;
}

function createApiContextForGithub(api: any, options: any): GitHubApiContext {
  const apiContext = new GitHubApiContext(api, options);
  return apiContext;
}

export function StripGitHubEntity(
  entityType: GitHubResponseType,
  incomingEntity: any,
  keyOrName: string
): any | null {
  let entityClone = null;
  if (!incomingEntity || typeof incomingEntity !== 'object') {
    return; // no change
  }
  const keepers = entityData.entityPropertiesToKeep.get(entityType) || emptySet;
  const droppers =
    entityData.entityPropertiesToDrop.get(entityType) || emptySet;
  const objects = entityData.entityPropertiesSubsets.get(entityType);
  const entityKeys = Object.getOwnPropertyNames(incomingEntity);
  for (let j = 0; j < entityKeys.length; j++) {
    const fieldName = entityKeys[j];
    const fieldObjectType = objects ? objects.get(fieldName) : null;
    if (keepers.has(fieldName)) {
      // Safe known field to keep
    } else if (droppers.has(fieldName)) {
      if (!entityClone) {
        entityClone = Object.assign({}, incomingEntity);
        if (debugShowStandardBehavior) {
          debugCacheOptimization(
            `stripping from response ${keyOrName} of type ${entityType}: (clone)`
          );
        }
      }
      delete entityClone[fieldName];
      if (debugShowStandardBehavior) {
        debugCacheOptimization(
          `field strip: ${fieldName} from ${keyOrName} entity (${entityType})`
        );
      }
    } else if (fieldObjectType) {
      // this property itself is a sub-object that might want to get parsed
      if (!entityClone) {
        entityClone = Object.assign({}, incomingEntity);
        if (debugShowStandardBehavior) {
          debugCacheOptimization(
            `stripping from response ${keyOrName} of type ${entityType}: (clone)`
          );
        }
      }
      const newSubObject = StripGitHubEntity(
        fieldObjectType,
        entityClone[fieldName],
        `${keyOrName}.${fieldName}`
      );
      if (newSubObject) {
        entityClone[fieldName] = newSubObject;
        if (debugShowStandardBehavior) {
          debugCacheOptimization(
            `replacing ${keyOrName}.${fieldName} sub-entity with a subset object (${entityType})`
          );
        }
      } else {
        if (debugShowStandardBehavior) {
          debugCacheOptimization(
            `no subset required for sub-entity ${keyOrName}.${fieldName} (${entityType})`
          );
        }
      }
    } else {
      debugCacheOptimization(
        `*NOT* stripping ${keyOrName}.${fieldName} (type ${entityType}) (not a registered field)`
      );
    }
  }
  return entityClone;
}
