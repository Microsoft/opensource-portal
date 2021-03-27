//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Octokit } from '@octokit/rest';
const githubPackage = require('@octokit/rest/package.json');

import * as restApi from './restApi';
import { flattenData } from './core';
import { CompositeIntelligentEngine } from './composite';
import { RestCollections } from './collections';
import { CrossOrganizationCollator } from './crossOrganization';
import { LinkMethods } from './links';
import RedisHelper from '../caching/redis';
import {
  IGetAuthorizationHeader,
  IAuthorizationHeaderValue,
} from '../../transitional';
import { ILinkProvider } from '../linkProviders';
import { ICacheHelper } from '../caching';

export enum CacheMode {
  ValidateCache = 'ValidateCache',
  BackgroundRefresh = 'BackgroundRefresh',
}

export interface IGitHubPostFunction {
  (
    awaitToken: IGetAuthorizationHeader,
    api: string,
    parameters: any
  ): Promise<any>;
}

// With the introduction of a breaking change in the underlying schema, any cache objects
// which are related to the GitHub library and have a SemVer equal to or less than this
// value will be discarded. The lack of a 'av' property (app version, originally) will
// also trigger a discard.
const breakingChangeGitHubPackageVersion = '6.0.0';

interface IRestLibraryOptions {
  config: any;
  cacheProvider: ICacheHelper;
  github?: Octokit;
  baseUrl?: string;
}

export class RestLibrary {
  public cacheProvider: ICacheHelper;
  private github: Octokit;

  private _collections: RestCollections;
  private _links: LinkMethods;
  private _crossOrganization: CrossOrganizationCollator;
  private githubEngine?: restApi.IntelligentGitHubEngine;

  defaultPageSize: number;

  public breakingChangeGitHubPackageVersion: string;
  public compositeEngine?: CompositeIntelligentEngine;

  constructor(options: IRestLibraryOptions) {
    const cacheProvider = options.cacheProvider;
    if (!cacheProvider) {
      throw new Error(
        'No Redis instance provided to the GitHub library context constructor.'
      );
    }
    this.cacheProvider = cacheProvider;

    let config = options.config;
    if (!config) {
      throw new Error(
        'No runtime configuration instance provided to the library context constructor'
      );
    }

    const nodeGithubVersion = `${githubPackage.name}/${githubPackage.version}`;
    let userAgent = nodeGithubVersion;
    if (
      config &&
      config.github &&
      config.github.library &&
      config.github.library.userAgent
    ) {
      userAgent = config.github.library.userAgent;
    }
    let github = options.github;
    if (!github) {
      github = new Octokit({
        userAgent,
        baseUrl: options.baseUrl,
      });
    }
    this.github = github;

    (this.defaultPageSize =
      config &&
      config.github &&
      config.github.api &&
      config.github.api.defaultPageSize
        ? config.github.api.defaultPageSize
        : 100),
      (this.breakingChangeGitHubPackageVersion = breakingChangeGitHubPackageVersion);

    this.githubEngine = new restApi.IntelligentGitHubEngine();
    this.compositeEngine = new CompositeIntelligentEngine();

    this.hasNextPage = hasNextPage.bind(this);

    this.call = this.call.bind(this);
    this.post = this.post.bind(this);
    this.request = this.request.bind(this);
  }

  get collections(): RestCollections {
    if (!this._collections) {
      this._collections = new RestCollections(this, this.call);
    }
    return this._collections;
  }

  get links(): LinkMethods {
    if (!this._links) {
      this._links = new LinkMethods(this);
    }
    return this._links;
  }

  get crossOrganization(): CrossOrganizationCollator {
    if (!this._crossOrganization) {
      this._crossOrganization = new CrossOrganizationCollator(
        this,
        this.collections
      );
    }
    return this._crossOrganization;
  }

  hasNextPage?: (any) => boolean;

  private async resolveAuthorizationHeader(
    authorizationHeader:
      | IGetAuthorizationHeader
      | IAuthorizationHeaderValue
      | string
  ): Promise<string | IAuthorizationHeaderValue> {
    let authorizationValue = null;
    try {
      if (typeof authorizationHeader === 'string') {
        authorizationValue = authorizationHeader as string;
      } else if (typeof authorizationHeader === 'function') {
        let asFunc = authorizationHeader as IGetAuthorizationHeader;
        let resolved = asFunc.call(null) as Promise<
          IAuthorizationHeaderValue | string
        >;
        authorizationValue = await resolved;
        if (typeof resolved === 'function') {
          asFunc = resolved as IGetAuthorizationHeader;
          resolved = asFunc.call(null) as Promise<
            IAuthorizationHeaderValue | string
          >;
          authorizationValue = await resolved;
        }
      } else if (authorizationHeader && authorizationHeader['value']) {
        authorizationValue = authorizationHeader as IAuthorizationHeaderValue;
      } else {
        throw new Error('Invalid resolveAuthorizationHeader');
      }
    } catch (getTokenError) {
      console.dir(getTokenError);
      throw getTokenError;
    }
    return authorizationValue;
  }

  async call(
    awaitToken: IGetAuthorizationHeader | IAuthorizationHeaderValue | string,
    api: string,
    options,
    cacheOptions = null
  ): Promise<any> {
    cacheOptions = cacheOptions || {};
    let massageData = (data) => flattenData(data);
    if (options.allowEmptyResponse) {
      delete options.allowEmptyResponse;
      massageData = (data) => data;
    }
    const apiContext = restApi.createFullContext(
      api,
      options,
      this.github,
      this
    );
    // CONSIDER: technically, callApi can wait to resolve the token by passing it into the context as-is
    apiContext.overrideToken(await this.resolveAuthorizationHeader(awaitToken));
    if (cacheOptions.maxAgeSeconds) {
      apiContext.maxAgeSeconds = cacheOptions.maxAgeSeconds;
    }
    if (cacheOptions.backgroundRefresh !== undefined) {
      apiContext.backgroundRefresh = cacheOptions.backgroundRefresh;
    }
    const data = await this.githubEngine.execute(apiContext);
    const result = massageData(data);
    return result;
  }

  request(token, restEndpoint, parameters: any, cacheOptions): Promise<any> {
    parameters = parameters || {};
    parameters['octokitRequest'] = restEndpoint;
    return this.call(token, 'request', parameters, cacheOptions);
  }

  requestAsPost(token, restEndpoint, parameters: any): Promise<any> {
    parameters = parameters || {};
    parameters['octokitRequest'] = restEndpoint;
    return this.post(token, 'request', parameters);
  }

  graphql(token, query, parameters: any): Promise<any> {
    parameters = parameters || {};
    parameters['octokitQuery'] = query;
    return this.post(token, 'graphql', parameters);
  }

  async post(
    awaitToken: IGetAuthorizationHeader | string,
    api: string,
    options: any
  ): Promise<any> {
    const method = restApi.IntelligentGitHubEngine.findLibraryMethod(
      this.github,
      api
    );
    if (!options.headers) {
      options.headers = {};
    }
    const noDataMassage = (data) => data;
    let massageData = (data) => flattenData(data);
    if (options.allowEmptyResponse) {
      delete options.allowEmptyResponse;
      massageData = noDataMassage;
    }
    if (!options.headers.authorization) {
      const value = await this.resolveAuthorizationHeader(awaitToken);
      options.headers.authorization =
        typeof value === 'string'
          ? (value as string)
          : (value as IAuthorizationHeaderValue).value;
    }
    try {
      let value = null;
      if (api === 'request' && options.octokitRequest) {
        const endpoint = options.octokitRequest;
        delete options.octokitRequest;
        value = (await method.call(
          this.github,
          endpoint,
          options
        )) as Promise<any>;
      } else if (api === 'graphql') {
        massageData = noDataMassage;
        const query = options.octokitQuery;
        delete options.octokitQuery;
        value = (await method.call(
          this.github,
          query,
          options
        )) as Promise<any>;
      } else {
        value = (await method.call(this.github, options)) as Promise<any>;
      }
      const finalized = massageData(value);
      return finalized;
    } catch (error) {
      console.log(error.message);
      if (error.status) {
        console.log(`Status: ${error.status}`);
      }
      if (error.headers && error.headers['x-ratelimit-remaining']) {
        console.log(
          `Rate limit remaining: ${error.headers['x-ratelimit-remaining']}`
        );
      }
      if (error.request) {
        console.dir(error.request);
      }
      throw error;
    }
  }
}

// follows: deprecated functions that parse links out of the response headers

function getPageLinks(link: any): any {
  link = link.link || link.headers.link || '';
  const links = {};
  // link format:
  // '<https://api.github.com/users/aseemk/followers?page=2>; rel="next", <https://api.github.com/users/aseemk/followers?page=2>; rel="last"'
  link.replace(/<([^>]*)>;\s*rel="([\w]*)"/g, (m, uri, type) => {
    links[type] = uri;
  });
  return links;
}

function hasNextPage(link): string {
  return getPageLinks(link).next;
}
