//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IRestResponse, IIntelligentCacheResponseArray, flattenData } from './core';
import { CompositeApiContext } from './composite';
import { RestLibrary } from '.';
import { RestCollections } from './collections';
import { ICacheOptions, IGetOrganizationMembersOptions, IPagedCrossOrganizationCacheOptions, IPurposefulGetAuthorizationHeader, ITeamMembershipOptions } from '../../interfaces';
import { AppPurpose } from '../../github';

interface IOrganizationsResponse extends IRestResponse {
  orgs?: any;
}

interface ICrossOrganizationDataResponse extends IRestResponse {
//  data?: any;
}

interface ILocalOptionsParameters {
  per_page: number;
  id?: string;
  team_id?: string;
  owner?: string;
  repo?: string;
}

export class CrossOrganizationCollator {
  public libraryContext: RestLibrary;
  public collectionsClient: RestCollections;

  constructor(libraryContext: RestLibrary, collectionsClient: RestCollections) {
    this.libraryContext = libraryContext;
    this.collectionsClient = collectionsClient;
  }

  async orgMembers(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options: IGetOrganizationMembersOptions, cacheOptions: ICacheOptions): Promise<any> {
    options['apiTypePrefix'] = 'github.x#';
    const data = await this.getCrossOrganizationMethod(orgsAndTokens, 'orgMembers', 'getOrgMembers', options, cacheOptions);
    return flattenData(data);
  }

  async teams(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options, cacheOptions) {
    const allTeams = await this.getAllTeams(orgsAndTokens, options, cacheOptions);
    return flattenData(allTeams);
  }

  async teamMembers(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options: ITeamMembershipOptions, cacheOptions: ICacheOptions): Promise<any> {
    options['apiTypePrefix'] = 'github.x#';
    const capturedThis = this;
    const generalizedData = await this.generalizedCollectionMethod(
      orgsAndTokens,
      'teamMembers',
      capturedThis.crossOrganizationCollection(capturedThis, orgsAndTokens, options, cacheOptions, 'team', capturedThis.getAllTeams, 'getTeamMembers', 'members', true),
      options,
      cacheOptions);
    return flattenData(generalizedData);
  }

  async repos(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options, cacheOptions: ICacheOptions): Promise<any> {
    const allRepos = await this.getAllRepos(orgsAndTokens, options, cacheOptions);
    return flattenData(allRepos);
  }

  async repoCollaborators(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options, cacheOptions: ICacheOptions): Promise<any> {
    options.apiTypePrefix = 'github.x#';
    const capturedThis =  this;
    const generalizedData = await this.generalizedCollectionMethod(
      orgsAndTokens,
      'repoCollaborators',
      capturedThis.crossOrganizationCollection(capturedThis, orgsAndTokens, options, cacheOptions, 'repo', capturedThis.getAllRepos, 'getRepoCollaborators', 'collaborators', true),
      options,
      cacheOptions);
    return flattenData(generalizedData);
  }

  async repoTeams(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options, cacheOptions: ICacheOptions): Promise<any> {
    options.apiTypePrefix = 'github.x#';
    const capturedThis =  this;
    const generalizedData = await this.generalizedCollectionMethod(
      orgsAndTokens,
      'repoTeams',
      capturedThis.crossOrganizationCollection(capturedThis, orgsAndTokens, options, cacheOptions, 'repo', capturedThis.getAllRepos, 'getRepoTeams', 'teams', true),
      options,
      cacheOptions);
    return flattenData(generalizedData);
  }

  private async generalizedCollectionMethod(token, apiName, method, options, cacheOptions?: ICacheOptions): Promise<any> { // IIntelligentEngineResponse
    cacheOptions = cacheOptions || {};
    const apiContext = new CompositeApiContext(apiName, method, options);
    apiContext.maxAgeSeconds = cacheOptions.maxAgeSeconds || 600;
    apiContext.overrideToken(token);
    apiContext.libraryContext = this.libraryContext;
    if (cacheOptions.backgroundRefresh) {
      apiContext.backgroundRefresh = true;
    }
    const res = await this.libraryContext.compositeEngine.execute(apiContext);
    return res;
  }

  private async getCrossOrganizationMethod(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, apiName: string, methodName: string, options, cacheOptions: IPagedCrossOrganizationCacheOptions): Promise<any> {
    const method = this.collectionsClient[methodName];
    if (!method) {
      throw new Error(`No method called ${method} defined in the collections client.`);
    }
    const capturedThis = this;
    const crossOrgMethod = async function actualCrossOrgMethod(): Promise<any> {
      const values: IOrganizationsResponse = {
        headers: {},
        orgs: {},
        data: undefined,
      };
      const organizations = Array.from(orgsAndTokens.keys());
      for (let orgName of organizations) {
        const token = orgsAndTokens.get(orgName);
        const localOptions = Object.assign({}, options);
        localOptions.org = orgName;
        if (!localOptions.per_page) {
          localOptions.per_page = capturedThis.libraryContext.defaultPageSize;
        }
        const localCacheOptions = Object.assign({}, cacheOptions);
        if (localCacheOptions.individualMaxAgeSeconds) {
          localCacheOptions.maxAgeSeconds = localCacheOptions.individualMaxAgeSeconds;
        }
        try {
          const orgValues = await method.call(capturedThis.collectionsClient, token.bind(null, AppPurpose.Data), localOptions, localCacheOptions);
          if (!orgValues) {
            throw new Error('No result');
          }
          if (orgValues && orgValues.data) {
            console.warn(`${apiName} ${methodName} result has data that is being used instead of the parent object`);
            values.orgs[orgName] = orgValues.data;
          } else {
            values.orgs[orgName] = orgValues;
          }
        } catch (orgError) {
          throw orgError;
        }
      }
      const dataObject = {
        data: values,
        headers: values.headers,
      };
      delete values.headers;
      return dataObject;
    };
    return await this.generalizedCollectionMethod(orgsAndTokens, apiName, crossOrgMethod, options, cacheOptions);
  }

  private crossOrganizationCollection(capturedThis: CrossOrganizationCollator, orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options, cacheOptions: IPagedCrossOrganizationCacheOptions, innerKeyType, outerFunction, collectionMethodName: string, collectionKey, optionalSetOrganizationLogin) {
    return async (): Promise<any> => {
      const entities = [] as IIntelligentCacheResponseArray;
      entities.headers = {};
      let data = null;
      try {
        data = await outerFunction.call(capturedThis, orgsAndTokens, {}, cacheOptions);
      } catch (outerError) {
        throw outerError;
      }
      let entitiesByOrg = null;
      if (data && !data.data) {
        throw new Error('crossOrganizationCollection inner outerFunction returned an entity but no entity.data property was present');
      } else if (data && data.data) {
        entitiesByOrg = data.data;
      }
      const localCacheOptions = Object.assign({}, cacheOptions);
      if (localCacheOptions.individualMaxAgeSeconds) {
        localCacheOptions.maxAgeSeconds = localCacheOptions.individualMaxAgeSeconds;
      }
      entities.headers = {};
      const orgNames = Object.getOwnPropertyNames(entitiesByOrg.orgs);
      for (let i = 0; i < orgNames.length; i++) {
        const orgName = orgNames[i];
        const orgEntities = entitiesByOrg.orgs[orgName];
        for (const orgEntity of orgEntities) {
          const cloneTarget = optionalSetOrganizationLogin ? {
            organization: {
              login: orgName,
            }
          } : {};
          const entityClone = Object.assign(cloneTarget, orgEntity);
          const localOptionsTarget: ILocalOptionsParameters = {
            per_page: capturedThis.libraryContext.defaultPageSize,
          };
          switch (innerKeyType) {
            case 'team':
              localOptionsTarget.team_id = orgEntity.id;
              break;
            case 'repo':
              localOptionsTarget.owner = orgName;
              localOptionsTarget.repo = orgEntity.name;
              break;
            default:
              throw new Error(`Unsupported inner key type ${innerKeyType}`);
          }
          const localOptions = Object.assign(localOptionsTarget, options);
          delete localOptions.maxAgeSeconds;
          delete localOptions.backgroundRefresh;
          const token = orgsAndTokens.get(orgName.toLowerCase()) as IPurposefulGetAuthorizationHeader;
          if (!token) {
            throw new Error(`No token available for the organization ${orgName}`);
          }
          let innerEntities = null;
          let collectionsError = null;
          try {
            innerEntities = await this.collectionsClient[collectionMethodName](token(AppPurpose.Data), localOptions, localCacheOptions);
          } catch (error) {
            // This is a silent error for now, because there
            // are valid scenarios, i.e. team deletion, to consider.
            // In the future, get smarter here.
            collectionsError = error;
          }
          if (!collectionsError && innerEntities && innerEntities.data) {
            collectionsError = new Error(`innerEntities.data set from the ${collectionMethodName} collection method call`);
          }
          if (!collectionsError) {
            entityClone[collectionKey] = innerEntities;
            entities.push(entityClone);
          }
        }
      }
      const projectedToDataEntity: ICrossOrganizationDataResponse = {
        data: entities,
      };
      if (entities.cost) {
        projectedToDataEntity.cost = entities.cost;
        delete entities.cost;
      }
      if (entities.headers) {
        projectedToDataEntity.headers = entities.headers;
        delete entities.headers;
      }
      return projectedToDataEntity;
    }
  }

  private async getAllTeams(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options, cacheOptions: IPagedCrossOrganizationCacheOptions): Promise<any> {
    options.apiTypePrefix = 'github.x#';
    const data = await this.getCrossOrganizationMethod(orgsAndTokens, 'teams', 'getOrgTeams', options, cacheOptions);
    return data;
  }

  private async getAllRepos(orgsAndTokens: Map<string, IPurposefulGetAuthorizationHeader>, options, cacheOptions: IPagedCrossOrganizationCacheOptions): Promise<any> {
    options.apiTypePrefix = 'github.x#';
    const data = await this.getCrossOrganizationMethod(orgsAndTokens, 'repos', 'getOrgRepos', options, cacheOptions);
    return data;
  }
}
