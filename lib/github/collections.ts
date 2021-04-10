//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import _ from 'lodash';

const debug = require('debug')('restapi');

import cost from './cost';
import { IRestResponse, flattenData } from './core';
import { CompositeApiContext, CompositeIntelligentEngine } from './composite';
import { Collaborator } from '../../business/collaborator';
import { Team } from '../../business/team';
import { IPagedCacheOptions, IGetAuthorizationHeader, IDictionary } from '../../interfaces';
import { RestLibrary } from '.';
import { sleep } from '../../utils';
import GitHubApplication from '../../business/application';
import { RepositoryPrimaryProperties } from '../../business/primaryProperties';

export interface IGetAppInstallationsParameters {
  app_id: string;
}

export enum GitHubPullRequestState {
  Open = 'open',
  Closed = 'closed',
  All = 'all',
}

export enum GitHubIssueState {
  Open = 'open',
  Closed = 'closed',
  All = 'all',
}

export enum GitHubPullRequestSort {
  Created = 'created',
  Updated = 'updated',
  Popularity = 'popularity', // comment count
  LongRunning = 'long-running', // age, filtering by pulls updated in the last month
}

export enum GitHubSortDirection {
  Ascending = 'asc',
  Descending = 'desc',
}

export interface IListPullsParameters {
  owner: string;
  repo: string;
  state?: GitHubPullRequestState;
  head?: string;
  base?: string;
  sort?: GitHubPullRequestSort;
  direction?: GitHubSortDirection;
}

const branchDetailsToCopy = [
  'name',
  'commit',
  'protected',
];
const repoDetailsToCopy = RepositoryPrimaryProperties;
const teamDetailsToCopy = Team.PrimaryProperties;
const memberDetailsToCopy = Collaborator.PrimaryProperties;
const appInstallDetailsToCopy = GitHubApplication.PrimaryInstallationProperties;
const contributorsDetailsToCopy = [
  ...Collaborator.PrimaryProperties,
  'contributions',
];

const teamPermissionsToCopy = [
  'id',
  'name',
  'slug',
  'description',
  'members_count',
  'repos_count',
  'privacy',
  'permission',
];
const teamRepoPermissionsToCopy = [
  'id',
  'name',
  'full_name',
  'description',
  'private',
  'fork',
  'permissions',
];
const pullDetailsToCopy = [
  'id',
  'number',
  'state',
  'locked',
  'title',
  // user
  'body',
  // labels
  // milestone
  // active_lock_reason
  'created_at',
  'updated_at',
  'closed_at',
  'merged_at',
  'merge_commit_sha',
  'assignee',
  // assignees
  // requested_reviewers
  // requested_teams
  'head', // PERF: large user of list storage
  'base', // PERF: large user of list storage
  'author_association',
  'draft',
];

interface IRequestWithData {
  data: unknown;
  requests: IPageRequest[];
}

interface IPageRequest {
  cost: unknown;
  headers: IDictionary<string>;
  status?: number;
}

export class RestCollections {
  private libraryContext: RestLibrary;
  private githubCall: unknown;

  constructor(libraryContext: RestLibrary, githubCall: unknown) {
    this.libraryContext = libraryContext;
    this.githubCall = githubCall;
  }

  getOrgRepos(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('orgRepos', 'repos.listForOrg', repoDetailsToCopy, token, options, cacheOptions);
  }

  getOrgTeams(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('orgTeams', 'teams.list', teamDetailsToCopy, token, options, cacheOptions);
  }

  getTeamChildTeams(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('teamChildTeams', 'teams.listChildInOrg', teamDetailsToCopy, token, options, cacheOptions);
  }

  getUserActivity(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('userActivity', 'activity.listEventsForAuthenticatedUser', null /*activityDetailsToCopy*/, token, options, cacheOptions);
  }

  getOrgMembers(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('orgMembers', 'orgs.listMembers', memberDetailsToCopy, token, options, cacheOptions);
  }

  getAppInstallations(token: string | IGetAuthorizationHeader, parameters: IGetAppInstallationsParameters, cacheOptions: IPagedCacheOptions): Promise<any> {
    if (!parameters.app_id) {
      throw new Error('parameters.app_id required');
    }
    const projectedOptions = {
      additionalDifferentiationParameters: parameters,
    };
    return this.generalizedCollectionWithFilter(`appInstallations`, 'apps.listInstallations', appInstallDetailsToCopy, token, projectedOptions, cacheOptions);
  }

  getRepoIssues(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any[]> {
    return this.generalizedCollectionWithFilter('repoIssues', 'issues.listForRepo', null, token, options, cacheOptions);
  }

  getRepoTeams(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('repoTeamPermissions', 'repos.listTeams', teamPermissionsToCopy, token, options, cacheOptions);
  }

  getRepoContributors(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('repoListContributors', 'repos.listContributors', contributorsDetailsToCopy, token, options, cacheOptions);
  }

  getRepoCollaborators(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('repoCollaborators', 'repos.listCollaborators', memberDetailsToCopy, token, options, cacheOptions);
  }

  getRepoBranches(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('repoBranches', 'repos.listBranches', branchDetailsToCopy, token, options, cacheOptions);
  }

  getRepoPullRequests(token: string | IGetAuthorizationHeader, options: IListPullsParameters, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('repoPullRequests', 'pulls.list', pullDetailsToCopy, token, options, cacheOptions);
  }

  getTeamMembers(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('teamMembers', 'teams.listMembersInOrg', memberDetailsToCopy, token, options, cacheOptions);
  }

  getTeamRepos(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    return this.generalizedCollectionWithFilter('teamRepos', 'teams.listReposInOrg', teamRepoPermissionsToCopy, token, options, cacheOptions);
  }

  private async getGithubCollection(token: string | IGetAuthorizationHeader, methodName, options, cacheOptions: IPagedCacheOptions): Promise<IRequestWithData> {
    const hasNextPage = this.libraryContext.hasNextPage;
    const githubCall = this.githubCall;
    let done = false;
    let results = [];
    let recentResult = null;
    let requests = [];
    let pages = 0;
    let currentPage = 0;
    const pageLimit = options.pageLimit || cacheOptions['pageLimit'] || Number.MAX_VALUE;
    const pageRequestDelay = cacheOptions.pageRequestDelay || null;
    while (!done) {
      const method = githubCall;
      const args = [];
      const currentToken = typeof(token) === 'string' ? token : await token();
      args.push(currentToken);
      const clonedOptions = Object.assign({}, options);
      if (++currentPage > 1) {
        clonedOptions.page = currentPage;
      }
      args.push(methodName, clonedOptions);
      let error = null;
      let result = null;
      try {
        result = await (method as any).apply(null, args);
        recentResult = result;
        if (result) {
          ++pages;
          if (Array.isArray(result)) {
            results = results.concat(result);
          }
          requests.push({
            cost: result.cost,
            headers: result.headers,
          });
        }
        try {
          done = pages >= pageLimit || !hasNextPage(result);
        } catch (nextPageError) {
          error = nextPageError;
          done = true;
        }
      } catch (iterationError) {
        done = true;
        error = iterationError;
      }
      if (!done && !error && result.headers && result.headers['retry-after']) { // actual retry headers win
        const delaySeconds = result.headers['retry-after'];
        debug(`Retry-After header was present. Delaying before next page ${delaySeconds}s.`);
        await sleep(delaySeconds * 1000);
      } else if (pageRequestDelay) {
        const to = typeof(pageRequestDelay);
        let evaluatedTime = 0;
        if (to === 'number') {
          evaluatedTime = pageRequestDelay as number;
        } else if (to === 'function') {
          evaluatedTime = (pageRequestDelay as unknown as any)();
        } else {
          throw new Error(`Unsupported pageRequestDelay type: ${to}`);
        }
        await sleep(evaluatedTime);
      }
      if (error) {
        throw error;
      }
    }
    const data = {
      data: results,
    };
    return { data, requests };
  }

  private async getFilteredGithubCollection(token: string | IGetAuthorizationHeader, methodName, options, cacheOptions: IPagedCacheOptions, propertiesToKeep): Promise<IRequestWithData> {
    const keepAll = !propertiesToKeep;
    try {
      // IRequestWithData
      const getCollectionResponse = await this.getGithubCollection(token, methodName, options, cacheOptions);
      if (!getCollectionResponse) {
        throw new Error('No response');
      }
      const root = getCollectionResponse.data as any;
      if (!root) {
        throw new Error('No object, no data');
      }
      if (!root.data) {
        throw new Error('The resulting object did not contain a data property');
      }
      const requests = getCollectionResponse.requests;
      const results = root.data;
      const repos = [];
      for (let i = 0; i < results.length; i++) {
        const doNotModify = results[i];
        if (doNotModify) {
          const r = {};
          _.forOwn(doNotModify, (value, key) => {
            if (keepAll || propertiesToKeep.indexOf(key) >= 0) {
              r[key] = value;
            }
          });
          repos.push(r);
        }
      }
      const filteredData = {
        data: repos,
      };
      return {
        data: filteredData,
        requests,
      }
    } catch (error) {
      throw error;
    }
  }

  private async getFilteredGithubCollectionWithMetadataAnalysis(token: string | IGetAuthorizationHeader, methodName, options, cacheOptions: IPagedCacheOptions, propertiesToKeep): Promise<IRestResponse> {
    const collectionResults = await this.getFilteredGithubCollection(token, methodName, options, cacheOptions, propertiesToKeep);
    const results = collectionResults.data as IRestResponse;
    const requests = collectionResults.requests;
    const pages = [];
    let dirty = false;
    let dirtyModified = [];
    let compositeCost = cost.create();
    for (let i = 0; i < requests.length; i++) {
      if (requests[i] && requests[i].headers && requests[i].headers.etag) {
        pages.push(requests[i].headers.etag);
      } else {
        throw new Error('Invalid set of responses for pages');
      }
      if (requests[i] && requests[i].status && requests[i].status !== 304) {
        dirty = true;
        let lastModified = requests[i].headers['last-modified'];
        if (lastModified) {
          dirtyModified.push(lastModified);
        }
      }
      if (requests[i] && requests[i].cost) {
        cost.add(compositeCost, requests[i].cost);
      }
    }
    if (dirtyModified.length > 0) {
      debug('Last-Modified response was present. This work is not yet implemented.');
      // Some types, typically direct entities, will return this value; collections do not.
      // Would want to use the Last-Modified over the refresh time, sorting to find the latest.
    }
    results.headers = {
      pages,
      dirty,
    };
    results.cost = compositeCost;
    return results;
  }

  private generalizedCollectionMethod(token: string | IGetAuthorizationHeader, apiName: string, method, options, cacheOptions: IPagedCacheOptions): Promise<IRestResponse> {
    const apiContext = new CompositeApiContext(apiName, method, options);
    apiContext.maxAgeSeconds = cacheOptions.maxAgeSeconds || 600;
    apiContext.overrideToken(token);
    apiContext.libraryContext = this.libraryContext;
    if (cacheOptions.backgroundRefresh) {
      apiContext.backgroundRefresh = true;
    }
    const compositeEngine = this.libraryContext.compositeEngine as CompositeIntelligentEngine;
    return compositeEngine.execute(apiContext);
  }

  private getCollectionAndFilter(token: string | IGetAuthorizationHeader, options, cacheOptions: IPagedCacheOptions, githubClientMethod, propertiesToKeep) {
    const capturedThis = this;
    return function (token, options) {
      return capturedThis.getFilteredGithubCollectionWithMetadataAnalysis(token, githubClientMethod, options, cacheOptions, propertiesToKeep);
    };
  }

  private async generalizedCollectionWithFilter(name, githubClientMethod, propertiesToKeep, token, options, cacheOptions: IPagedCacheOptions): Promise<any> {
    const rows = await this.generalizedCollectionMethod(token, name, this.getCollectionAndFilter(token, options, cacheOptions, githubClientMethod, propertiesToKeep), options, cacheOptions);
    const flattened = flattenData(rows);
    return flattened;
  }
}
