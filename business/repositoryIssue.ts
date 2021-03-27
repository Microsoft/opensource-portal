//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Repository } from './repository';
import {
  IPurposefulGetAuthorizationHeader,
  ICacheOptions,
  IGetAuthorizationHeader,
  IOperationsInstance,
  throwIfNotGitHubCapable,
} from '../transitional';
import { wrapError } from '../utils';
import { AppPurpose } from '../github';
import { CacheDefault, getMaxAgeSeconds } from '.';

export enum GitHubIssueState {
  Closed = 'closed',
  Open = 'open',
}

export enum GitHubIssueQuery {
  Closed = 'closed',
  Open = 'open',
  All = 'all',
}

export interface IIssueLabel {
  id: number;
  node_id: string;
  // url: string;
  name: string;
  description?: string;
  color?: string;
  default?: boolean;
}

export class RepositoryIssue {
  private _operations: IOperationsInstance;
  private _getAuthorizationHeader: IPurposefulGetAuthorizationHeader;

  private _number: number;
  private _repository: Repository;

  private _entity: any;

  constructor(
    repository: Repository,
    issueNumber: number,
    operations: IOperationsInstance,
    getAuthorizationHeader: IPurposefulGetAuthorizationHeader,
    entity?: any
  ) {
    this._getAuthorizationHeader = getAuthorizationHeader;
    this._repository = repository;
    this._number = issueNumber;
    this._operations = operations;
    if (entity) {
      this._entity = entity;
    }
  }

  get id(): number {
    return this._entity?.id as number;
  }
  get title(): string {
    return this._entity?.title as string;
  }
  get body(): string {
    return this._entity.body as string;
  }
  get state(): GitHubIssueState {
    return this._entity?.state as GitHubIssueState;
  }
  get labels(): IIssueLabel[] {
    if (this._entity) {
      return this._entity.labels as IIssueLabel[];
    } else {
      return null;
    }
  }

  get number(): number {
    return this._number;
  }

  getEntity(): any {
    return this._entity;
  }

  get repository(): Repository {
    return this._repository;
  }

  async update(patch: any): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = Object.assign(patch, {
      owner: this.repository.organization.name,
      repo: this.repository.name,
      issue_number: this.number,
    });
    // Operations has issue write permissions
    const details = await operations.github.post(
      this.authorize(AppPurpose.Operations),
      'issues.update',
      parameters
    );
    return details;
  }

  async comment(commentBody: string): Promise<any> {
    const operations = throwIfNotGitHubCapable(this._operations);
    const parameters = Object.assign(
      {
        body: commentBody,
      },
      {
        owner: this.repository.organization.name,
        repo: this.repository.name,
        issue_number: this.number,
      }
    );
    // Operations has issue write permissions
    const comment = await operations.github.post(
      this.authorize(AppPurpose.Operations),
      'issues.createComment',
      parameters
    );
    return comment;
  }

  async getDetails(options?: ICacheOptions): Promise<any> {
    if (this._entity) {
      return this._entity;
    }
    options = options || {};
    const operations = throwIfNotGitHubCapable(this._operations);
    if (!this._repository.name) {
      throw new Error('repository.name required');
    }
    const parameters = {
      owner: this._repository.organization.name,
      repo: this._repository.name,
      issue_number: this._number,
    };
    const cacheOptions: ICacheOptions = {
      // NOTE: just reusing repo details stale time
      maxAgeSeconds: getMaxAgeSeconds(
        operations,
        CacheDefault.orgRepoDetailsStaleSeconds,
        options
      ),
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    try {
      const entity = await operations.github.call(
        this.authorize(AppPurpose.Data),
        'issues.get',
        parameters,
        cacheOptions
      );
      this._entity = entity;
      return entity;
    } catch (error) {
      const notFound = error.status && error.status == /* loose */ 404;
      error = wrapError(
        error,
        notFound
          ? 'The issue could not be found.'
          : 'Could not get details about the issue.',
        notFound
      );
      if (notFound) {
        error.status = 404;
      }
      throw error;
    }
  }

  async isDeleted(options?: ICacheOptions): Promise<boolean> {
    try {
      await this.getDetails(options);
    } catch (maybeDeletedError) {
      if (
        maybeDeletedError &&
        maybeDeletedError.status &&
        maybeDeletedError.status === 404
      ) {
        return true;
      }
    }
    return false;
  }

  private authorize(purpose: AppPurpose): IGetAuthorizationHeader | string {
    const getAuthorizationHeader = this._getAuthorizationHeader.bind(
      this,
      purpose
    ) as IGetAuthorizationHeader;
    return getAuthorizationHeader;
  }
}
