//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Team } from './team';
import { Repository } from './repository';
import { IOperationsInstance } from '../transitional';

export class TeamRepositoryPermission {
  private _team: Team;
  private _operations: IOperationsInstance;
  private _permissions: any;
  private _repository: Repository;
  private _id: number;

  constructor(team: Team, entity: any, operations: IOperationsInstance) {
    this._team = team;
    if (!entity) {
      throw new Error('TeamRepositoryPermission: requires entity');
    }
    this._permissions = entity.permissions;
    this._repository = team.organization.repositoryFromEntity(entity);
    this._id = this._repository.id;
    this._operations = operations;
  }

  asJson() {
    const repo = this._repository.asJson();
    const permissions = this._permissions;
    const combined = { ...repo, permissions };
    return combined;
  }

  get team(): Team {
    return this._team;
  }

  get repository(): Repository {
    return this._repository;
  }

  get id(): number {
    return this._id;
  }

  get permissions(): any {
    return this._permissions;
  }

  get name(): string {
    return this._repository.name;
  }
}
