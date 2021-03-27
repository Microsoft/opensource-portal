//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import util from 'util';

import { Organization } from './organization';
import { Team, TeamJsonFormat } from './team';
import {
  GitHubTeamPrivacy,
  ICacheOptions,
  IOperationsInstance,
} from '../transitional';
import { TeamMember } from './teamMember';
import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';

export class TeamPermission {
  private _organization: Organization;
  private _operations: IOperationsInstance;

  private _team: Team;

  private _permission: GitHubRepositoryPermission;
  private _privacy: GitHubTeamPrivacy;

  private _teamMembersIfSet: TeamMember[];

  [util.inspect.custom](depth, options) {
    return `GitHub Team Permission: team=${
      this.team?.slug || this.team?.id
    } permission=${this._permission}`;
  }

  asJson() {
    const members = this._teamMembersIfSet;
    return {
      permission: this._permission,
      privacy: this._privacy,
      team: this._team?.asJson(TeamJsonFormat.Augmented),
      members: members ? members.map((member) => member.asJson()) : undefined,
    };
  }

  get permission(): GitHubRepositoryPermission {
    return this._permission;
  }

  get privacy(): GitHubTeamPrivacy {
    return this._privacy;
  }

  get team(): Team {
    return this._team;
  }

  constructor(
    organization: Organization,
    entity: any,
    operations: IOperationsInstance
  ) {
    this._organization = organization;

    this._permission = entity.permission;
    this._privacy = entity.privacy;

    if (!entity || !entity.id) {
      throw new Error('TeamPermission requires entity.id');
    }
    const id = entity.id;
    this._team = organization.team(id, entity);

    this._operations = operations;
  }

  get relativeJoinLink() {
    return this.relativePortalLink + 'join';
  }

  get relativePortalLink() {
    return `/${this._organization.name}/teams/${
      this._team.slug || this._team.name
    }/`;
  }

  get members(): TeamMember[] {
    if (this._teamMembersIfSet) {
      return this._teamMembersIfSet;
    }
  }

  async resolveTeamMembers(options?: ICacheOptions): Promise<void> {
    this._teamMembersIfSet = await this.team.getMembers(options);
  }
}
