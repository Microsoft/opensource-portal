//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';
import * as common from './common';

const memberPrimaryProperties = ['id', 'login', 'permissions', 'avatar_url'];

export interface IGitHubCollaboratorPermissions {
  admin: boolean;
  pull: boolean;
  push: boolean;
  // triage and maintain do not appear today by the GitHub API (sigh), it's in V4 GraphQL but not in V3 REST
}

export class Collaborator {
  public static PrimaryProperties = memberPrimaryProperties;

  private _avatar_url: string;
  private _id: number;
  private _login: string;
  private _permissions: IGitHubCollaboratorPermissions;

  constructor(entity: unknown) {
    if (entity) {
      common.assignKnownFieldsPrefixed(
        this,
        entity,
        'member',
        memberPrimaryProperties
      );
    }
  }

  asJson() {
    return {
      avatar_url: this.avatar_url,
      id: this._id,
      login: this._login,
      permissions: this._permissions,
    };
  }

  get permissions(): IGitHubCollaboratorPermissions {
    return this._permissions;
  }

  getHighestPermission() {
    if (!this._permissions) {
      return GitHubRepositoryPermission.None;
    }
    const permissions = this._permissions;
    if (permissions.admin) {
      return GitHubRepositoryPermission.Admin;
    } else if (permissions.push) {
      return GitHubRepositoryPermission.Push;
    } else if (permissions.pull) {
      return GitHubRepositoryPermission.Pull;
    }
    throw new Error(`Unsupported permission type by getHighestPermission`);
  }

  get id(): number {
    return this._id;
  }

  get login(): string {
    return this._login;
  }

  get avatar_url(): string {
    return this._avatar_url;
  }
}
