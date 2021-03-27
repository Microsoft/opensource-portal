//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import * as common from './common';
import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';

const repoPermissionProperties = ['permission', 'user'];

export enum GitHubCollaboratorPermissionLevel {
  Admin = 'admin',
  Write = 'write',
  Read = 'read',
  None = 'none',
}

export function ConvertGitHubCollaboratorPermissionLevelToGitHubRepositoryPermission(
  level: GitHubCollaboratorPermissionLevel
): GitHubRepositoryPermission {
  switch (level) {
    case GitHubCollaboratorPermissionLevel.None:
      return null;
    case GitHubCollaboratorPermissionLevel.Admin:
      return GitHubRepositoryPermission.Admin;
    case GitHubCollaboratorPermissionLevel.Write:
      return GitHubRepositoryPermission.Push;
    case GitHubCollaboratorPermissionLevel.Read:
      return GitHubRepositoryPermission.Pull;
    default:
      throw new Error(
        `ConvertGitHubCollaboratorPermissionLevelToGitHubRepositoryPermission unrecognized value ${level} cannot be translated`
      );
  }
}

export class RepositoryPermission {
  private _id: string;
  private _user: any;

  private _permission: GitHubCollaboratorPermissionLevel;

  constructor(entity: unknown) {
    if (entity) {
      common.assignKnownFieldsPrefixed(
        this,
        entity,
        'repositoryPermission',
        repoPermissionProperties
      );
      if (this._user) {
        this._id = this._user.id;
      }
    }
  }

  get id(): string {
    return this._id;
  }
  get permission(): GitHubCollaboratorPermissionLevel {
    return this._permission;
  }
  get user(): any {
    return this._user;
  }

  public asGitHubRepositoryPermission(): GitHubRepositoryPermission {
    return ConvertGitHubCollaboratorPermissionLevelToGitHubRepositoryPermission(
      this._permission
    );
  }
}
