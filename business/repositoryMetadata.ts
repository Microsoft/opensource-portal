//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';

const currentRepositoryMetadataSchema = 'rm1.0';

export interface IRepositoryMetadataPermissionPair {
  id: string;
  permission: GitHubRepositoryPermission;
}

export interface IRepositoryMetadata {
  schema: string;

  apiVersion: string;
  correlationId: string;

  id: string;

  created: Date;

  name: string;
  description: string;
  visibility: string;

  policy: string;
  policyUrl: string;

  license: string;
  legalEntity: string;

  template: string;
  gitIgnoreTemplate: string;

  teamPermissions: IRepositoryMetadataPermissionPair[];
}

export function ParseRepositoryMetadataSchema(
  fields: any
): IRepositoryMetadata {
  if (fields.schema === currentRepositoryMetadataSchema) {
    return fields as IRepositoryMetadata;
  }

  return new RepositoryMetadataLegacySchema(fields);
}

class RepositoryMetadataLegacySchema implements IRepositoryMetadata {
  public readonly schema: string = undefined;

  constructor(private _fields: any) {}

  static translateToOldSchemaValues(fields: IRepositoryMetadata): any {
    // translates to the old fields
  }

  get id(): string {
    return this._fields.repoId;
  }

  get created(): Date {
    return this._fields.requested;
  }

  get name(): string {
    return this._fields.repoName;
  }

  get description(): string {
    return this._fields.repoDescription;
  }

  get visibility(): string {
    return this._fields.repoVisibility;
  }

  get policy(): string {
    return this._fields.approvalType;
  }

  get policyUrl(): string {
    return this._fields.approvalUrl;
  }

  get template(): string {
    return this._fields.template;
  }

  get gitIgnoreTemplate(): string {
    return this._fields.gitignore_template;
  }

  get license(): string {
    return this._fields.license;
  }

  get legalEntity(): string {
    return this._fields.claEntity;
  }

  get apiVersion(): string {
    return this._fields.apiVersion;
  }

  get correlationId(): string {
    return this._fields.correlationId;
  }

  get teamPermissions(): IRepositoryMetadataPermissionPair[] {
    const data: IRepositoryMetadataPermissionPair[] = [];
    const count = (this._fields.teamsCount as number) || 0;
    for (let i = 0; i < count; i++) {
      const idFieldName = `teamid${i}`;
      const permissionFieldName = `${idFieldName}p`;
      const id = this._fields[idFieldName];
      const permissionStringValue = this._fields[permissionFieldName];
      let permission: GitHubRepositoryPermission =
        GitHubRepositoryPermission.Pull;
      if (permissionStringValue === 'push') {
        permission = GitHubRepositoryPermission.Push;
      } else if (permissionStringValue === 'admin') {
        permission = GitHubRepositoryPermission.Admin;
      }
      data.push({
        id,
        permission,
      });
    }
    return data;
  }

  getLegacyFields(): any {
    // direct access to preserve any extra fields
    return this._fields;
  }
}
