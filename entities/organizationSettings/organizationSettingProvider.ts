//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IEntityMetadata,
  EntityMetadataBase,
  IEntityMetadataBaseOptions,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  OrganizationSetting,
  OrganizationSettingFixedQueryAll,
  OrganizationSettingFixedQueryMostRecentlyUpdatedActive,
} from './organizationSetting';
import { EntityImplementation } from './organizationSetting';

const thisProviderType = EntityImplementation.Type;

export interface IOrganizationSettingCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IOrganizationSettingProvider {
  initialize(): Promise<void>;

  getOrganizationSetting(organizationId: string): Promise<OrganizationSetting>;
  createOrganizationSetting(metadata: OrganizationSetting): Promise<string>;
  updateOrganizationSetting(metadata: OrganizationSetting): Promise<void>;
  queryAllOrganizations(): Promise<OrganizationSetting[]>;
  deleteOrganizationSetting(metadata: OrganizationSetting): Promise<void>;
}

export class OrganizationSettingProvider
  extends EntityMetadataBase
  implements IOrganizationSettingProvider {
  constructor(options: IOrganizationSettingCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async deleteOrganizationSetting(
    metadata: OrganizationSetting
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.deleteMetadata(entity);
  }

  async getOrganizationSetting(
    organizationId: string
  ): Promise<OrganizationSetting> {
    this.ensureHelpers(thisProviderType);
    let metadata: IEntityMetadata = null;
    metadata = await this._entities.getMetadata(
      thisProviderType,
      organizationId
    );
    if (!metadata) {
      const error = new Error(
        `No metadata available for organization ${organizationId}`
      );
      error['status'] = 404;
      throw error;
    }
    return this.deserialize<OrganizationSetting>(thisProviderType, metadata);
  }

  async queryAllOrganizations(): Promise<OrganizationSetting[]> {
    const query = new OrganizationSettingFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<OrganizationSetting>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async createOrganizationSetting(
    metadata: OrganizationSetting
  ): Promise<string> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.setMetadata(entity);
    return entity.entityId;
  }

  async updateOrganizationSetting(
    metadata: OrganizationSetting
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, metadata);
    await this._entities.updateMetadata(entity);
  }
}
