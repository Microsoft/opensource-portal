//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IApprovalProvider } from './approvalProvider';
import {
  IEntityMetadataBaseOptions,
  EntityMetadataBase,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  TeamJoinApprovalEntity,
  TeamJoinRequestFixedQueryByTeam,
  TeamJoinRequestFixedQueryByTeams,
  TeamJoinRequestFixedQueryByThirdPartyUserId,
  TeamJoinRequestFixedQueryAll,
  EntityImplementation,
} from './teamJoinApproval';

const thisProviderType = EntityImplementation.Type;

export interface ITeamJoinApprovalProviderOptions
  extends IEntityMetadataBaseOptions {}

export class TeamJoinApprovalProvider
  extends EntityMetadataBase
  implements IApprovalProvider {
  constructor(options: ITeamJoinApprovalProviderOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getApprovalEntity(approvalId: string): Promise<TeamJoinApprovalEntity> {
    this.ensureHelpers(thisProviderType);
    const metadata = await this._entities.getMetadata(
      thisProviderType,
      approvalId
    );
    return this.deserialize<TeamJoinApprovalEntity>(thisProviderType, metadata);
  }

  async queryPendingApprovalsForTeam(
    id: string
  ): Promise<TeamJoinApprovalEntity[]> {
    const query = new TeamJoinRequestFixedQueryByTeam(id);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamJoinApprovalEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryPendingApprovalsForTeams(
    ids: string[]
  ): Promise<TeamJoinApprovalEntity[]> {
    const query = new TeamJoinRequestFixedQueryByTeams(ids);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamJoinApprovalEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryPendingApprovalsForThirdPartyId(
    thirdPartyId: string
  ): Promise<TeamJoinApprovalEntity[]> {
    const query = new TeamJoinRequestFixedQueryByThirdPartyUserId(thirdPartyId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamJoinApprovalEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryAllApprovals(): Promise<TeamJoinApprovalEntity[]> {
    const query = new TeamJoinRequestFixedQueryAll();
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<TeamJoinApprovalEntity>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async deleteAllRequests(): Promise<void> {
    await this._entities.clearMetadataStore(thisProviderType);
  }

  async createTeamJoinApprovalEntity(
    approval: TeamJoinApprovalEntity
  ): Promise<string> {
    const entity = this.serialize(thisProviderType, approval);
    await this._entities.setMetadata(entity);
    return entity.entityId;
  }

  async updateTeamApprovalEntity(
    approval: TeamJoinApprovalEntity
  ): Promise<void> {
    const entity = this.serialize(thisProviderType, approval);
    return await this._entities.updateMetadata(entity);
  }
}
