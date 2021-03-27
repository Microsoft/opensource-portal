//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { TeamJoinApprovalEntity } from './teamJoinApproval';
import { IEntityMetadataProvider } from '../../lib/entityMetadataProvider/entityMetadataProvider';

export interface IApprovalProvider {
  initialize(): Promise<void>;

  getApprovalEntity(approvalId: string): Promise<TeamJoinApprovalEntity>;
  createTeamJoinApprovalEntity(
    approval: TeamJoinApprovalEntity
  ): Promise<string>;
  updateTeamApprovalEntity(approval: TeamJoinApprovalEntity): Promise<void>;

  queryPendingApprovalsForTeam(id: string): Promise<TeamJoinApprovalEntity[]>;
  queryPendingApprovalsForTeams(
    ids: string[]
  ): Promise<TeamJoinApprovalEntity[]>;
  queryPendingApprovalsForThirdPartyId(
    thirdPartyId: string
  ): Promise<TeamJoinApprovalEntity[]>;
  queryAllApprovals(): Promise<TeamJoinApprovalEntity[]>;

  deleteAllRequests(): Promise<void>;
}

export interface IApprovalProviderCreateOptions {
  entityMetadataProvider: IEntityMetadataProvider;
}
