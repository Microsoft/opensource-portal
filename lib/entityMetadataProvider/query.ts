//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IEntityMetadataProvider } from './entityMetadataProvider';
import { EntityMetadataType, EntityMetadataBase } from './entityMetadata';

export enum FixedQueryType {
  // Team join
  ActiveTeamJoinApprovalsByTeam,
  ActiveTeamJoinApprovalsByTeams,
  ActiveTeamJoinApprovalsByThirdPartyId,
  AllActiveTeamJoinApprovals,
  AllTeamJoinApprovals,

  // Repo metadata
  AllRepositoryMetadata,
  RepositoryMetadataByRepositoryId,

  // Audit log entries
  AuditLogUndoCandidateRecordsByThirdPartyId,
  AuditLogRecordsByRepositoryId,
  AuditLogRecordsByTeamId,
  AuditLogRecordsByActorThirdPartyId,
  AuditLogRecordsByUserThirdPartyId,

  // Event entries
  EventRecordContributionsByThirdPartyId,
  EventRecordContributionsByDateRange,
  EventRecordContributionsByDateRangeAndId,
  EventRecordContributionsByDateRangeAndCorporateId,
  EventRecordDistinctOrganizations,
  EventRecordDistinctEligibleContributorsByDateRange,
  EventRecordPopularContributionsByDateRange,

  // Tokens
  TokensByCorporateId,
  TokensGetAll,

  // Key extension keys
  LocalExtensionKeysGetAll,

  // Orgs
  OrganizationMemberCacheGetAll,
  OrganizationMemberCacheByOrganizationId,
  OrganizationMemberCacheByUserId,
  OrganizationCacheGetAllBasics,
  OrganizationMemberCacheDeleteByOrganizationId,
  OrganizationOwnersCache,

  // Repo cache
  RepositoryCacheGetAll,
  RepositoryCacheGetByOrganizationId,
  RepositoryCacheDeleteByOrganizationId,
  RepositoryCacheGetOrganizationIds,

  // Repo collaborator cache
  RepositoryCollaboratorCacheGetAll,
  RepositoryCollaboratorCacheByOrganizationId,
  RepositoryCollaboratorCacheByRepositoryId,
  RepositoryCollaboratorCacheByUserId,
  RepositoryCollaboratorCacheDeleteByOrganizationId,
  RepositoryCollaboratorCacheDeleteByRepositoryId,
  RepositoryCollaboratorCacheGetOrganizationIds,

  // Repository team cache
  RepositoryTeamCacheGetAll,
  RepositoryTeamCacheByRepositoryId,
  RepositoryTeamCacheByTeamId,
  RepositoryTeamCacheByTeamIds,
  RepositoryTeamCacheByOrganizationId,
  RepositoryTeamCacheDeleteByOrganizationId,
  RepositoryTeamCacheDeleteByRepositoryId,
  RepositoryTeamCacheGetOrganizationIds,

  // Team cache
  TeamCacheGetAll,
  TeamCacheGetByOrganizationId,
  TeamCacheGetOrganizationIds,
  TeamCacheDeleteByOrganizationId,

  // Team member cache
  TeamMemberCacheGetAll,
  TeamMemberCacheGetByOrganizationId,
  TeamMemberCacheGetByUserId,
  TeamMemberCacheGetByTeamId,
  TeamMemberCacheGetByOrganizationIdAndUserId,
  TeamMemberCacheGetOrganizationIds,
  TeamMemberCacheDeleteByOrganizationId,

  // Organization settings
  OrganizationSettingsGetAll,
  OrganizationSettingsGetMostRecentlyUpdatedActive,

  // Shim to evolve this code
  Shim,
}

export interface IEntityMetadataFixedQuery {
  // A fixed set of query provider is less useful than a dynamic querying system
  // but also will help support a specific set of queries without too much
  // trouble or a generic query builder interface until needed

  fixedQueryType: FixedQueryType;
}

export abstract class QueryBase<T> implements IEntityMetadataFixedQuery {
  fixedQueryType = FixedQueryType.Shim;

  public async discover(
    provider: EntityMetadataBase,
    entities: IEntityMetadataProvider,
    thisProviderType: EntityMetadataType
  ): Promise<T[]> {
    const metadatas = await entities.fixedQueryMetadata(thisProviderType, this);
    const deserializeArray = provider['deserializeArray']; // using as an internal call
    if (!deserializeArray) {
      throw new Error('No provider.deserializeArray private method');
    }
    const deserialize = deserializeArray.bind(provider, thisProviderType);
    const results = deserialize(metadatas);
    return results as T[];
  }

  public async generic(
    provider: EntityMetadataBase,
    entities: IEntityMetadataProvider,
    thisProviderType: EntityMetadataType
  ): Promise<any[]> {
    const metadatas = await entities.fixedQueryMetadata(thisProviderType, this);
    return metadatas as any[];
  }
}
