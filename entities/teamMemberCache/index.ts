//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  ITeamMemberCacheProvider,
  ITeamMemberCacheCreateOptions,
  TeamMemberCacheProvider,
} from './teamMemberCacheProvider';
import {
  FixedQueryType,
  IEntityMetadataFixedQuery,
} from '../../lib/entityMetadataProvider/query';

export async function CreateTeamMemberCacheProviderInstance(
  options?: ITeamMemberCacheCreateOptions
): Promise<ITeamMemberCacheProvider> {
  const provider = new TeamMemberCacheProvider(options);
  await provider.initialize();
  return provider;
}

export class TeamMemberCacheFixedQueryAll implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamMemberCacheGetAll;
}

export class TeamMemberCacheGetOrganizationIdsQuery
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamMemberCacheGetOrganizationIds;
}

export class TeamMemberCacheDeleteByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamMemberCacheDeleteByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class TeamMemberCacheFixedQueryByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamMemberCacheGetByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class TeamMemberCacheFixedQueryByTeamId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamMemberCacheGetByTeamId;
  constructor(public teamId: string) {
    if (typeof this.teamId !== 'string') {
      throw new Error(`${teamId} must be a string`);
    }
  }
}

export class TeamMemberCacheFixedQueryByUserId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamMemberCacheGetByUserId;
  constructor(public userId: string) {
    if (typeof this.userId !== 'string') {
      throw new Error(`${userId} must be a string`);
    }
  }
}

export class TeamMemberCacheFixedQueryByOrganizationIdAndUserId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamMemberCacheGetByOrganizationIdAndUserId;
  constructor(public organizationId: string, public userId: string) {
    if (typeof this.userId !== 'string') {
      throw new Error(`userId ${userId} must be a string`);
    }
    if (typeof this.organizationId !== 'string') {
      throw new Error(`organizationId ${organizationId} must be a string`);
    }
  }
}
