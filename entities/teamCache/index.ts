//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  ITeamCacheProvider,
  ITeamCacheCreateOptions,
  TeamCacheProvider,
} from './teamCacheProvider';
import {
  FixedQueryType,
  IEntityMetadataFixedQuery,
} from '../../lib/entityMetadataProvider/query';

export async function CreateTeamCacheProviderInstance(
  options?: ITeamCacheCreateOptions
): Promise<ITeamCacheProvider> {
  const provider = new TeamCacheProvider(options);
  await provider.initialize();
  return provider;
}

export class TeamCacheFixedQueryAll implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamCacheGetAll;
}

export class TeamCacheFixedQueryByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamCacheGetByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}

export class TeamCacheGetOrganizationIdsQuery
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamCacheGetOrganizationIds;
}

export class TeamCacheDeleteByOrganizationId
  implements IEntityMetadataFixedQuery {
  public readonly fixedQueryType: FixedQueryType =
    FixedQueryType.TeamCacheDeleteByOrganizationId;
  constructor(public organizationId: string) {
    if (typeof this.organizationId !== 'string') {
      throw new Error(`${organizationId} must be a string`);
    }
  }
}
