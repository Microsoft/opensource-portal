//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IProviders } from '../../transitional';
import { MicrosoftGraphProvider } from './microsoftGraphProvider';

export enum GraphUserType {
  Unknown = '', // most employees
  Guest = 'Guest',
  Member = 'Member', // some users, like LinkedIn employees, are a member
}

export interface IGraphEntry {
  displayName: string;
  givenName: string;
  id: string;
  mail: string;
  userPrincipalName: string;
  userType?: GraphUserType;
  mailNickname?: string;
  // alias?: string;
  jobTitle?: string;
}

export interface IGraphGroupMember {
  id: string;
  userPrincipalName: string;
}

export interface IGraphGroup {
  id: string;
  displayName: string;
  mailNickname: string;

  description?: string;
  mail?: string;
}

export interface IGraphEntryWithManager extends IGraphEntry {
  manager: IGraphEntry;
}

export interface IGraphProvider {
  getUserById(id: string): Promise<IGraphEntry>;

  getUserIdByNickname(nickname: string): Promise<string>;

  getUserAndManagerById(corporateId: string): Promise<IGraphEntryWithManager>;

  getManagerById(corporateId: string): Promise<IGraphEntry>;
  getManagementChain(corporateId: string): Promise<IGraphEntry[]>;

  getMailAddressByUsername(corporateUsername: string): Promise<string>;
  getUserIdByUsername(corporateUsername: string): Promise<string>;

  getUsersBySearch(minimum3Characters: string): Promise<IGraphEntry[]>;
  getUsersByIds(userIds: string[]): Promise<IGraphEntry[]>;
  getUsersByMailNicknames(mailNicknames: string[]): Promise<IGraphEntry[]>;

  getGroupsById(corporateId: string): Promise<string[]>;
  getGroupsByMail(mailAddress: string): Promise<string[]>;
  getGroupsByNickname(nickname: string): Promise<string[]>;
  getGroupsStartingWith(minimum3Characters: string): Promise<IGraphGroup[]>;
  getGroupMembers(corporateGroupId: string): Promise<IGraphGroupMember[]>;
  getGroup(corporateGroupId: string): Promise<IGraphGroup>;
  isUserInGroup(corporateId: string, securityGroupId: string): Promise<boolean>;

  getToken(): Promise<string>;
}

export function CreateGraphProviderInstance(
  providers: IProviders,
  config: any,
  callback
) {
  const activeDirectoryConfig = config.activeDirectory;
  const graphConfig = Object.assign(
    {
      tenantId: activeDirectoryConfig.tenantId,
    },
    config.graph
  );
  if (!graphConfig) {
    return callback(new Error('No graph provider configuration.'));
  }
  const provider = graphConfig.provider;
  if (!provider) {
    return callback(new Error('No graph provider set in the graph config.'));
  }
  let providerInstance: IGraphProvider = null;
  try {
    switch (provider) {
      case 'microsoftGraphProvider':
        if (providers?.cacheProvider) {
          graphConfig.cacheProvider = providers.cacheProvider;
        }
        providerInstance = new MicrosoftGraphProvider(graphConfig);
        break;
      default:
        break;
    }
  } catch (createError) {
    return callback(createError);
  }

  if (!providerInstance) {
    return callback(
      new Error(
        `The graph provider "${provider}" is not implemented or configured at this time.`
      )
    );
  }

  return callback(null, providerInstance);
}
