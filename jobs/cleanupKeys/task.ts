//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import throat from 'throat';

import { IReposJob, IReposJobResult } from '../../interfaces';
import { sleep } from '../../utils';
import { IGraphProvider } from '../../lib/graphProvider';
import { LocalExtensionKey } from '../../entities/localExtensionKey/localExtensionKey';

async function lookupCorporateId(graphProvider: IGraphProvider, knownUsers: Map<string, any>, corporateId: string): Promise<any> {
  let entry = knownUsers.get(corporateId);
  if (entry === false) {
    return false;
  } else if (entry) {
    return true;
  }

  try {
    const userDetails = await graphProvider.getUserById(corporateId);
    if (!userDetails || !userDetails.userPrincipalName) {
      knownUsers.set(corporateId, false);
      return false;
    }
    knownUsers.set(corporateId, userDetails);
    return true;
  } catch (otherUserError) {
    console.dir(otherUserError);
    throw otherUserError;
  }
}

export default async function cleanup({ providers }: IReposJob): Promise<IReposJobResult> {
  const graphProvider = providers.graphProvider;
  const localExtensionKeyProvider = providers.localExtensionKeyProvider;
  const insights = providers.insights;

  console.log('reading all keys');
  const allKeys = await localExtensionKeyProvider.getAllKeys();
  console.log(`read ${allKeys.length}`);

  insights.trackEvent({ name: 'JobCleanupTokensKeysTokens', properties: { tokens: String(allKeys.length) } });

  let errors = 0;

  let deleted = 0;
  let okUserTokens = 0;

  const parallelUsers = 2;
  const secondsDelayAfterSuccess = 0.25;

  const knownUsers = new Map<string, any>();

  const throttle = throat(parallelUsers);
  await Promise.all(allKeys.map((key: LocalExtensionKey) => throttle(async () => {
    const corporateId = key.corporateId;
    const userStatus = await lookupCorporateId(graphProvider, knownUsers, corporateId);
    if (!userStatus) {
      try {
        ++deleted;
        console.log(`${deleted}: Deleting key for ${corporateId} that could not be found`);
        await localExtensionKeyProvider.delete(key);
      } catch (tokenDeleteError) {
        --deleted;
        console.dir(tokenDeleteError);
        ++errors;
        insights.trackException({ exception: tokenDeleteError });
      }
    } else {
      ++okUserTokens;
      console.log(`${okUserTokens}: valid`);
    }

    await sleep(secondsDelayAfterSuccess * 1000);

  })));

  console.log(`deleted: ${deleted}`);
  console.log(`okUserTokens: ${okUserTokens}`);
  console.log();

  return {
    successProperties: {
      deleted,
      okUserTokens,
      errors,
    },
  };
}
