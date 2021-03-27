//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

// This is a transition migration job that takes the former link source of truth -
// table links - and copies those links into the configured provider if it is different.

// Assumes the source and destination providers are of different types so
// that the default configuration is required for both.
//
// Also requires migration environment variables:
// LINK_MIGRATION_DESTINATION_TYPE
// LINK_MIGRATION_OVERWRITE  values : 'overwrite', 'skip'

import throat from 'throat';

import app, { IReposJob } from '../../app';
import {
  createAndInitializeLinkProviderInstance,
  ILinkProvider,
} from '../../lib/linkProviders';
import { ICorporateLink } from '../../business/corporateLink';
import { ErrorHelper } from '../../transitional';

const parallelWorkLimit = 5;

export default async function migration({
  providers,
}: IReposJob): Promise<void> {
  // const sourceLinkProvider = providers.linkProvider;
  const config = providers.config;
  const sourceLinkProviderName = 'table';
  console.log(`creating source ${sourceLinkProviderName} provider`);
  const sourceLinkProvider = await createAndInitializeLinkProviderInstance(
    providers,
    config,
    sourceLinkProviderName
  );

  const destinationLinkProviderName = 'postgres';

  console.log(`creating destination ${destinationLinkProviderName} provider`);
  const destinationLinkProvider = await createAndInitializeLinkProviderInstance(
    providers,
    config,
    destinationLinkProviderName
  );

  console.log('downloading all source links');
  const allSourceLinks = await sourceLinkProvider.getAll();
  console.log(`SOURCE: ${allSourceLinks.length} links`);

  // const clearDestinationLinksFirst = false;

  const overwriteDestinationLinks = false;

  console.log(`migrating ${allSourceLinks.length} links...`);
  let errors = 0;
  let errorList = [];

  const throttle = throat(parallelWorkLimit);
  await Promise.all(
    allSourceLinks.map((sourceLink: ICorporateLink) =>
      throttle(async () => {
        const existingLink = await getThirdPartyLink(
          destinationLinkProvider,
          sourceLink.thirdPartyId
        );
        if (existingLink && overwriteDestinationLinks) {
          console.warn('Removing existing destination link...');
          await destinationLinkProvider.deleteLink(existingLink);
        } else if (existingLink && overwriteDestinationLinks === false) {
          return '$';
        }

        console.log(
          `Creating link in destination provider for corp ${sourceLink.corporateUsername} 3p ${sourceLink.thirdPartyUsername}...`
        );
        try {
          if (!sourceLink.corporateId) {
            // need to use the graph!
            const id = await getUserIdByUpn(
              providers.graphProvider,
              sourceLink.corporateUsername
            );
            if (id === null) {
              throw new Error(
                `not found user ${sourceLink.corporateUsername} in graph`
              );
            }
            console.log(
              `discovered id ${id} for upn ${sourceLink.corporateUsername}`
            );
            sourceLink.corporateId = id;
          }

          const newLinkId = await destinationLinkProvider.createLink(
            sourceLink
          );
          console.log(`OK: new link ID in destination: ${newLinkId}`);
        } catch (linkCreateError) {
          console.log('Issue with link:');
          console.dir(sourceLink);
          console.warn(linkCreateError);
          ++errors;
          errorList.push(linkCreateError);
          return 'e';
          // throw linkCreateError;
        }
        console.log('[next]');
        return 'x';
      })
    )
  );

  console.log('All done with ' + errors + ' errors');
  console.dir(errorList);
  console.log();
}

async function getThirdPartyLink(
  linkProvider: ILinkProvider,
  thirdPartyId: string
): Promise<ICorporateLink> {
  try {
    return await linkProvider.getByThirdPartyId(thirdPartyId);
  } catch (error) {
    if (ErrorHelper.IsNotFound(error)) {
      return null;
    }
    throw error;
  }
}

async function getUserIdByUpn(graphProvider, upn: string): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    graphProvider.getUserById(upn, (err, info) => {
      if (err && err['status'] === 404) {
        console.log('User no longer around');
        return resolve(null);
      }
      if (err) {
        return reject(err);
      }
      return resolve(info.id);
    });
  });
}
