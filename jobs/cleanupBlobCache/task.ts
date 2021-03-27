//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import BlobCache from '../../lib/caching/blob';
import { IReposJob } from '../../app';

export default async function go({ providers }: IReposJob): Promise<void> {
  for (const providerName in providers) {
    const provider = providers[providerName];
    if (provider && provider['expiringBlobCache']) {
      const expiringBlobCache = provider['expiringBlobCache'] as BlobCache;
      if (expiringBlobCache && expiringBlobCache.deleteExpiredBlobs) {
        console.log(`blob cache found: ${providerName}`);
        try {
          const stats = await expiringBlobCache.deleteExpiredBlobs();
          console.log(`
Provider name:   ${providerName}
Expired blobs:   ${stats.expired}
Blobs processed: ${stats.processedBlobs}
Pages processed: ${stats.processedPages}
Errors:          ${stats.errors.length}
`);
          if (stats.errors.length) {
            console.dir(stats.errors);
          }
        } catch (expiringDeleteError) {
          console.log(
            `Unhandled error while deleting expiring blobs for provider ${providerName}`
          );
          console.dir(expiringDeleteError);
        }
      }
    }
  }
}
