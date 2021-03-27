//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

const debug = require('debug')('startup');

import path from 'path';

import { IReposApplication } from '../app';
import { IApplicationProfile } from '../transitional';

export default async function initializeAlternateApps(
  config,
  app: IReposApplication,
  appName: string
): Promise<IApplicationProfile> {
  const appPath = path.resolve(path.join(__dirname, '..', appName, '/'));
  debug(`Alternate app requested: name=${appName}, path=${appPath}`);
  try {
    let setupApp = require(appPath);
    // support modern imports
    if (typeof setupApp !== 'function' && setupApp.default) {
      setupApp = setupApp.default;
    } else if (typeof setupApp !== 'function') {
      throw new Error(
        `Could not prepare default import for alternate app name=${appName}, path=${appPath}`
      );
    }
    return await setupApp(config, app);
  } catch (loadAlternateAppError) {
    console.log(`Loading error: ${loadAlternateAppError}`);
    throw loadAlternateAppError;
  }
}
