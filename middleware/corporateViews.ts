//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { constants as fsConstants, promises as fs } from 'fs';
import path from 'path';

import { IProviders, stripDistFolderName } from '../transitional';

// providers.corporateViews:
// ---
// This provider initializes a set of defined views discovered at startup time
// in the directory ./views/corporate/...
//
// This allows for dynamic inclusion of corporate content without too much impact
// on other users of the project, for the time being.
//
// All views served through the 'webcontext' for the site include the
// result.
//
// Use may look similar to this in a Pug template:
//
// if corporateViews && corporateViews.directoryname.viewname
//   include 'corporate/directoryname/viewname'
//

export default async function initializeCorporateViews(
  providers: IProviders,
  dirname: string
): Promise<any> {
  const { config } = providers.config;
  const appDirectory =
    config && config.typescript && config.typescript.appDirectory
      ? config.typescript.appDirectory
      : stripDistFolderName(dirname);
  const corporateViewsRoot = path.resolve(
    path.join(appDirectory, 'views', 'corporate')
  );
  try {
    await fs.access(corporateViewsRoot, fsConstants.R_OK);
  } catch (err) {
    return null;
  }
  return await recurseDirectory(corporateViewsRoot);
}

async function recurseDirectory(dir: string) {
  const thisLevel = {};
  for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      thisLevel[entry.name] = await recurseDirectory(
        path.join(dir, entry.name)
      );
    } else if (entry.isFile()) {
      const baseName = path.basename(entry.name, path.extname(entry.name));
      thisLevel[baseName] = true;
    }
  }
  return thisLevel;
}
