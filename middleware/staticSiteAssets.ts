//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import appPackage from '../package.json';

const debug = require('debug')('startup');

import favicon from 'serve-favicon';
import path from 'path';

const defaultPublicAssetsPackageName = '../../default-assets-package/';
const staticAssetspackageName = appPackage['static-site-assets-package-name'] || defaultPublicAssetsPackageName;

const ospoAssetsDistPath = require(staticAssetspackageName);
const ospoAssetsPackage = require(`${staticAssetspackageName}/package.json`);

export function StaticSiteAssets(app, express) {
  // Serve/host the static site assets from our private NPM

  debug(`Hosting site assets version ${ospoAssetsPackage.version} from ${ospoAssetsDistPath} on path '/'`);
  app.use(express.static(ospoAssetsDistPath));
};

export function StaticSiteFavIcon(app) {
  try {
    app.use(favicon(path.join(ospoAssetsDistPath, 'favicon.ico')));
  } catch (nofavicon) {
    console.error(`The static site assets in "${ospoAssetsDistPath}" does not include a favicon. You may need to run 'npm install' in the package folder first.`);
    throw nofavicon;
  }
}
