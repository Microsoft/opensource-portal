//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ReposAppRequest } from '../interfaces';
import thisPackage = require('../package.json');
import { getProviders } from '../transitional';

import { Router } from 'express';
const router: Router = Router();

var cachedPackageInformation = null;

// Super-synchronous but rarely used page...
function getPackageInfo(config) {
  if (cachedPackageInformation) {
    return cachedPackageInformation;
  }
  // var thisPackage = require('/package.json');
  cachedPackageInformation = {};
  const privateFeedScope = config && config.npm && config.npm.privateFeedScope ? config.npm.privateFeedScope : 'no-configured-private-feed-scope';
  for (var dependency in thisPackage.dependencies) {
    var componentPackage = require('../node_modules/' + dependency + '/package.json');
    if (componentPackage && componentPackage.name && !componentPackage.name.includes(`@${privateFeedScope}`)) {
      cachedPackageInformation[dependency] = {
        homepage: componentPackage.homepage,
        description: componentPackage.description,
      };
    }
  }
  return cachedPackageInformation;
}

router.get('/', function (req: ReposAppRequest, res) {
  const { config: completeConfig} = getProviders(req);
  const config = completeConfig.obfuscatedConfig;
  const components = getPackageInfo(config);
  res.render('thanks', {
    config: config,
    components: components,
    serviceBanner: config && config.serviceMessage ? config.serviceMessage.banner : undefined,
    title: 'Open Source Components',
  });
});

export default router;
