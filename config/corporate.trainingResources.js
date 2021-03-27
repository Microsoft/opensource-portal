//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

const fs = require('fs');
const path = require('path');

const debug = require('debug')('startup');

const pkg = require('../package.json');

const painlessConfigEnvPkgName = 'painlessConfigEnvironments';
const resourcesEnvironmentName = 'urls';
const painlessConfigEnvironmentVariableName = 'CONFIGURATION_ENVIRONMENT';

const typescriptConfig = require('./typescript');

module.exports = function (graphApi) {
  const environmentProvider = graphApi.environment;
  const environmentName =
    environmentProvider.get(painlessConfigEnvironmentVariableName) ||
    environmentProvider.get('ENV');

  let resources = null;

  // 1: load URL/resource links from a parallel painless config environment
  if (pkg && pkg[painlessConfigEnvPkgName] && environmentName) {
    try {
      resources = require(pkg[painlessConfigEnvPkgName])(
        environmentName,
        resourcesEnvironmentName
      );
      debug(
        `resources and URL links loaded from ${painlessConfigEnvPkgName}/${environmentName},${resourcesEnvironmentName}`
      );
    } catch (painlessConfigError) {
      debug(
        `attempted to load resources and URL links loaded from ${painlessConfigEnvPkgName}/${environmentName},${resourcesEnvironmentName}`
      );
      console.warn(painlessConfigError);
      throw painlessConfigError;
    }
  } else {
    // 2: load UJRL/resource links data from a local JSON file
    try {
      const filename = path.join(
        typescriptConfig.appDirectory,
        'data',
        'resources.json'
      );
      const str = fs.readFileSync(filename, 'utf8');
      resources = JSON.parse(str);
      debug(`resources and URL links loaded from file ${filename}`);
    } catch (notFound) {
      console.warn(notFound);
      throw notFound;
    }
  }

  return resources;
};
