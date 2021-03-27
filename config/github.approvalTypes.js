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
const repoApprovalsEnvironmentName = 'repo.approvals';

const arrayFromString = require('./utils/arrayFromString');
const typescriptConfig = require('./typescript');

const approvalFieldsFileVariableName = 'GITHUB_APPROVAL_FIELDS_FILE';
const painlessConfigEnvironmentVariableName = 'CONFIGURATION_ENVIRONMENT';

module.exports = function (graphApi) {
  const environmentProvider = graphApi.environment;
  const fieldsFile = environmentProvider.get(approvalFieldsFileVariableName);
  const environmentName =
    environmentProvider.get(painlessConfigEnvironmentVariableName) ||
    environmentProvider.get('ENV');
  let approvalFields = undefined;
  if (fieldsFile) {
    // Environment approach 1 (legacy):
    // Look for the approval fields file and use that for the approval data
    try {
      const filename = path.join(
        typescriptConfig.appDirectory,
        'data',
        `${fieldsFile}.json`
      );
      const str = fs.readFileSync(filename, 'utf8');
      approvalFields = JSON.parse(str);
      debug(`repo approval types loaded from file ${filename}`);
    } catch (notFound) {
      /* no action required */
      console.warn(notFound);
    }
  } else if (pkg && pkg[painlessConfigEnvPkgName] && environmentName) {
    // Painless config environment approach 2 (newer):
    // Uses the painless config environment + separate env type to get the data
    // This is also a partial hack; if there are multiple environments, this will fail.
    try {
      approvalFields = require(pkg[painlessConfigEnvPkgName])(
        environmentName,
        repoApprovalsEnvironmentName
      );
      debug(
        `repo approval types loaded from painlessConfigEnvPkgName/${environmentName},${repoApprovalsEnvironmentName}`
      );
    } catch (painlessConfigError) {
      debug(
        `attempted to load repo approval types loaded from painlessConfigEnvPkgName/${environmentName},${repoApprovalsEnvironmentName}`
      );
      console.warn(painlessConfigError);
      throw painlessConfigError;
    }
  }

  return {
    repo: arrayFromString(
      environmentProvider.get('REPO_APPROVAL_TYPES') || 'github'
    ),
    teamJoin: arrayFromString(
      environmentProvider.get('TEAM_JOIN_APPROVAL_TYPES') || 'github'
    ),
    fields: approvalFields,
  };
};
