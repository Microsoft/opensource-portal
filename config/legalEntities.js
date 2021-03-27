//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

const arrayFromString = require('./utils/arrayFromString');

const envName = 'LEGAL_ENTITIES';
const defaultEnvName = 'GITHUB_ORGANIZATIONS_DEFAULT_LEGAL_ENTITIES';

module.exports = function (graphApi) {
  const environmentProvider = graphApi.environment;
  const value = environmentProvider.get(envName);

  return {
    entities: arrayFromString(value || ''),
    defaultOrganizationEntities: arrayFromString(
      environmentProvider.get(defaultEnvName) || ''
    ),
  };
};
