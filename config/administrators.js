//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

const arrayFromString = require('./utils/arrayFromString');

const administratorsEnvironmentName =
  'AUTHORIZED_CORPORATE_ADMINISTRATOR_USERNAMES';

module.exports = function (graphApi) {
  const environmentProvider = graphApi.environment;
  const value = environmentProvider.get(administratorsEnvironmentName);

  return {
    corporateUsernames: arrayFromString(value || ''),
  };
};
