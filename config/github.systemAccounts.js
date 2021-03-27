//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn"] }] */

'use strict';

const arrayFromString = require('./utils/arrayFromString');

module.exports = (graphApi) => {
  const environmentProvider = graphApi.environment;
  const systemAccounts = arrayFromString(
    environmentProvider.get('GITHUB_SYSTEM_ACCOUNT_USERNAMES')
  );
  return {
    logins: systemAccounts,
  };
};
