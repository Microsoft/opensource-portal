//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

// SPECIALIZED FILE
// This is intended to allow a specific set of defined guest accounts to use the site,
// designed to allow our operations vendor to have access.

const arrayFromString = require('./utils/arrayFromString');

const speciallyAuthorizedGuestUserIdsEnvironmentName =
  'AAD_AUTHORIZED_GUEST_IDS';

module.exports = function (graphApi) {
  const environmentProvider = graphApi.environment;
  const speciallyAuthorizedIdsValue = environmentProvider.get(
    speciallyAuthorizedGuestUserIdsEnvironmentName
  );

  return {
    authorizedIds: arrayFromString(speciallyAuthorizedIdsValue || ''),
  };
};
