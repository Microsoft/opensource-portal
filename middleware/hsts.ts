//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import hsts from 'hsts';

export default hsts({
  maxAge: 10886400000, // Must be at least 18 weeks to be approved
  includeSubDomains: true, // Must be enabled to be approved
  preload: true,
});
