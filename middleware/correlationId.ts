//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { v4 as uuidV4 } from 'uuid';

// ----------------------------------------------------------------------------
// Generate a correlation ID
// ----------------------------------------------------------------------------
module.exports = function (req, res, next) {
  req.correlationId = uuidV4();
  next();
};
