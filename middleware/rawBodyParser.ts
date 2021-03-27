//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

export default function rawBodyParser(req, res, next) {
  // Since we have a site-wide implementation of body parser, this allows
  // routines access to the raw body if it is needed
  req._raw = '';
  req.on('data', (chunk) => {
    req._raw += chunk;
  });

  next();
}
