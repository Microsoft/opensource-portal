//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

export default function (params) {
  return function (req, res, next) {
    // lowercase parameters
    Object.getOwnPropertyNames(req.params).forEach((param) => {
      if (params.indexOf(param) > -1) {
        req.params[param] = lowerCaser(req.params[param]);
      }
    });
    // lowercase query string
    Object.getOwnPropertyNames(req.query).forEach((query) => {
      if (params.indexOf(query) > -1) {
        req.query[query] = lowerCaser(req.query[query]);
      }
    });
    next();
  };
}

function lowerCaser(param) {
  if (typeof param === 'string') {
    return param.toLowerCase();
  }
  if (Array.isArray(param)) {
    return param.map((item) => {
      return item.toLowerCase();
    });
  }
  return param;
}
