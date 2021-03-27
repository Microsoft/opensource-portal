//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Scrub the incoming URL value(s) in the request, replacing tokens and other
// secrets.
export default function (req, res, next) {
  var url = req.originalUrl || req.url;
  var secretKeys = ['code', 'token'];
  for (var i = 0; i < secretKeys.length; i++) {
    var key = secretKeys[i];
    var value = req.query[key];
    if (value !== undefined) {
      url = url.replace(key + '=' + value, key + '=*****');
    }
  }
  req.scrubbedUrl = url;
  next();
}
