//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// I no longer believe this file FYI... I don't think we need this any longer.

import { getProviders } from '../transitional';

// Office uses a specialized pre-fetch to learn more about hyperlinks before
// opening. As a result, if the Office user agent is in use, and the
// request is not authenticated, we should still just respond with a simple
// HTTP 200 and message. Office technically creates a HEAD request. Once a
// successful response is found, it opens a web browser to the specific
// URL. This middleware performs these actions.

const GenericOfficeUserAgent = 'ms-office';
const WordUserAgent = 'Microsoft Office Word';

export default function supportOfficeHyperlinks(req, res, next) {
  const { insights } = getProviders(req);
  const userAgent = req.headers['user-agent'];
  const isAuthenticated = req.isAuthenticated ? req.isAuthenticated() : false;
  if (
    userAgent &&
    userAgent.includes &&
    !isAuthenticated &&
    (userAgent.includes(GenericOfficeUserAgent) ||
      userAgent.includes(WordUserAgent))
  ) {
    insights?.trackEvent({
      name: 'InterceptOfficeHyperlinkRequest',
      properties: {
        userAgent: userAgent,
        isAuthenticated: isAuthenticated,
        originalUrl: req.originalUrl,
        httpMethod: req.method,
        responseType: 200,
      },
    });
    return res.send(
      `When using Microsoft Office, you need to open the hyperlink in your browser to authenticate if needed`
    );
  }
  return next();
}
