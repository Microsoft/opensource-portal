//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import os from 'os';

import { getProviders, ReposAppRequest } from '../transitional';

export default function (req: ReposAppRequest, res, next) {
  const { config, viewServices } = getProviders(req);
  req.app.locals.correlationId = req.correlationId;
  req.app.locals.scrubbedUrl = req.scrubbedUrl;
  req.app.locals.serverAddress = req.hostname;
  req.app.locals.serverName = os.hostname();
  req.app.locals.websiteHostname = process.env.WEBSITE_HOSTNAME;
  req.app.locals.appInsightsKey = config?.telemetry?.applicationInsightsKey;
  req.app.locals.googleAnalyticsKey = config?.telemetry?.googleAnalyticsKey;
  req.app.locals.viewServices = viewServices;
  return next();
}
