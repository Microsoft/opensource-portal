//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';

import { jsonError } from '../../middleware';
import { getProviders, ReposAppRequest } from '../../transitional';

const router = express.Router();

// TODO: move to modern w/administration experience, optionally

router.get('/', (req: ReposAppRequest, res, next) => {
  const { config } = getProviders(req);
  const text = config?.serviceMessage?.banner || null;
  const link = config.serviceMessage?.link;
  const details = config.serviceMessage?.details;
  let banner = text ? { text, link, details } : null;
  return res.json({ banner });
});

router.use('*', (req, res, next) => {
  return next(
    jsonError('no API or function available within this banner route', 404)
  );
});

export default router;
