//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
const router: Router = Router();

import { ReposAppRequest } from '../interfaces';
import { getProviders } from '../transitional';

// These are Microsoft-specific, we'll remove these eventually.
// TODO: remove from open source version since not helpful having random routes in place

router.use('/data', (req: ReposAppRequest, res) => {
  const { config } = getProviders(req);
  const exploreUrl = config?.urls?.explore;
  res.redirect(exploreUrl ? `${exploreUrl}resources/insights` : '/');
});

router.use('/use', (req: ReposAppRequest, res) => {
  const { config } = getProviders(req);
  const exploreUrl = config?.urls?.explore;
  res.redirect(exploreUrl ? `${exploreUrl}resources/use` : '/');
});

router.use('/release', (req: ReposAppRequest, res) => {
  const { config } = getProviders(req);
  const exploreUrl = config?.urls?.explore;
  res.redirect(exploreUrl ? `${exploreUrl}resources/release` : '/');
});

export default router;
