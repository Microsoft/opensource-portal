//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import { Repository } from '../../business';
import { ReposAppRequest } from '../../interfaces';
import { jsonError } from '../../middleware';
import { getProviders } from '../../transitional';
import JsonPager from './jsonPager';
import { RepositorySearchSortOrder, searchRepos } from './organization/repos';

const router: Router = Router();

router.get('/', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const providers = getProviders(req);
  const pager = new JsonPager<Repository>(req, res);
  const searchOptions = {
    q: (req.query.q || '') as string,
    type: (req.query.type || '') as string, // CONSIDER: TS: stronger typing
  }
  try {
    const repos = await searchRepos(providers, null, RepositorySearchSortOrder.Updated, searchOptions);
    const slice = pager.slice(repos);
    return pager.sendJson(slice.map(repo => {
      return repo.asJson();
    }));
  } catch (repoError) {
    console.dir(repoError);
    return next(jsonError(repoError));
  }
}));

router.use('*', (req, res, next) => {
  return next(jsonError('no API or function available within this cross-organization repps list', 404));
});

export default router;
