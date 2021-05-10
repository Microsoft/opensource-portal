//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import { ReposAppRequest } from '../interfaces';
const router: Router = Router();

router.get('/', function (req: ReposAppRequest, res) {
  req.reposContext = req.reposContext || {};
  req.reposContext.releaseTab = true;
  req.individualContext.webContext.render({
    view: './emberApp',
    title: 'Releases',
  });
});

export default router;
