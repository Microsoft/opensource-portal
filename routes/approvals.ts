//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
const router: Router = Router();

router.get('/', (req, res) => {
  return res.redirect('/settings/approvals');
});

router.get('/:id', (req, res) => {
  const id = req.params.id;
  return res.redirect(`/settings/approvals/${id}`);
});

export default router;
