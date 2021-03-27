//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import { ReposAppRequest, getProviders } from '../../transitional';
import { AddLinkToRequest } from '../../middleware/links/';

import approvalsRoute from './approvals';
import authorizationsRoute from './authorizations';
import digestReportsRoute from './digestReports';
import personalAccessTokensRoute from './personalAccessTokens';

import contributionDataRoute from './contributionData';
import campaignsRoute from './campaigns';

router.use(asyncHandler(AddLinkToRequest));

router.get(
  '/',
  asyncHandler(async (req: ReposAppRequest, res) => {
    const providers = getProviders(req);
    const link = req.individualContext.link;
    let legalContactInformation = null;
    try {
      if (providers.corporateContactProvider) {
        legalContactInformation = await providers.corporateContactProvider.lookupContacts(
          link.corporateUsername
        );
      }
    } catch (ignoredError) {
      /* ignored */
    }
    req.individualContext.webContext.render({
      view: 'settings',
      title: 'Settings',
      state: {
        legalContactInformation,
        link,
      },
    });
  })
);

router.use('/approvals', approvalsRoute);
router.use('/authorizations', authorizationsRoute);
router.use('/campaigns', campaignsRoute);
router.use('/digestReports', digestReportsRoute);
router.use('/security/tokens', personalAccessTokensRoute);
router.use('/contributionData', contributionDataRoute);

export default router;
