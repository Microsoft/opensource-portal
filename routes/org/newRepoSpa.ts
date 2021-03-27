//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import { ReposAppRequest, getProviders } from '../../transitional';
import NewRepositoryLockdownSystem from '../../features/newRepositoryLockdown';
import { Organization } from '../../business/organization';

router.get(
  '/',
  asyncHandler(async function (req: ReposAppRequest, res) {
    const providers = getProviders(req);
    const individualContext = req.individualContext;
    const existingRepoId = req.query.existingrepoid as string;
    const organization = req.organization as Organization;
    if (organization.createRepositoriesOnGitHub) {
      throw new Error(
        'This organization requires that repositories are either directly created on GitHub, or by an organization owner.'
      );
    }
    if (existingRepoId && organization.isNewRepositoryLockdownSystemEnabled) {
      try {
        const metadata = await providers.repositoryMetadataProvider.getRepositoryMetadata(
          existingRepoId
        );
        await NewRepositoryLockdownSystem.ValidateUserCanConfigureRepository(
          metadata,
          individualContext
        );
      } catch (noExistingMetadata) {
        if (noExistingMetadata.status === 404) {
          throw new Error(
            'This repository does not have any metadata available regarding who can setup it up. No further actions available.'
          );
        }
        throw noExistingMetadata;
      }
    }
    const orgName = organization.name.toLowerCase();
    req.individualContext.webContext.render({
      view: 'emberApp',
      title: 'New repository',
      state: {
        orgName: orgName,
        organization: organization,
      },
    });
  })
);

export default router;
