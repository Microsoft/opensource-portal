//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';

import { jsonError } from '../../../middleware';
import {
  ErrorHelper,
  getProviders,
  LocalApiRepoAction,
  ReposAppRequest,
} from '../../../transitional';
import { Repository } from '../../../business';
import { IndividualContext } from '../../../user';
import NewRepositoryLockdownSystem from '../../../features/newRepositoryLockdown';
import {
  AddRepositoryPermissionsToRequest,
  getContextualRepositoryPermissions,
} from '../../../middleware/github/repoPermissions';
import { renameRepositoryDefaultBranchEndToEnd } from '../../../routes/org/repos';
import getCompanySpecificDeployment from '../../../middleware/companySpecificDeployment';

import RouteRepoPermissions from './repoPermissions';

type RequestWithRepo = ReposAppRequest & {
  repository: Repository;
};

const router = express.Router();

const deployment = getCompanySpecificDeployment();
deployment?.routes?.api?.organization?.repo &&
  deployment?.routes?.api?.organization?.repo(router);

router.use('/permissions', RouteRepoPermissions);

router.get(
  '/',
  asyncHandler(async (req: RequestWithRepo, res, next) => {
    const { repository } = req;
    try {
      await repository.getDetails();

      const clone = Object.assign({}, repository.getEntity());
      delete clone.temp_clone_token; // never share this back
      delete clone.cost;

      return res.json(repository.getEntity());
    } catch (repoError) {
      if (ErrorHelper.IsNotFound(repoError)) {
        // // Attempt fallback by ID (?)
      }
      return next(jsonError(repoError));
    }
  })
);

router.get(
  '/exists',
  asyncHandler(async (req: RequestWithRepo, res, next) => {
    let exists = false;
    let name: string = undefined;
    const { repository } = req;
    try {
      const originalName = repository.name;
      await repository.getDetails();
      if (repository && repository.name) {
        name = repository.getEntity().name as string;
        if (name.toLowerCase() !== originalName.toLowerCase()) {
          // A renamed repository will return the new name here
          exists = false;
        } else {
          exists = true;
        }
      }
    } catch (repoError) {}
    return res.json({ exists, name });
  })
);

router.patch(
  '/renameDefaultBranch',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: RequestWithRepo, res, next) {
    const providers = getProviders(req);
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    const repoPermissions = getContextualRepositoryPermissions(req);
    const targetBranchName = req.body.default_branch;
    const { repository } = req;
    try {
      const result = await renameRepositoryDefaultBranchEndToEnd(
        providers,
        activeContext,
        repoPermissions,
        repository,
        targetBranchName,
        true /* wait for refresh before sending response */
      );
      return res.json(result);
    } catch (error) {
      return next(jsonError(error));
    }
  })
);

router.post(
  '/archive',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: RequestWithRepo, res, next) {
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    const providers = getProviders(req);
    const { insights } = providers;
    const repoPermissions = getContextualRepositoryPermissions(req);
    if (!repoPermissions.allowAdministration) {
      return next(
        jsonError('You do not have permission to archive this repo', 403)
      );
    }
    const insightsPrefix = 'ArchiveRepo';
    const { repository } = req;
    try {
      insights?.trackEvent({
        name: `${insightsPrefix}Started`,
        properties: {
          requestedById: activeContext.link.corporateId,
          repoName: repository.name,
          orgName: repository.organization.name,
          repoId: repository.id ? String(repository.id) : 'unknown',
        },
      });
      const currentRepositoryState = deployment?.features?.repositoryActions
        ?.getCurrentRepositoryState
        ? await deployment.features.repositoryActions.getCurrentRepositoryState(
            providers,
            repository
          )
        : null;
      await repository.archive();
      if (deployment?.features?.repositoryActions?.sendActionReceipt) {
        deployment.features.repositoryActions
          .sendActionReceipt(
            providers,
            activeContext,
            repository,
            LocalApiRepoAction.Archive,
            currentRepositoryState
          )
          .then((ok) => {})
          .catch(() => {});
      }
      insights?.trackMetric({
        name: `${insightsPrefix}s`,
        value: 1,
      });
      insights?.trackEvent({
        name: `${insightsPrefix}Success`,
        properties: {
          requestedById: activeContext.link.corporateId,
          repoName: repository.name,
          orgName: repository.organization.name,
          repoId: repository.id ? String(repository.id) : 'unknown',
        },
      });
      //return res.json(result);
      return res.json({
        message: `You archived: ${repository.full_name}`,
      });
    } catch (error) {
      insights?.trackException({ exception: error });
      insights?.trackEvent({
        name: `${insightsPrefix}Failed`,
        properties: {
          requestedById: activeContext.link.corporateId,
          repoName: repository.name,
          orgName: repository.organization.name,
          repoId: repository.id ? String(repository.id) : 'unknown',
        },
      });
      return next(jsonError(error));
    }
  })
);

router.delete(
  '/',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: RequestWithRepo, res, next) {
    // NOTE: duplicated code from /routes/org/repos.ts
    const providers = getProviders(req);
    const { insights } = providers;
    const insightsPrefix = 'DeleteRepo';
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    const { organization, repository } = req;
    const repoPermissions = getContextualRepositoryPermissions(req);
    if (repoPermissions.allowAdministration) {
      try {
        insights?.trackEvent({
          name: `${insightsPrefix}Started`,
          properties: {
            requestedById: activeContext.link.corporateId,
            repoName: repository.name,
            orgName: repository.organization.name,
            repoId: repository.id ? String(repository.id) : 'unknown',
          },
        });
        const currentRepositoryState = deployment?.features?.repositoryActions
          ?.getCurrentRepositoryState
          ? await deployment.features.repositoryActions.getCurrentRepositoryState(
              providers,
              repository
            )
          : null;
        await repository.delete();
        if (deployment?.features?.repositoryActions?.sendActionReceipt) {
          deployment.features.repositoryActions
            .sendActionReceipt(
              providers,
              activeContext,
              repository,
              LocalApiRepoAction.Delete,
              currentRepositoryState
            )
            .then((ok) => {})
            .catch(() => {});
        }
        insights?.trackMetric({
          name: `${insightsPrefix}s`,
          value: 1,
        });
        insights?.trackEvent({
          name: `${insightsPrefix}Success`,
          properties: {
            requestedById: activeContext.link.corporateId,
            repoName: repository.name,
            orgName: repository.organization.name,
            repoId: repository.id ? String(repository.id) : 'unknown',
          },
        });
        return res.json({
          message: `You deleted: ${repository.full_name}`,
        });
      } catch (error) {
        insights?.trackException({ exception: error });
        insights?.trackEvent({
          name: `${insightsPrefix}Failed`,
          properties: {
            requestedById: activeContext.link.corporateId,
            repoName: repository.name,
            orgName: repository.organization.name,
            repoId: repository.id ? String(repository.id) : 'unknown',
          },
        });
        return next(jsonError(error));
      }
    }
    if (!organization.isNewRepositoryLockdownSystemEnabled) {
      return next(
        jsonError(
          'This endpoint is not available as configured in this app.',
          400
        )
      );
    }
    const daysAfterCreateToAllowSelfDelete = 21; // could be a config setting if anyone cares
    try {
      // make sure ID is known
      if (await repository.isDeleted()) {
        return next(jsonError('The repository has already been deleted', 404));
      }
      const metadata = await repository.getRepositoryMetadata();
      await NewRepositoryLockdownSystem.ValidateUserCanSelfDeleteRepository(
        repository,
        metadata,
        activeContext,
        daysAfterCreateToAllowSelfDelete
      );
    } catch (noExistingMetadata) {
      if (noExistingMetadata.status === 404) {
        return next(
          jsonError(
            'This repository does not have any metadata available regarding who can setup it up. No further actions available.',
            400
          )
        );
      }
      return next(jsonError(noExistingMetadata, 404));
    }
    const { operations, repositoryMetadataProvider } = getProviders(req);
    const lockdownSystem = new NewRepositoryLockdownSystem({
      operations,
      organization,
      repository,
      repositoryMetadataProvider,
    });
    await lockdownSystem.deleteLockedRepository(
      false /* delete for any reason */,
      true /* deleted by the original user instead of ops */
    );
    return res.json({
      message: `You deleted your repo, ${repository.full_name}.`,
    });
  })
);

router.use('*', (req, res, next) => {
  console.warn(req.baseUrl);
  return next(
    jsonError('no API or function available within this specific repo', 404)
  );
});

export default router;
