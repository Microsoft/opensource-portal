//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import { ICorporateLink, UnlinkPurpose } from '../../interfaces';
import { jsonError } from '../../middleware';
import { IApiRequest } from '../../middleware/apiReposAuth';
import { getProviders } from '../../transitional';

const router: Router = Router();

interface ILinksApiRequestWithUnlink extends IApiRequest {
  unlink?: ICorporateLink;
}

router.use(function (req: ILinksApiRequestWithUnlink, res, next) {
  const token = req.apiKeyToken;
  if (!token.scopes) {
    return next(jsonError('The key is not authorized for specific APIs', 401));
  }
  if (!token.hasScope('unlink')) {
    return next(jsonError('The key is not authorized to use the unlink API', 401));
  }
  return next();
});

router.use('/github/id/:id', asyncHandler(async (req: ILinksApiRequestWithUnlink, res, next) => {
  const { linkProvider } = getProviders(req);
  const id = req.params.id;
  try {
    const link = await linkProvider.getByThirdPartyId(id);
    if (!link) {
      throw new Error(`Could not locate a link for GitHub user ID ${id}`);
    }
    req.unlink = link;
    return next();
  } catch (error) {
    return next(jsonError(error));
  }
}));

router.use('*', (req: ILinksApiRequestWithUnlink, res, next) => {
  return next(req.unlink ? undefined : jsonError('No link available for operation', 404));
});

router.delete('*', (req: ILinksApiRequestWithUnlink, res, next) => {
  const { operations } = getProviders(req);
  const link = req.unlink;
  let purpose: UnlinkPurpose = null;
  try {
    purpose = apiUnlinkPurposeToEnum((req.headers['unlink-purpose'] || 'termination') as string);
  } catch (purposeError) {
    return next(jsonError(purposeError, 400));
  }
  const options = { purpose };
  return operations.terminateLinkAndMemberships(link.thirdPartyId, options).then(results => {
    res.json({
      messages: Array.isArray(results) ? (results as any as string[]).reverse() : results,
    });
  }).catch(problem => {
    return next(jsonError(problem, 500));
  });
});

function apiUnlinkPurposeToEnum(purpose: string): UnlinkPurpose {
  switch (purpose) {
    case 'operations':
      throw new Error('The unlink purpose "operations" is not supported by API');
    case 'termination':
      return UnlinkPurpose.Termination;
    case 'self':
      throw new Error('The unlink purpose "self" is not supported by API');
    default:
      return UnlinkPurpose.Unknown;
  }
}

export default router;
