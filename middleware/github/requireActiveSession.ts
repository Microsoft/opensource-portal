//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ReposAppRequest } from '../../transitional';
import { storeOriginalUrlAsReferrer } from '../../utils';

export default function RequireActiveGitHubSession(
  req: ReposAppRequest,
  res,
  next
) {
  const identity = req.individualContext.getSessionBasedGitHubIdentity();
  if (!identity) {
    return storeOriginalUrlAsReferrer(req, res, '/signin/github');
  }
  return next();
}
