//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// This route does not use the portal administrator team but instead an explicit
// approved list of corporate usernames.

import { getProviders, ReposAppRequest } from '../../transitional';

import { wrapError } from '../../utils';

function denyRoute(next) {
  next(
    wrapError(
      null,
      "These aren't the droids you are looking for. You do not have permission to be here.",
      true
    )
  );
}

export function AuthorizeOnlyCorporateAdministrators(
  req: ReposAppRequest,
  res,
  next
) {
  const individualContext = req.individualContext;
  const config = getProviders(req).config;
  const administrators: string[] =
    config && config.administrators && config.administrators.corporateUsernames
      ? config.administrators.corporateUsernames
      : null;
  const username = individualContext.corporateIdentity.username;
  let isAuthorized = false;
  if (
    administrators &&
    username &&
    administrators.includes(username.toLowerCase())
  ) {
    return next();
  }
  res.header('x-username', username);
  return denyRoute(next);
}
