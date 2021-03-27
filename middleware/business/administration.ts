//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
const router = express.Router();

import { ReposAppRequest } from '../../transitional';
import { IndividualContext } from '../../user';

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

export function requirePortalAdministrationPermission(
  req: ReposAppRequest,
  res,
  next
) {
  req.individualContext
    .isPortalAdministrator()
    .then((isAdmin) => {
      return isAdmin === true ? next() : denyRoute(next);
    })
    .catch((err) => {
      denyRoute(next);
    });
}
