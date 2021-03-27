//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// This is a Microsoft-specific piece of middleware.

import { ReposAppRequest } from '../../transitional';

export function AuthorizeOnlyFullTimeEmployeesAndInterns(
  req: ReposAppRequest,
  res,
  next
) {
  const individualContext = req.individualContext;
  if (
    !individualContext.corporateIdentity ||
    !individualContext.corporateIdentity.username
  ) {
    return next(
      new Error('This resource is only available to authenticated users.')
    );
  }
  if (isEmployeeOrIntern(individualContext.corporateIdentity.username)) {
    return next();
  }
  return next(
    new Error(
      `This resource is only available to full-time employees and interns at this time. Username: ${individualContext.corporateIdentity.username}`
    )
  );
}

export function isEmployeeOrIntern(upn: string): boolean {
  if (!upn) {
    return false;
  }
  upn = upn.toLowerCase();
  const dashIndex = upn.indexOf('-');
  if (dashIndex < 0 || upn.startsWith('t-')) {
    return true;
  }
  return false;
}
