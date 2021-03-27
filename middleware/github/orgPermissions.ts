//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ReposAppRequest } from '../../transitional';
import { wrapError } from '../../utils';
import { OrganizationMembershipState } from '../../business/organization';

const orgPermissionsCacheKeyName = 'orgPermissions';
const orgOwnersCacheKeyName = 'orgOwners';
const orgOwnersSetCacheKeyName = 'orgOwnersSet';

export interface IRequestOrganizationPermissions {
  allowAdministration: boolean;
  owner: boolean;
  sudo: boolean;
  membershipStatus: OrganizationMembershipState;
}

export function GetOrganizationPermissionsFromRequest(req: ReposAppRequest) {
  return req[orgPermissionsCacheKeyName];
}

export async function AddOrganizationPermissionsToRequest(
  req: ReposAppRequest,
  res,
  next
) {
  // Only compute once per request
  if (req[orgPermissionsCacheKeyName]) {
    return next();
  }
  const individualContext = req.individualContext;
  const login = individualContext.getGitHubIdentity().username;
  const ghIdAsString = individualContext.getGitHubIdentity().id;
  const id = ghIdAsString ? parseInt(ghIdAsString, 10) : null;
  const organization = req.organization;
  const orgPermissions: IRequestOrganizationPermissions = {
    allowAdministration: false,
    owner: false,
    sudo: false,
    membershipStatus: null,
  };
  if (id && !login) {
    return next(
      new Error(
        `While your technical GitHub ID ${id} is known, your GitHub username is not currently known.`
      )
    );
  }
  req[orgPermissionsCacheKeyName] = orgPermissions;
  const isSudoer = await organization.isSudoer(login, individualContext.link);
  const isPortalSudoer = await individualContext.isPortalAdministrator();

  // Indicate that the user is has sudo rights
  if (isSudoer === true || isPortalSudoer === true) {
    orgPermissions.sudo = true;
  }

  // Get the organization owners
  const owners = await organization.getOwners();

  // +MIDDLEWARE: provide this later if it is needed elsewhere
  req[orgOwnersCacheKeyName] = owners;
  const set = new Set();
  for (let i = 0; i < owners.length; i++) {
    set.add(owners[i].id);
  }
  if (set.has(id)) {
    orgPermissions.owner = true;
  }
  req[orgOwnersSetCacheKeyName] = set;

  // Make a permission decision
  if (orgPermissions.owner || orgPermissions.sudo) {
    orgPermissions.allowAdministration = true;
  }

  // Are they even an organization member?
  const membershipCacheOptions = {
    maxAgeSeconds: 30,
    backgroundRefresh: false,
  };

  try {
    const membershipStatus = await organization.getMembership(
      login,
      membershipCacheOptions
    );
    orgPermissions.membershipStatus =
      membershipStatus && membershipStatus.state
        ? membershipStatus.state
        : null;
    return next();
  } catch (getMembershipError) {
    // if (getMembershipError && getMembershipError.innerError && getMembershipError.innerError.status === 404) {
    //   getMembershipError = null;
    //   membershipStatus = null;
    // }
    const reason = getMembershipError.message;
    return next(
      wrapError(
        getMembershipError,
        `Unable to successfully validate whether you are already a member of the ${organization.name} organization on GitHub. ${reason}`
      )
    );
  }
}
