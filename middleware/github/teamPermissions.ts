//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ReposAppRequest } from '../../transitional';
import {
  GitHubTeamRole,
  ITeamMembershipRoleState,
  Team,
} from '../../business/team';
import { IndividualContext } from '../../user';
import { OrganizationMembershipState } from '../../business/organization';

// --- team2 context

const contextualTeamContextKey = 'team2';

export function setContextualTeam(req: ReposAppRequest, team: Team) {
  req[contextualTeamContextKey] = team;
}

export function getContextualTeam(req: ReposAppRequest) {
  return req[contextualTeamContextKey] as Team;
}

// --- team membership

const teamStatusCacheKeyName = '$teamMembershipStatus';

export interface IRequestTeamMembershipStatus {
  isLinked: boolean;
  membershipStatus: GitHubTeamRole;
  membershipState: OrganizationMembershipState;
}

export function getTeamMembershipFromRequest(req: ReposAppRequest) {
  return req[teamStatusCacheKeyName] as IRequestTeamMembershipStatus;
}

export async function AddTeamMembershipToRequest(
  req: ReposAppRequest,
  res,
  next
) {
  if (req[teamStatusCacheKeyName]) {
    return next();
  }
  const team2 = req['team2'] as Team;
  if (!team2) {
    return next(new Error('team2 required'));
  }
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  if (!activeContext.link) {
    const noLink: IRequestTeamMembershipStatus = {
      membershipStatus: null,
      membershipState: null,
      isLinked: false,
    };
    req[teamStatusCacheKeyName] = noLink;
  } else {
    const login = activeContext.getGitHubIdentity().username;
    try {
      const statusResult = await team2.getMembershipEfficiently(login);
      const value: IRequestTeamMembershipStatus = {
        membershipStatus:
          statusResult && (statusResult as ITeamMembershipRoleState).role
            ? (statusResult as ITeamMembershipRoleState).role
            : null,
        membershipState:
          statusResult && (statusResult as ITeamMembershipRoleState).state
            ? (statusResult as ITeamMembershipRoleState).state
            : null,
        isLinked: true,
      };
      req[teamStatusCacheKeyName] = value;
    } catch (problem) {
      console.dir(problem);
    }
  }
  return next();
}

// -- team permissions

const teamPermissionsCacheKeyName = 'teamPermissions';

export interface IRequestTeamPermissions {
  allowAdministration: boolean;
  maintainer: boolean;
  sudo: boolean;
  isLinked: boolean;
}

export function getTeamPermissionsFromRequest(req: ReposAppRequest) {
  return req[teamPermissionsCacheKeyName] as IRequestTeamPermissions;
}

export async function AddTeamPermissionsToRequest(
  req: ReposAppRequest,
  res,
  next
) {
  if (req[teamPermissionsCacheKeyName]) {
    return next();
  }
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  const teamPermissions: IRequestTeamPermissions = {
    isLinked: false,
    allowAdministration: false,
    maintainer: false,
    sudo: false,
  };
  req[teamPermissionsCacheKeyName] = teamPermissions;
  if (activeContext.link) {
    teamPermissions.isLinked = true;
    const login = activeContext.getGitHubIdentity().username;
    const organization = req.organization;
    if (!organization) {
      return next(new Error('organization required'));
    }
    const isSudoer = await organization.isSudoer(login, activeContext.link);
    const isPortalSudoer = await activeContext.isPortalAdministrator();

    // Indicate that the user is has sudo rights
    if (isSudoer === true || isPortalSudoer === true) {
      teamPermissions.sudo = true;
    }
  }

  // Get the team maintainers
  const team2 = req['team2'] as Team;
  const maintainers = await team2.getMaintainers();

  // +MIDDLEWARE: providing this later to speed up getting this data
  req['teamMaintainers'] = maintainers;

  if (activeContext.link) {
    const id = activeContext.getGitHubIdentity().id;
    for (let i = 0; i < maintainers.length; i++) {
      if (String(maintainers[i].id) === id) {
        teamPermissions.maintainer = true;
        break;
      }
    }
  }

  // Make a permission decision
  if (teamPermissions.maintainer || teamPermissions.sudo) {
    teamPermissions.allowAdministration = true;
  }
  return next();
}
