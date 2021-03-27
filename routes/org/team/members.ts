//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import {
  ReposAppRequest,
  RequestTeamMemberAddType,
  UserAlertType,
  getProviders,
  validateGitHubLogin,
} from '../../../transitional';
import { Team } from '../../../business/team';
import { TeamMember } from '../../../business/teamMember';

import RoutePeopleSearch from '../../peopleSearch';
import MiddlewareTeamAdminRequired from './teamAdminRequired';

interface ILocalTeamRequest extends ReposAppRequest {
  team2?: Team;
  refreshedMembers?: any;
  teamUrl?: any;
  team2AddType?: RequestTeamMemberAddType;
  team2RemoveType?: any;
}

function refreshMembers(
  team2: Team,
  backgroundRefresh: boolean,
  maxSeconds: number,
  firstPageOnly: boolean
): Promise<TeamMember[]> {
  const options = {
    maxAgeSeconds: maxSeconds || 60,
    backgroundRefresh: backgroundRefresh,
    pageLimit: undefined,
  };
  if (firstPageOnly) {
    options.pageLimit = 1;
  }
  return team2.getMembers(options);
}

async function refreshMembersAndSummary(team2: Team, when): Promise<void> {
  await refreshMembers(
    team2,
    false /* immediately refresh */,
    when === 'now' ? -1 : null,
    true /* start with just the first page */
  );
  await refreshMembers(
    team2,
    false /* immediate */,
    when === 'now' ? -1 : null,
    false /* refresh all pages */
  );
}

router.use(
  asyncHandler(async (req: ILocalTeamRequest, res, next) => {
    // Always make sure to have a relatively up-to-date membership cache available
    const team2 = req.team2 as Team;
    req.refreshedMembers = await refreshMembers(
      team2,
      true /* background refresh ok */,
      null,
      false /* refresh all pages */
    );
    return next();
  })
);

router.get(
  '/refresh',
  asyncHandler(async (req: ILocalTeamRequest, res, next) => {
    // Refresh all the pages and also the cached single-page view shown on the team page
    const team2 = req.team2 as Team;
    await refreshMembersAndSummary(team2, 'whenever');
    return res.redirect(req.teamUrl);
  })
);

// Browse members
router.use(
  '/browse',
  (req: ILocalTeamRequest, res, next) => {
    req.team2RemoveType = 'member';
    return next();
  },
  RoutePeopleSearch
);

// Add org members to the team
router.use(
  '/add',
  MiddlewareTeamAdminRequired,
  (req: ILocalTeamRequest, res, next) => {
    req.team2AddType = RequestTeamMemberAddType.Member;
    return next();
  },
  RoutePeopleSearch
);

router.post(
  '/remove',
  MiddlewareTeamAdminRequired,
  asyncHandler(async (req: ILocalTeamRequest, res, next) => {
    const { operations } = getProviders(req);
    const username = validateGitHubLogin(req.body.username);
    const team2 = req.team2 as Team;
    await team2.removeMembership(username);
    req.individualContext.webContext.saveUserAlert(
      `${username} has been removed from the team ${team2.name}.`,
      'Team membership update',
      UserAlertType.Success
    );
    await refreshMembersAndSummary(team2, 'now');
    return res.redirect(`${req.teamUrl}members/browse/`);
  })
);

router.post(
  '/add',
  MiddlewareTeamAdminRequired,
  asyncHandler(async (req: ILocalTeamRequest, res, next) => {
    const { operations } = getProviders(req);
    const username = validateGitHubLogin(req.body.username);
    const organization = req.organization;
    const team2 = req.team2;
    const refreshedMembers = req.refreshedMembers;
    // Allow a one minute org cache for self-correcting validation
    const orgOptions = {
      maxAgeSeconds: 60,
      backgroundRefresh: true,
    };
    // Validate that the user is a current org member
    try {
      const membership = await organization.getMembership(username);
      if (!membership) {
        throw new Error(
          `Membership information in the ${organization.name} organization is unknown for ${username}`
        );
      }
      if (membership.state !== 'active') {
        throw new Error(
          `${username} has the organization state of ${membership.state}. The user is not an active member and so cannot be added to the team at this time.`
        );
      }
    } catch (error) {
      if (error && error.innerError && error.innerError.status === 404) {
        error = new Error(
          `${username} is not a member of the ${organization.name} organization and so cannot be added to the team until they have joined the org.`
        );
      }
      return next(error);
    }

    // Make sure they are not already a member
    const lc = username.toLowerCase();
    for (let i = 0; i < refreshedMembers.length; i++) {
      const member = refreshedMembers[i];
      if (member.login.toLowerCase() === lc) {
        return next(
          new Error(`The user ${username} is already a member of the team.`)
        );
      }
    }
    await team2.addMembership(username);
    req.individualContext.webContext.saveUserAlert(
      `Added ${username} to the ${team2.name} team.`,
      'Team membership update',
      UserAlertType.Success
    );
    await refreshMembersAndSummary(team2, 'now');
    return res.redirect(req.teamUrl + 'members/browse/');
  })
);

export default router;
