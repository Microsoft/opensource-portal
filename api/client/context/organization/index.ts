//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
import {
  GitHubTeamRole,
  Organization,
  OrganizationMembershipRole,
  Team,
  TeamJsonFormat,
} from '../../../../business';
import { jsonError } from '../../../../middleware';
import getCompanySpecificDeployment from '../../../../middleware/companySpecificDeployment';
import { ReposAppRequest } from '../../../../transitional';
import { IndividualContext } from '../../../../user';

import RouteRepos from './repos';
import RouteTeams from './teams';

const router = express.Router();

router.get(
  '/',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { organization } = req;
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    if (!activeContext.link) {
      return res.json(false);
    }
    const membership = await organization.getOperationalMembership(
      activeContext.getGitHubIdentity().username
    );
    if (!membership) {
      return res.json(false);
    }
    return res.json({
      user: toSanitizedUser(membership.user),
      organization: toSanitizedOrg(membership.organization),
      role: membership.role,
      state: membership.state,
    });
  })
);

router.get(
  '/sudo',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { organization } = req;
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    if (!activeContext.link) {
      return res.json({ isSudoer: false });
    }
    return res.json({
      isSudoer: await organization.isSudoer(
        activeContext.getGitHubIdentity().username,
        activeContext.link
      ),
    });
  })
);

router.get(
  '/isOwner',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { organization } = req;
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    if (!activeContext.link) {
      return res.json({ isOrganizationOwner: false });
    }
    try {
      const username = activeContext.getGitHubIdentity().username;
      const membership = await organization.getOperationalMembership(username);
      const isOrganizationOwner =
        membership?.role === OrganizationMembershipRole.Admin;
      return res.json({
        isOrganizationOwner,
      });
    } catch (error) {
      return jsonError(error, 400);
    }
  })
);

router.delete(
  '/',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    // "Leave" / remove my context
    const { organization } = req;
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    if (!activeContext.link) {
      return next(jsonError('You are not linked', 400));
    }
    const login = activeContext.getGitHubIdentity().username;
    const id = activeContext.getGitHubIdentity().id;
    try {
      await organization.removeMember(login, id);
      return res.json({
        message: `Your ${login} account has been removed from ${organization.name}.`,
      });
    } catch (error) {
      console.warn(error);
      return next(jsonError(error, 400));
    }
  })
);

router.get(
  '/personalizedTeams',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    try {
      const organization = req.organization as Organization;
      const activeContext = (req.individualContext ||
        req.apiContext) as IndividualContext;
      if (!activeContext.link) {
        return res.json({ personalizedTeams: [] });
      }
      const userAggregateContext = activeContext.aggregations;
      const maintainedTeams = new Set<string>();
      const broadTeams = new Set<number>(req.organization.broadAccessTeams);
      const userTeams = userAggregateContext.reduceOrganizationTeams(
        organization,
        await userAggregateContext.teams()
      );
      userTeams.maintainer.map((maintainedTeam) =>
        maintainedTeams.add(maintainedTeam.id.toString())
      );
      const combinedTeams = new Map<string, Team>();
      userTeams.maintainer.map((team) =>
        combinedTeams.set(team.id.toString(), team)
      );
      userTeams.member.map((team) =>
        combinedTeams.set(team.id.toString(), team)
      );
      const personalizedTeams = Array.from(combinedTeams.values()).map(
        (combinedTeam) => {
          const entry = combinedTeam.asJson(TeamJsonFormat.Augmented);
          entry.role = maintainedTeams.has(combinedTeam.id.toString())
            ? GitHubTeamRole.Maintainer
            : GitHubTeamRole.Member;
          return entry;
        }
      );
      return res.json({
        personalizedTeams,
      });
    } catch (error) {
      return next(jsonError(error, 400));
    }
  })
);

router.use('/repos', RouteRepos);
router.use('/teams', RouteTeams);

const deployment = getCompanySpecificDeployment();
deployment?.routes?.api?.context?.organization?.index &&
  deployment?.routes?.api?.context?.organization?.index(router);

router.use('*', (req, res, next) => {
  return next(
    jsonError('no API or function available: client>organization', 404)
  );
});

const toSanitizedUser = (user) => {
  if (!user || !user.login) {
    return undefined;
  }
  return {
    id: user.id,
    login: user.login,
    avatar_url: user.avatar_url,
  };
};
const toSanitizedOrg = (org) => {
  if (!org || !org.login) {
    return undefined;
  }
  return {
    id: org.id,
    login: org.login,
    avatar_url: org.avatar_url,
    description: org.description,
  };
};

export default router;
