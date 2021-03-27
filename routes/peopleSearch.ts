//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import {
  RequestWithSystemwidePermissions,
  RequestTeamMemberAddType,
  getProviders,
} from '../transitional';

import {
  ensureAllLinksInMemory,
  getAllLinksFromRequest,
} from '../middleware/business/allLinks';

import lowercaser from '../middleware/lowercaser';

import {
  Operations,
  ICrossOrganizationMembersResult,
} from '../business/operations';
import { MemberSearch } from '../business';
import { Team } from '../business';
import { TeamMember } from '../business';
import { OrganizationMember } from '../business';
import { Organization } from '../business';

interface IPeopleSearchRequest extends RequestWithSystemwidePermissions {
  organization?: any;
  team2?: any;
  cachedLinks?: any;
  team2AddType?: RequestTeamMemberAddType;
  team2RemoveType?: any;
  teamUrl?: any;
  teamPermissions?: any;
}

interface IOptionalFilter {
  filter?: string;
}

interface IGetPeopleResult {
  organizationMembers?: OrganizationMember[];
  crossOrganizationMembers?: ICrossOrganizationMembersResult;
  teamMembers?: TeamMember[];
}

router.use(asyncHandler(ensureAllLinksInMemory));

async function getPeopleForOrganization(
  operations: Operations,
  org: string | null,
  options,
  team2: Team
): Promise<IGetPeopleResult> {
  const organization = operations.getOrganization(org);
  const organizationMembers = await organization.getMembers(options);
  if (team2) {
    const teamMembers = await team2.getMembers();
    return { teamMembers, organizationMembers };
  } else {
    return { organizationMembers };
  }
}

async function getPeopleAcrossOrganizations(
  operations: Operations,
  options,
  team2: Team
): Promise<IGetPeopleResult> {
  const crossOrganizationMembers = await operations.getMembers(options);
  if (team2) {
    const teamMembers = await team2.getMembers();
    return { teamMembers, crossOrganizationMembers };
  } else {
    return { crossOrganizationMembers };
  }
}

router.get(
  '/',
  lowercaser(['sort']),
  asyncHandler(async (req: IPeopleSearchRequest, res, next) => {
    const linksFromMiddleware = getAllLinksFromRequest(req);
    const { operations } = getProviders(req);
    const org = req.organization ? req.organization.name : null;
    const orgId = req.organization
      ? (req.organization as Organization).id
      : null;
    const isPortalSudoer =
      req.systemWidePermissions &&
      req.systemWidePermissions.allowAdministration === true;
    let twoFactor = req.query.twoFactor;
    const team2 = req.team2 as Team;
    let options: IOptionalFilter = {};
    if (twoFactor === 'off') {
      options.filter = '2fa_disabled';
    }
    const { crossOrganizationMembers, organizationMembers, teamMembers } = org
      ? await getPeopleForOrganization(operations, org, options, team2)
      : await getPeopleAcrossOrganizations(operations, options, team2);
    const page = req.query.page_number ? Number(req.query.page_number) : 1;
    let phrase = req.query.q as string;
    let type = req.query.type as string;
    const validTypes = new Set([
      'linked',
      'active',
      'unlinked',
      'former',
      'serviceAccount',
      'unknownAccount',
      'owners',
    ]);
    if (!validTypes.has(type)) {
      type = null;
    }
    if (/*twoFactor !== 'on' && */ twoFactor !== 'off') {
      twoFactor = null;
    }
    const filters = [];
    if (type) {
      filters.push({
        type: 'type',
        value: type,
        displayValue: type === 'former' ? 'formerly known' : type,
        displaySuffix: 'members',
      });
    }
    if (phrase) {
      filters.push({
        type: 'phrase',
        value: phrase,
        displayPrefix: 'matching',
      });
    }
    if (twoFactor) {
      filters.push({
        type: 'twoFactor',
        value: twoFactor,
        displayValue: twoFactor === 'on' ? 'secured' : 'without 2fa',
        // displayPrefix: 'matching',
      });
    }
    const search = new MemberSearch({
      phrase,
      type,
      links: linksFromMiddleware,
      providers: operations.providers,
      orgId,
      organizationMembers,
      crossOrganizationMembers,
      isOrganizationScoped: !!org, // Whether this view is specific to an org or not
      team2AddType: req.team2AddType, // Used to enable the "add a member" or maintainer experience for teams
      teamMembers, // Used to filter team members in ./org/ORG/team/TEAM/members and other views
    });

    await search.search(page, req.query.sort as string);
    let maillist = '';
    search.members.forEach(function (element) {
      if (maillist != '' && element.link != undefined) {
        maillist = maillist + ', ';
      }
      maillist += (element.link && element.link.corporateUsername) || '';
    });
    req.individualContext.webContext.render({
      view: 'people/',
      title: 'People',
      state: {
        search,
        filters,
        query: {
          phrase,
          twoFactor,
          type,
        },
        organization: req.organization || undefined,
        lightupSudoerLink: type === 'former' && isPortalSudoer,
        reposDataAgeInformation: null, // CUT: ageInformation,
        team2,
        team2AddType: req.team2AddType,
        team2RemoveType: req.team2RemoveType,
        teamUrl: req.teamUrl,
        specificTeamPermissions: req.teamPermissions,
        maillist,
        operations,
      },
    });
  })
);

export default router;
