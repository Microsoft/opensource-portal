//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { getProviders } from '../../transitional';
import { wrapError } from '../../utils';
import { Organization } from '../../business';
import { OrganizationMembershipState, ReposAppRequest, UserAlertType } from '../../interfaces';

interface IOrganizationMembershipState {
  state: OrganizationMembershipState;
  org: Organization;
}

interface ILocalLeaveRequest extends ReposAppRequest {
  orgLeave?: {
    state: OrganizationMembershipState;
    memberOfOrgs?;
  };
}

router.use(asyncHandler(async (req: ILocalLeaveRequest, res, next) => {
  const organization = req.organization as Organization;
  req.orgLeave = {
    state: null,
  };
  const username = req.individualContext.getGitHubIdentity().username;
  const memberOfOrgs: IOrganizationMembershipState[] = [];
  const stateResult = await organization.getOperationalMembership(username);
  const state = stateResult ? stateResult.state : null;
  req.orgLeave.state = state;
  if (!req.orgLeave.state) {
    return res.redirect('/');
  } else {
    req.orgLeave.memberOfOrgs = memberOfOrgs;
    req.individualContext.webContext.pushBreadcrumb('Leave');
    return next();
  }
}));

router.get('/', function (req: ILocalLeaveRequest, res) {
  const organization = req.organization;
  req.individualContext.webContext.render({
    view: 'org/leave',
    title: 'Leave ' + organization.name,
    state: {
      organization,
    },
  });
});

router.post('/', asyncHandler(async function (req: ReposAppRequest, res, next) {
  const organization = req.organization;
  const providers = getProviders(req);
  const operations = providers.operations;
  const username = req.individualContext.getGitHubIdentity().username;
  const id = req.individualContext.getGitHubIdentity().id;
  try {
    await organization.removeMember(username, id);
    req.individualContext.webContext.saveUserAlert(`You have been removed from the ${organization.name} and are no longer a member.`, organization.name, UserAlertType.Success);
    return res.redirect(operations.baseUrl || '/');
  } catch (error) {
    return next(wrapError(error, `We received an error code back from GitHub when trying to remove your membership from ${organization.name}.`));
  }
}));

export default router;
