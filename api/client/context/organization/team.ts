//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import { TeamJoinApprovalEntity } from '../../../../entities/teamJoinApproval/teamJoinApproval';
import { ReposAppRequest, OrganizationMembershipState, ITeamMembershipRoleState } from '../../../../interfaces';
import { jsonError } from '../../../../middleware';
import { AddTeamMembershipToRequest, AddTeamPermissionsToRequest, getContextualTeam, getTeamMembershipFromRequest, getTeamPermissionsFromRequest } from '../../../../middleware/github/teamPermissions';
import { submitTeamJoinRequest } from '../../../../routes/org/team';
import { postActionDecision, TeamApprovalDecision } from '../../../../routes/org/team/approval';
import { PermissionWorkflowEngine } from '../../../../routes/org/team/approvals';
import { getProviders } from '../../../../transitional';
import { IndividualContext } from '../../../../user';

const router: Router = Router();

interface ITeamRequestJsonResponse {
  request?: TeamJoinApprovalEntity;
}

interface ITeamApprovalsJsonResponse {
  allowAdministration: boolean;
  approvals?: TeamJoinApprovalEntity[];
}

router.get('/permissions',
  asyncHandler(AddTeamPermissionsToRequest),
  asyncHandler(AddTeamMembershipToRequest),
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const membership = getTeamMembershipFromRequest(req);
    const permissions = getTeamPermissionsFromRequest(req);
    return res.json({ permissions, membership });
  })
);

router.get('/join/request', asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { approvalProvider } = getProviders(req);
  const team = getContextualTeam(req);
  const activeContext = (req.individualContext || req.apiContext) as IndividualContext;
  let request: TeamJoinApprovalEntity = null;
  if (activeContext.link) {
    // no point query currently implemented
    let approvals = await approvalProvider.queryPendingApprovalsForTeam(String(team.id));
    approvals = approvals.filter(approval => approval.corporateId === activeContext.corporateIdentity.id);
    request = approvals.length > 0 ? approvals[0] : null;
  }
  const response: ITeamRequestJsonResponse = { request };
  return res.json(response);
}));

router.post('/join',
  asyncHandler(AddTeamMembershipToRequest),
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    try {
      const providers = getProviders(req);
      const { approvalProvider } = providers;
      const membership = getTeamMembershipFromRequest(req);
      if (!membership.isLinked) {
        return res.json({ error: 'You have not linked your GitHub account to your corporate identity yet' });
      }
      if (membership.membershipState === OrganizationMembershipState.Active) {
        return res.json({ error: 'You already have an active team membership' });
      }
      const team = getContextualTeam(req);
      const activeContext = (req.individualContext || req.apiContext) as IndividualContext;
      // no point query currently implemented
      let approvals = await approvalProvider.queryPendingApprovalsForTeam(String(team.id));
      approvals = approvals.filter(approval => approval.corporateId === activeContext.corporateIdentity.id);
      const request = approvals.length > 0 ? approvals[0] : null;
      if (request) {
        return res.json({ error: 'You already have a pending team join request' });
      }
      //
      const justification = (req.body.justification || '') as string;
      const hostname = req.hostname;
      const correlationId = req.correlationId;
      const outcome = await submitTeamJoinRequest(providers, activeContext, team, justification, correlationId, hostname);
      return res.json(outcome);
    } catch (error) {
      return next(jsonError(error));
    }
  }));

router.post('/join/approvals/:approvalId',
  asyncHandler(AddTeamPermissionsToRequest),
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { approvalId: id } = req.params;
    if (!id) {
      return next(jsonError('invalid approval', 400));
    }
    const permissions = getTeamPermissionsFromRequest(req);
    if (!permissions.allowAdministration) {
      return next(jsonError('you do not have permission to administer this team', 401));
    }
    const providers = getProviders(req);
    const { approvalProvider, operations } = providers;
    const team = getContextualTeam(req);
    const request = await approvalProvider.getApprovalEntity(id);
    if (String(request.teamId) !== String(team.id)) {
      return next(jsonError('mismatch on team', 400));
    }
    const requestingUser = await operations.getAccountWithDetailsAndLink(request.thirdPartyId);
    const approvalPackage = { request, requestingUser, id };
    const engine = new PermissionWorkflowEngine(team, approvalPackage);
    const activeContext = (req.individualContext || req.apiContext) as IndividualContext;
    const text = req.body.text as string;
    const dv = req.body.decision as string;
    let decision: TeamApprovalDecision = null;
    switch (dv) {
      case 'approve':
        decision = TeamApprovalDecision.Approve;
        break;
      case 'deny':
        decision = TeamApprovalDecision.Deny;
        break;
      case 'reopen':
        decision = TeamApprovalDecision.Reopen;
        break;
      default:
        return next(jsonError('invalid or no decision type', 400));
    }
    const teamBaseUrl = `/orgs/${team.organization.name}/teams/${team.slug}/`; // trailing?
    try {
      const outcome = await postActionDecision(providers, activeContext, engine, teamBaseUrl, decision, text);
      if (outcome.error) {
        throw outcome.error;
      }
      return res.json(outcome);
    } catch (outcomeError) {
      return next(jsonError(outcomeError, 500));
    }
  }));

router.get('/join/approvals/:approvalId',
  asyncHandler(AddTeamPermissionsToRequest),
  asyncHandler(async (req: ReposAppRequest, res, next) => {
  const { approvalId: id } = req.params;
  if (!id) {
    return next(jsonError('invalid approval', 400));
  }
  const permissions = getTeamPermissionsFromRequest(req);
  if (!permissions.allowAdministration) {
    return next(jsonError('you do not have permission to administer this team', 401));
  }
  const providers = getProviders(req);
  const { approvalProvider, operations } = providers;
  const team = getContextualTeam(req);
  const request = await approvalProvider.getApprovalEntity(id);
  if (String(request.teamId) !== String(team.id)) {
    return next(jsonError('mismatch on team', 400));
  }
  return res.json({ approval: request });
}));

router.get('/join/approvals',
  asyncHandler(AddTeamPermissionsToRequest),
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { approvalProvider } = getProviders(req);
    const team = getContextualTeam(req);
    const permissions = getTeamPermissionsFromRequest(req);
    let response: ITeamApprovalsJsonResponse = {
      allowAdministration: false,
    };
    if (permissions.allowAdministration) {
      response.allowAdministration = permissions.allowAdministration;
      response.approvals = await approvalProvider.queryPendingApprovalsForTeam(String(team.id));
    }
    return res.json(response);
}));

router.post('/role/:login',
  asyncHandler(AddTeamPermissionsToRequest),
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { role } = req.body;
    const { login } = req.params;
    if (!login) {
      return next(jsonError('invalid login', 400));
    }
    const permissions = getTeamPermissionsFromRequest(req);
    if (!permissions.allowAdministration) {
      return next(jsonError('you do not have permission to administer this team', 401));
    }
    const team = getContextualTeam(req);
    try {
      const currentRole = await team.getMembership(login, { backgroundRefresh : false, maxAgeSeconds: -1 });
      if (!currentRole || (currentRole as ITeamMembershipRoleState).state !== OrganizationMembershipState.Active) {
        return next(jsonError(`${login} is not currently a member of the team`, 400));
      }
      const response = await team.addMembership(login, { role });
      return res.json(response);
    } catch (outcomeError) {
      return next(jsonError(outcomeError, 500));
    }
  }));

router.use('*', (req, res, next) => {
  return next(jsonError('no API or function available for contextual team', 404));
});

export default router;
