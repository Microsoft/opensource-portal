//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import RouteApproval from './approval';

import { getProviders } from '../../../transitional';
import { wrapError } from '../../../utils';
import { Team } from '../../../business';
import { TeamJoinApprovalEntity } from '../../../entities/teamJoinApproval/teamJoinApproval';
import { Account } from '../../../business';
import { ReposAppRequest, IRequestTeams } from '../../../interfaces';

// Not a great place for these, should move into independent files eventually...

interface IPermissionWorkflowApprovalPackage {
  request: TeamJoinApprovalEntity;
  id: string;
  requestingUser: Account; // ??  string;
}

enum PermissionWorkflowDecision {
  Approve = 'approve',
  Deny = 'deny',
}

export class PermissionWorkflowEngine {
  public team: Team;
  public request: TeamJoinApprovalEntity;
  public user: Account; // string;
  public id: string;
  public typeName: string;

  constructor(team: Team, approvalPackage: IPermissionWorkflowApprovalPackage) {
    this.team = team;
    if (!team) {
      throw new Error('No team instance');
    }
    this.request = approvalPackage.request;
    this.user = approvalPackage.requestingUser;
    this.id = approvalPackage.id;
    this.typeName = 'Team Join';
  }

  getDecisionEmailViewName() {
    return 'membershipApprovals/decision';
  }

  getDecisionEmailSubject(approved, request: TeamJoinApprovalEntity) {
    return approved ? `Welcome to the ${request.teamName} ${request.organizationName} GitHub team` : `Your ${request.teamName} permission request was not approved`;
  }

  getDecisionEmailHeadline(approved/*, request*/) {
    return approved ? 'Welcome' : 'Sorry';
  }

  messageForAction(action: PermissionWorkflowDecision) {
    let message = null;
    if (action === PermissionWorkflowDecision.Deny) {
      message = 'This team join request has not been approved at this time.';
    } else if (action === PermissionWorkflowDecision.Approve) {
      message = 'Permission request approved.';
    }
    return message;
  }

  performApprovalOperation(callback) {
    const team = this.team;
    const username = this.request.thirdPartyUsername;
    return team.addMembership(username).then(ok => {
      return callback();
    }).catch(error => {
      error = wrapError(error, `The GitHub API returned an error trying to add the user ${username} to team ID ${team.id}.`);
      return callback(error);
    });
  }
}

// Find the request and assign the workflow engine

router.use(function (req: ReposAppRequest, res, next) {
  req.individualContext.webContext.pushBreadcrumb('Approvals');
  next();
});

router.get('/', asyncHandler(async (req: IRequestTeams, res, next) => {
  const team = req.team2 as Team;
  const approvals = await team.getApprovals();
  req.individualContext.webContext.render({
    view: 'org/team/approvals',
    title: 'Approvals for ' + team.name,
    state: {
      team: team,
      pendingApprovals: approvals,
      teamUrl: req.teamUrl,
    },
  });
}));


interface IRequestPlusApprovalEngine extends IRequestTeams {
  approvalEngine?: PermissionWorkflowEngine;
}

router.use('/:requestid', asyncHandler(async function (req: IRequestPlusApprovalEngine, res, next) {
  const team = req.team2 as Team;
  const requestid = req.params.requestid;
  const { approvalProvider, operations } = getProviders(req);
  if (!approvalProvider) {
    return next(new Error('No approval provider instance available'));
  }
  try {
    const pendingRequest = await approvalProvider.getApprovalEntity(requestid);
    const requestingUserAccount = await operations.getAccountWithDetailsAndLink(pendingRequest.thirdPartyId);
    const approvalPackage = {
      request: pendingRequest,
      requestingUser: requestingUserAccount,
      id: requestid,
    };
    if (String(approvalPackage.request.teamId) !== String(team.id)) {
      res.status(400);
      return next(new Error('Mismatch on team'));
    }
    const engine = new PermissionWorkflowEngine(team, approvalPackage);
    req.individualContext.webContext.pushBreadcrumb(engine.typeName + ' Request');
    req.approvalEngine = engine;
    return next();
  } catch (error) {
    return next(wrapError(error, 'The pending request you are looking for does not seem to exist.'));
  }
}));

// Pass on to the context-specific routes.
router.use('/:requestid', RouteApproval);

export default router;
