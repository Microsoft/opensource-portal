//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { getProviders } from '../../../../transitional';
import { Team } from '../../../../business';
import { PermissionWorkflowEngine } from '../approvals';
import RenderHtmlMail from '../../../../lib/emailRender';
import { IndividualContext } from '../../../../user';
import { ReposAppRequest, UserAlertType, IProviders } from '../../../../interfaces';

interface ILocalRequest extends ReposAppRequest {
  team2?: any;
  approvalEngine?: PermissionWorkflowEngine;
  teamUrl?: any;
}

router.get('/', function (req: ILocalRequest, res) {
  const approvalRequest = req.approvalEngine.request;
  const team2 = req.team2 as Team;
  // Ignoring any errors for now.
  if (approvalRequest.created) {
    // legacy compat, can remove with a view update
    approvalRequest['requestedTime'] = approvalRequest.created;
  }
  req.individualContext.webContext.render({
    view: 'org/team/approveStatus',
    title: 'Request Status',
    state: {
      metadata: approvalRequest,
      requestingUser: req.approvalEngine.user,
      team: team2,
      teamUrl: req.teamUrl,
    },
  });
});

router.get('/setNote/:action', function (req: ILocalRequest, res) {
  var engine = req.approvalEngine;
  var action = req.params.action;
  if (action == 'approveWithComment') {
    action = 'approve';
  }
  const team2 = req.team2;
  req.individualContext.webContext.render({
    view: 'org/team/approveStatusWithNote',
    title: 'Record your comment for request ' + engine.id + ' (' + action + ')',
    state: {
      metadata: engine.request,
      action: action,
      requestingUser: engine.user,
      team: team2,
      teamUrl: req.teamUrl,
    },
  });
});

router.post('/', asyncHandler(async (req: ILocalRequest, res, next) => {
  const providers = getProviders(req);
  const { individualContext } = req;
  const engine = req.approvalEngine as PermissionWorkflowEngine;
  let message = req.body.text as string;
  const teamBaseUrl = req.teamUrl as string;
  let decision: TeamApprovalDecision = null;
  if (req.body.reopen) {
    decision = TeamApprovalDecision.Reopen;
  } else if (req.body.approve || req.body.approveWithComment) {
    decision = TeamApprovalDecision.Approve;
  } else if (req.body.deny) {
    decision = TeamApprovalDecision.Deny;
  }
  if (!decision) {
    throw new Error('No valid decision');
  }
  const outcome = await postActionDecision(providers, individualContext, engine, teamBaseUrl, decision, message);
  if (outcome.message) {
    req.individualContext.webContext.saveUserAlert(outcome.message, engine.typeName, UserAlertType.Success);
  }
  if (outcome.error) {
    req.insights.trackException({
      exception: outcome.error,
    });
    return next(outcome.error);
  }
  return res.redirect(outcome.redirect || teamBaseUrl);
}));

export enum TeamApprovalDecision {
  Approve = 'Approved',
  Deny = 'Denied',
  Reopen = 'Reopened',
}

export interface IPostActionDecisionOutcome {
  redirect?: string;
  message?: string;
  error?: any;
}

function performApprovalWithEngine(engine: PermissionWorkflowEngine): Promise<void> {
  return new Promise((resolve, reject) => {
    engine.performApprovalOperation((error: Error) => {
      return error ? reject(error) : resolve();
    });
  });
}

export async function postActionDecision(providers: IProviders, individualContext: IndividualContext, engine: PermissionWorkflowEngine, teamBaseUrl: string, decision: TeamApprovalDecision, messageToRequestor: string): Promise<IPostActionDecisionOutcome> {
  if (!individualContext || !individualContext.getGitHubIdentity().username) {
    return { error: new Error('No individual context') };
  }
  const approvalRequest = engine.request;
  const requestid = engine.id;
  const { approvalProvider: teamJoinApprovalProvider, config, mailAddressProvider, mailProvider, insights } = providers;
  if (decision === TeamApprovalDecision.Reopen) {
    approvalRequest.active = true;
    try {
      await teamJoinApprovalProvider.updateTeamApprovalEntity(approvalRequest);
      return {
        message: 'Request reopened',
        redirect: `${teamBaseUrl}approvals/${requestid}`,
      };
    } catch (error) {
      return { error };
    }
  }
  const action = decision === TeamApprovalDecision.Approve ? 'approve' : 'deny';
  const username = individualContext.getGitHubIdentity().username;
  let userMailAddress: string = null;
  const decisionMessage = messageToRequestor || decision;
  const pendingRequest = engine.request;
  try {
    const upn = pendingRequest.corporateUsername;
    userMailAddress = await mailAddressProvider.getAddressFromUpn(upn);
    pendingRequest.decision = action;
    pendingRequest.active = false;
    pendingRequest.decisionTime = new Date();
    pendingRequest.decisionThirdPartyUsername = username;
    pendingRequest.decisionThirdPartyId = individualContext.getGitHubIdentity().id;
    pendingRequest.decisionMessage = decisionMessage;
    pendingRequest.decisionCorporateUsername = individualContext.corporateIdentity.username;
    pendingRequest.decisionCorporateId = individualContext.corporateIdentity.id;
    await teamJoinApprovalProvider.updateTeamApprovalEntity(pendingRequest);
    if (decision == TeamApprovalDecision.Approve) {
      await performApprovalWithEngine(engine);
    }
  } catch (error) {
    return { error };
  }
  const message = `Thanks for your ${action.toUpperCase()} decision`;
  const approvalMail = individualContext.link.corporateMailAddress || individualContext.link.corporateUsername;
  if (mailProvider) {
    const wasApproved = decision === TeamApprovalDecision.Approve;
    const contentOptions = {
      correlationId: individualContext.webContext?.correlationId,
      pendingRequest,
      version: config.logging.version,
      results: [], // no longer a used field, used to be called 'output'
      wasApproved,
      decisionBy: username,
      decisionNote: decisionMessage,
      decisionEmail: approvalMail,
      reason: (`You are receiving this e-mail because of a request that you created, and a decision has been made.
                This mail was sent to: ${userMailAddress}`),
      headline: engine.getDecisionEmailHeadline(wasApproved),
      notification: wasApproved ? 'information' : 'warning',
      service: (config.brand?.companyName || 'Corporate') + ' GitHub',
      companyName: config.brand.companyName,
    };
    if (!engine.getDecisionEmailViewName || !engine.getDecisionEmailSubject) {
      return { message, redirect: teamBaseUrl };
    }
    // req.individualContext.webContext.saveUserAlert('Thanks for your ' + action.toUpperCase() + ' decision.', engine.typeName, 'success');
    const getDecisionEmailViewName = engine.getDecisionEmailViewName();
    let content = null;
    try {
      content = await RenderHtmlMail(config.typescript.appDirectory, getDecisionEmailViewName, contentOptions);
    } catch (renderError) {
    }
    if (content) {
      const mail = {
        to: [userMailAddress],
        subject: engine.getDecisionEmailSubject(wasApproved, pendingRequest),
        content,
        correlationId: individualContext.webContext?.correlationId,
      };
      try {
        const mailResult = await mailProvider.sendMail(mail);
        insights?.trackEvent({
          name: 'ReposRequestDecisionMailSuccess', properties: Object.assign({
            receipt: mailResult,
          }, contentOptions)
        });
      } catch (mailError) {
        insights?.trackException({
          exception: mailError, properties: Object.assign({
            eventName: 'ReposRequestDecisionMailFailure',
          }, contentOptions)
        });
      }
    }
  }
  return { message, redirect: teamBaseUrl };
}

export default router;
