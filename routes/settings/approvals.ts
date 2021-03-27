//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import {
  IReposError,
  ReposAppRequest,
  UserAlertType,
  getProviders,
} from '../../transitional';
import { IApprovalProvider } from '../../entities/teamJoinApproval/approvalProvider';
import { TeamJoinApprovalEntity } from '../../entities/teamJoinApproval/teamJoinApproval';
import { safeLocalRedirectUrl } from '../../utils';
import { Operations } from '../../business';
import { Team } from '../../business';
import { Organization } from '../../business';
import { IAggregateUserTeams } from '../../user/aggregate';

export interface ApprovalPair {
  team: Team;
  request: TeamJoinApprovalEntity;
}

export async function Approvals_getTeamMaintainerApprovals(
  operations: Operations,
  aggregateTeams: IAggregateUserTeams,
  approvalProvider: IApprovalProvider
): Promise<ApprovalPair[]> {
  // TODO: move to team object?
  const ownedTeamIdsAsStrings = aggregateTeams.maintainer.map((team) =>
    team.id.toString()
  );
  if (ownedTeamIdsAsStrings.length === 0) {
    return [];
  }
  const pendingApprovals = await approvalProvider.queryPendingApprovalsForTeams(
    ownedTeamIdsAsStrings
  );
  const pairs: ApprovalPair[] = [];
  for (const request of pendingApprovals) {
    try {
      const pair = await hydrateRequest(operations, request);
      if (pair) {
        pairs.push(pair);
      }
    } catch (ignored) {
      /* ignored */
    }
  }
  return pairs;
}

async function hydrateRequest(
  operations: Operations,
  request: TeamJoinApprovalEntity
): Promise<ApprovalPair> {
  const teamIdAsNumber = Number(request.teamId);
  const organizationName = request.organizationName;
  const team = operations.getTeamByIdWithOrganization(
    teamIdAsNumber,
    organizationName
  );
  await team.getDetails();
  if (team && team.name) {
    return { team, request };
  }
}

export async function Approvals_getUserRequests(
  operations: Operations,
  thirdPartyIdAsString: string,
  approvalProvider: IApprovalProvider
): Promise<ApprovalPair[]> {
  const pendingApprovals = await approvalProvider.queryPendingApprovalsForThirdPartyId(
    thirdPartyIdAsString
  );
  const pairs: ApprovalPair[] = [];
  for (const request of pendingApprovals) {
    try {
      const pair = await hydrateRequest(operations, request);
      if (pair) {
        pairs.push(pair);
      }
    } catch (ignored) {
      /* ignored */
    }
  }
  return pairs;
}

router.get(
  '/',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const { approvalProvider, operations } = getProviders(req);
    if (!approvalProvider) {
      return next(new Error('No approval provider instance available'));
    }
    req.individualContext.webContext.pushBreadcrumb('Requests');
    // CONSIDER: Requests on GitHub.com should be shown, too, now that that's integrated in many cases
    const id = req.individualContext.getGitHubIdentity().id;
    const aggregateTeams = await req.individualContext.aggregations.teams();
    const state = {
      teamResponsibilities: await Approvals_getTeamMaintainerApprovals(
        operations,
        aggregateTeams,
        approvalProvider
      ),
      usersRequests: await Approvals_getUserRequests(
        operations,
        id.toString(),
        approvalProvider
      ),
    };
    req.individualContext.webContext.render({
      view: 'settings/approvals',
      title: 'Review My Approvals',
      state,
    });
  })
);

router.post('/:requestid/cancel', function (req: ReposAppRequest, res, next) {
  const { approvalProvider } = getProviders(req);
  if (!approvalProvider) {
    return next(new Error('No approval provider instance available'));
  }
  const safeReturnUrl = safeLocalRedirectUrl(
    req.body.returnUrl || req.params.returnUrl
  );
  const requestid = req.params.requestid;
  const id = req.individualContext.getGitHubIdentity().id;
  approvalProvider
    .getApprovalEntity(requestid)
    .catch((error) => {
      return next(
        new Error(
          'The pending request you are looking for does not seem to exist.'
        )
      );
    })
    .then((pendingRequest: TeamJoinApprovalEntity) => {
      if (pendingRequest.thirdPartyId == id) {
        pendingRequest.active = false;
        pendingRequest.decisionMessage = 'canceled-by-user';
        pendingRequest.decisionTime = new Date(); // (new Date().getTime()).toString();
        approvalProvider
          .updateTeamApprovalEntity(pendingRequest)
          .catch((error) => {
            if (error) {
              return next(error);
            }
          })
          .then((unused) => {
            return res.redirect(safeReturnUrl || '/settings/approvals/');
          });
      } else {
        return next(
          new Error('You are not authorized to cancel this request.')
        );
      }
    });
});

router.get(
  '/:requestid',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const requestid = req.params.requestid;
    const { approvalProvider, operations } = getProviders(req);
    req.individualContext.webContext.pushBreadcrumb('Your Request');
    let isMaintainer = false;
    let pendingRequest: TeamJoinApprovalEntity = null;
    let team2: Team = null;
    let maintainers = null;
    const username = req.individualContext.getGitHubIdentity().username;
    const id = req.individualContext.getGitHubIdentity().id;
    let organization: Organization = null;
    try {
      pendingRequest = await approvalProvider.getApprovalEntity(requestid);
      organization = operations.getOrganization(
        pendingRequest.organizationName
      );
      team2 = organization.team(Number(pendingRequest.teamId));
      await team2.getDetails();
      const isOrgSudoer = await organization.isSudoer(
        username,
        req.individualContext?.link
      );
      isMaintainer = isOrgSudoer;
      const maintainers = await team2.getOfficialMaintainers();
      if (!isMaintainer) {
        for (let i = 0; i < maintainers.length; i++) {
          if (String(maintainers[i].id) == String(id)) {
            isMaintainer = true;
          }
        }
      }
      if (isMaintainer) {
        let err: IReposError = new Error(
          'Redirecting to the admin experience to approve'
        );
        let slugPreferred = team2.slug || team2.name;
        err.redirect =
          '/' +
          organization.name +
          '/teams/' +
          slugPreferred +
          '/approvals/' +
          requestid;
        throw err;
      }
      if (pendingRequest.thirdPartyId != /* loose */ id) {
        let msg: IReposError = new Error(
          'This request does not exist or was created by another user.'
        );
        msg.skipLog = true;
        throw msg;
      }
    } catch (error) {
      if (error.redirect) {
        return res.redirect(error.redirect);
      }
      // Edge case: the team no longer exists.
      if (
        error.innerError &&
        error.innerError.innerError &&
        error.innerError.innerError.statusCode == 404
      ) {
        return closeOldRequest(
          false /* not a JSON client app */,
          pendingRequest,
          req,
          res,
          next
        );
      }
      return next(error);
    }
    req.individualContext.webContext.render({
      view: 'org/userApprovalStatus',
      title: 'Review your request',
      state: {
        entry: pendingRequest,
        team: team2,
      },
    });
  })
);

export function closeOldRequest(
  isJsonClient: boolean,
  pendingRequest: TeamJoinApprovalEntity,
  req: ReposAppRequest,
  res,
  next
) {
  const { approvalProvider } = getProviders(req);
  const config = getProviders(req).config;
  const repoApprovalTypesValues = config.github.approvalTypes.repo;
  if (repoApprovalTypesValues.length === 0) {
    return next(new Error('No repo approval providers configured.'));
  }
  const repoApprovalTypes = new Set(repoApprovalTypesValues);
  const mailProviderInUse = repoApprovalTypes.has('mail');
  if (!mailProviderInUse) {
    return next(new Error('No configured approval providers configured.'));
  }
  if (!isJsonClient) {
    req.individualContext.webContext.saveUserAlert(
      'The team this request was for no longer exists. The request has been canceled.',
      'Team gone!',
      UserAlertType.Success
    );
  }
  if (pendingRequest.active === false) {
    return isJsonClient ? res.json({}) : res.redirect('/');
  }
  closeRequest(
    approvalProvider,
    pendingRequest.approvalId,
    'Team no longer exists.',
    (closeError: Error) => {
      if (closeError) {
        return next(closeError);
      }
      return isJsonClient ? res.json({}) : res.redirect('/');
    }
  );
}

function closeRequest(
  approvalProvider: IApprovalProvider,
  requestid: string,
  note: string,
  callback
) {
  approvalProvider
    .getApprovalEntity(requestid)
    .then((pendingRequest) => {
      pendingRequest.active = false;
      pendingRequest.decisionMessage = note;
      return approvalProvider
        .updateTeamApprovalEntity(pendingRequest)
        .then((ok) => {
          return callback(null, ok);
        });
    })
    .catch((error) => {
      return callback(error);
    });
}

export default router;
