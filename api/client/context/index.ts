//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';

import { Organization } from '../../../business';

import { jsonError } from '../../../middleware';
import getCompanySpecificDeployment from '../../../middleware/companySpecificDeployment';
import {
  ErrorHelper,
  getProviders,
  IProviders,
  ReposAppRequest,
} from '../../../transitional';
import { IndividualContext } from '../../../user';

import RouteApprovals from './approvals';
import RouteIndividualContextualOrganization from './organization';
import RouteOrgs from './orgs';
import RouteRepos from './repos';
import RouteTeams from './teams';

const router = express.Router();

const deployment = getCompanySpecificDeployment();
deployment?.routes?.api?.context?.index &&
  deployment?.routes?.api?.context?.index(router);

router.use('/approvals', RouteApprovals);

router.get('/', (req: ReposAppRequest, res) => {
  const { config } = getProviders(req);
  const { continuousDeployment } = config;
  const activeContext = (req.individualContext ||
    req.apiContext) as IndividualContext;
  const isGitHubAuthenticated = !!activeContext.getSessionBasedGitHubIdentity()
    ?.id;
  const data = {
    corporateIdentity: activeContext.corporateIdentity,
    githubIdentity: activeContext.getGitHubIdentity(),
    isAuthenticated: true,
    isGitHubAuthenticated,
    isLinked: !!activeContext.link,
    build: continuousDeployment,
  };
  return res.json(data);
});

router.get(
  '/accountDetails',
  asyncHandler(async (req: ReposAppRequest, res) => {
    const { operations } = getProviders(req);
    const activeContext = (req.individualContext ||
      req.apiContext) as IndividualContext;
    const gh = activeContext.getGitHubIdentity();
    if (!gh || !gh.id) {
      res.status(400);
      res.end();
    }
    const accountFromId = operations.getAccount(gh.id);
    const accountDetails = await accountFromId.getDetails();
    res.json(accountDetails);
  })
);

router.get('/orgs', RouteOrgs);

router.get('/repos', RouteRepos);

router.get('/teams', RouteTeams);

router.use(
  '/orgs/:orgName',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { orgName } = req.params;
    const { operations } = getProviders(req);
    // const activeContext = (req.individualContext || req.apiContext) as IndividualContext;
    // if (!activeContext.link) {
    //   return next(jsonError('Account is not linked', 400));
    // }
    let organization: Organization = null;
    try {
      organization = operations.getOrganization(orgName);
      // CONSIDER: what if they are not currently a member of the org?
      req.organization = organization;
      return next();
    } catch (noOrgError) {
      if (ErrorHelper.IsNotFound(noOrgError)) {
        res.status(404);
        return res.end();
      }
      return next(jsonError(noOrgError, 500));
    }
  })
);

router.use('/orgs/:orgName', RouteIndividualContextualOrganization);

router.use('*', (req: ReposAppRequest, res, next) => {
  return next(jsonError('Contextual API or route not found', 404));
});

export default router;
