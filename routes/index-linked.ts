//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { CreateError, hasStaticReactClientApp, getProviders } from '../transitional';
import { IndividualContext } from '../user';
import { storeOriginalUrlAsVariable } from '../utils';
import { AuthorizeOnlyCorporateAdministrators } from '../middleware/business/corporateAdministrators';

import RouteAdministration from './administration';
import RouteUndo from './undo';
import RouteTeams from './teams';
import RoutePeople from './people';
import RouteRepos from './repos';
import RouteLegacyOrganizationAdministration from './orgAdmin';

import unlinkRoute from './unlink';
import { Organization, Repository } from '../business';

import orgsRoute from './orgs';
import { injectReactClient } from '../middleware';
import { ReposAppRequest } from '../interfaces';

const hasReactApp = hasStaticReactClientApp();
const reactRoute = hasReactApp ? injectReactClient() : undefined;

//-----------------------------------------------------------------------------
//-----------------------------------------------------------------------------
// SECURITY ROUTE MARKER:
// Below this next call, all routes will require an active link to exist for
// the authenticated GitHub user.
//-----------------------------------------------------------------------------
// * only for the traditional app. The React app does not require a link to browse orgs.
//-----------------------------------------------------------------------------
if (!hasReactApp) {
  router.use(function (req: ReposAppRequest, res, next) {
    const individualContext = req.individualContext as IndividualContext;
    const link = individualContext.link;
    if (link && link.thirdPartyId) {
      return next();
    }
    storeOriginalUrlAsVariable(req, res, 'beforeLinkReferrer', '/', 'no linked github username');
  });
}
// end security route
//-----------------------------------------------------------------------------
//-----------------------------------------------------------------------------

router.use('/unlink', unlinkRoute);

router.use('/organization', RouteLegacyOrganizationAdministration); // admin UI, not in React

router.use('/teams', reactRoute || RouteTeams);
router.use('/people', reactRoute || RoutePeople);
router.use('/repos', reactRoute || RouteRepos);

// Routes not yet available in the client
router.use('/undo', RouteUndo);
router.use('/administration', AuthorizeOnlyCorporateAdministrators, RouteAdministration);

router.use('/https?*github.com/:org/:repo', asyncHandler(async (req: ReposAppRequest, res, next) => {
  // Helper method to allow pasting a GitHub URL into the app to go to a repo
  const { org, repo } = req.params;
  const { operations } = getProviders(req);
  if (org && repo) {
    let organization: Organization = null;
    try {
      organization = operations.getOrganization(org);
    } catch (error) {
      return next(CreateError.InvalidParameters(`Organization ${org} not managed by this system`));
    }
    let repository: Repository = null;
    try {
      repository = organization.repository(repo);
      await repository.getDetails();
    }
    catch (error) {
      return next(CreateError.NotFound(`The repository ${org}/${repo} no longer exists.`));
    }
    if (hasReactApp) {
      return res.redirect(`/orgs/${repository.organization.name}/repos/${repository.name}`);
    }
    return res.redirect(repository.baseUrl);
  }
  return next();
}));

router.use('/', orgsRoute);

export default router;
