//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import _ from 'lodash';

import { getProviders } from '../../transitional';
import { jsonError } from '../../middleware/jsonError';
import { IndividualContext } from '../../user';
import { Organization } from '../../business/organization';
import { CreateRepository, ICreateRepositoryApiResult, CreateRepositoryEntrypoint } from '../createRepo';
import { Team } from '../../business/team';
import { GitHubTeamRole, ReposAppRequest } from '../../interfaces';

// This file supports the client apps for creating repos.

interface ILocalApiRequest extends ReposAppRequest {
  apiVersion?: string;
  organization?: Organization;
  knownRequesterMailAddress?: string;
}

router.get('/metadata', (req: ILocalApiRequest, res, next) => {
  try {
    const options = {
      projectType: req.query.projectType,
    };
    const organization = req.organization as Organization;
    const metadata = organization.getRepositoryCreateMetadata(options);
    res.json(metadata);
  } catch (error) {
    return next(jsonError(error, 400));
  }
});

router.get('/personalizedTeams', asyncHandler(async (req: ILocalApiRequest, res, next) => {
  try {
    const organization = req.organization as Organization;
    const userAggregateContext = req.apiContext.aggregations;
    const maintainedTeams = new Set<string>();
    const broadTeams = new Set<number>(req.organization.broadAccessTeams);
    const userTeams = userAggregateContext.reduceOrganizationTeams(organization, await userAggregateContext.teams());
    userTeams.maintainer.map(maintainedTeam => maintainedTeams.add(maintainedTeam.id.toString()));
    const combinedTeams = new Map<string, Team>();
    userTeams.maintainer.map(team => combinedTeams.set(team.id.toString(), team));
    userTeams.member.map(team => combinedTeams.set(team.id.toString(), team));
    const personalizedTeams = Array.from(combinedTeams.values()).map(combinedTeam => {
      return {
        broad: broadTeams.has(Number(combinedTeam.id)),
        description: combinedTeam.description,
        id: Number(combinedTeam.id),
        name: combinedTeam.name,
        role: maintainedTeams.has(combinedTeam.id.toString()) ? GitHubTeamRole.Maintainer : GitHubTeamRole.Member,
      }
    });
    return res.json({
      personalizedTeams,
    });
  } catch (error) {
    return next(jsonError(error, 400));
  }
}));

router.get('/teams', asyncHandler(async (req: ILocalApiRequest, res, next) => {
  const providers = getProviders(req);
  const queryCache = providers.queryCache;
  const organization = req.organization as Organization;
  const broadTeams = new Set(organization.broadAccessTeams);
  if (req.query.refresh === undefined && queryCache && queryCache.supportsTeams) {
    // Use the newer method in this case...
    const organizationTeams = await queryCache.organizationTeams(organization.id.toString());
    return res.json({
      teams: organizationTeams.map(qt => {
        const team = qt.team;
        const t = team.toSimpleJsonObject();
        if (broadTeams.has(Number(t.id))) {
          t['broad'] = true;
        }
        return t;
      }),
    });
  }

  // By default, allow a 30-second old list of teams. If the cached
  // view is older, refresh this list in the background for use if
  // they refresh for a better user experience.
  const caching = {
    backgroundRefresh: true,
    maxAgeSeconds: 30,
  };

  // If the user forces a true refresh, force a true walk of all the teams
  // from GitHub. This will be slow the larger the org. Allow a short cache
  // window for the casewhere a  webhook processes the change quickly.
  if (req.query.refresh) {
    caching.backgroundRefresh = false;
    caching.maxAgeSeconds = 10;
  }

  try {
    const teams = await organization.getTeams();
    const simpleTeams = teams.map(team => {
      const t = team.toSimpleJsonObject();
      if (broadTeams.has(t.id)) {
        t['broad'] = true;
      }
      return t;
    });
    res.json({
      teams: simpleTeams,
    });
  } catch (getTeamsError) {
    return next(jsonError(getTeamsError, 400));
  }
}));

router.get('/repo/:repo', asyncHandler(async (req: ILocalApiRequest, res) => {
  const { insights } = getProviders(req);
  const repoName = req.params.repo;
  let error = null;
  try {
    const repo = await req.organization.repository(repoName).getDetails();
    res.json(repo);
  } catch (repoDetailsError) {
    res.status(404).end();
    error = repoDetailsError;
  }
  insights?.trackEvent({
    name: 'ApiClientNewRepoValidateAvailability',
    properties: {
      found: error ? true : false,
      repoName,
      org: req.organization.name,
    },
  });
}));

export async function discoverUserIdentities(req: ReposAppRequest, res, next) {
  const apiContext = req.apiContext as IndividualContext;
  const providers = getProviders(req);
  const mailAddressProvider = providers.mailAddressProvider;
  // Try and also learn if we know their e-mail address to send the new repo mail to
  const upn = apiContext.corporateIdentity.username;
  if (apiContext.link && apiContext.link.corporateMailAddress) {
    req['knownRequesterMailAddress'] = apiContext.link.corporateMailAddress;
    return next();
  }
  try {
    const mailAddress = await mailAddressProvider.getAddressFromUpn(upn);
    if (mailAddress) {
      req['knownRequesterMailAddress'] = mailAddress;
    }
  } catch (ignoredError) { /* ignored */ }
  return next();
}

router.post('/repo/:repo', asyncHandler(discoverUserIdentities), asyncHandler(createRepositoryFromClient));

export async function createRepositoryFromClient(req: ILocalApiRequest, res, next) {
  const providers = getProviders(req);
  const { insights, diagnosticsDrop, customizedNewRepositoryLogic, graphProvider } = providers;
  const individualContext = req.individualContext || req.apiContext;
  const config = getProviders(req).config;;
  const organization = (req.organization || (req as any).aeOrganization) as Organization;
  const existingRepoId = req.body.existingrepoid;
  const correlationId = req.correlationId;
  const debugValues = req.body.debugValues || {};
  const corporateId = individualContext.corporateIdentity.id;
  const customContext = customizedNewRepositoryLogic?.createContext(req);
  const additionalTelemetryProperties = customizedNewRepositoryLogic?.getAdditionalTelemetryProperties(customContext) || {};
  insights.trackEvent({
    name: 'CreateRepositoryFromClientStart',
    properties: Object.assign({
      debugValues,
      correlationId,
      body: JSON.stringify(req.body),
      query: JSON.stringify(req.query),
      corporateId,
    }, additionalTelemetryProperties),
  });
  if (diagnosticsDrop) {
    diagnosticsDrop.setObject(`newrepo.${organization.name}.${correlationId}`, Object.assign({
      debugValues,
      correlationId,
      body: req.body,
      query: req.query,
      corporateId,
    }, additionalTelemetryProperties));
  }
  if (organization.createRepositoriesOnGitHub && !(existingRepoId && organization.isNewRepositoryLockdownSystemEnabled())) {
    return next(jsonError(`The GitHub organization ${organization.name} is configured as "createRepositoriesOnGitHub": repos should be created on GitHub.com directly and not through this wizard.`, 400));
  }
  const body = req.body;
  if (!body) {
    return next(jsonError('No body', 400));
  }
  try {
    await customizedNewRepositoryLogic?.validateRequest(customContext, req);
  } catch (validationError) {
    return next(jsonError(validationError, 400));
  }
  req.apiVersion = (req.query['api-version'] || req.headers['api-version'] || '2017-07-27') as string;
  if (req.apiContext && req.apiContext.getGitHubIdentity()) {
    body['ms.onBehalfOf'] = req.apiContext.getGitHubIdentity().username;
  }
  // these fields do not need translation: name, description, private
  let isApproved = customizedNewRepositoryLogic ? customizedNewRepositoryLogic.skipApproval(customContext, body) : false;
  const approvalTypesToIds = config.github.approvalTypes.fields.approvalTypesToIds;
  if (!isApproved) {
    if (approvalTypesToIds[body.approvalType]) {
      body.approvalType = approvalTypesToIds[body.approvalType];
    } else {
      let valid = false;
      Object.getOwnPropertyNames(approvalTypesToIds).forEach(key => {
        if (approvalTypesToIds[key] === body.approvalType) {
          valid = true;
        }
      })
      if (!valid) {
        return next(jsonError('The approval type is not supported or approved at this time', 400));
      }
    }
  }
  // Property supporting private repos from the client
  if (body.visibility === 'private') {
    body.private = true;
    delete body.visibility;
  } else if (body.visibility === 'internal') {
    // visibility: internal is legitimate for GHEC or GHAE users
    body.private = true;
  }
  translateValue(body, 'approvalType', 'ms.approval');
  translateValue(body, 'approvalUrl', 'ms.approval-url');
  translateValue(body, 'justification', 'ms.justification');
  translateValue(body, 'legalEntity', 'ms.entity');
  translateValue(body, 'projectType', 'ms.project-type');
  // Team permissions
  let sufficientTeamsOk = false;
  if (customizedNewRepositoryLogic?.sufficientTeamsConfigured) {
    sufficientTeamsOk = customizedNewRepositoryLogic.sufficientTeamsConfigured(customContext, body);
  }
  if (!sufficientTeamsOk) {
    if (!body.selectedAdminTeams || !body.selectedAdminTeams.length) {
      return next(jsonError('No administration team(s) provided in the request', 400));
    }
  }
  translateTeams(body);
  try {
    // Initial repo contents and license
    const templates = _.keyBy(organization.getRepositoryCreateMetadata().templates, 'id');
    const template = templates[body.template];
    // if (!template) {
      // return next(jsonError('There was a configuration problem, the template metadata was not available for this request', 400));
    // }
    translateValue(body, 'template', 'ms.template');
    body['ms.license'] = template && (template.spdx || template.name); // Today this is the "template name" or SPDX if available
    translateValue(body, 'gitIgnoreTemplate', 'gitignore_template');
  } catch (templateError) {
    insights?.trackException({ exception: templateError });
    console.warn(`Optional template error: ${templateError}`);
  }
  if (!body['ms.notify']) {
    try {
      if (graphProvider && individualContext?.corporateIdentity?.id) {
        const info = await graphProvider.getUserById(individualContext.corporateIdentity.id);
        if (info?.mail) {
          body['ms.notify'] = info.mail;
        }
      }
    } catch (graphError) {
      insights?.trackException({ exception: graphError });
      console.warn(`Retrieving user mail address error: ${graphError}`);
    }
  }
  if (!body['ms.notify']) {
    body['ms.notify'] = individualContext?.link?.corporateMailAddress || req.knownRequesterMailAddress || config.brand.operationsMail || config.brand.supportMail;
  }
  // these fields are currently ignored: orgName
  delete body.orgName;
  customizedNewRepositoryLogic?.stripRequestBody(customContext, body);
  insights.trackEvent({
    name: 'ApiClientNewOrgRepoStart',
    properties: {
      body: JSON.stringify(req.body),
    },
  });
  let success: ICreateRepositoryApiResult = null;
  try {
    const newRepositoryParameters = customizedNewRepositoryLogic?.additionalCreateRepositoryParameters ?
      Object.assign(body, customizedNewRepositoryLogic.additionalCreateRepositoryParameters(customContext)) : body;
    success = await CreateRepository(req, organization, customizedNewRepositoryLogic, customContext, newRepositoryParameters, CreateRepositoryEntrypoint.Client, individualContext);
  } catch (createRepositoryError) {
    insights.trackEvent({
      name: 'ApiClientNewOrgRepoError',
      properties: {
        error: createRepositoryError.message,
        encoded: JSON.stringify(createRepositoryError),
      },
    });
    if (!createRepositoryError.json) {
      createRepositoryError = jsonError(createRepositoryError, 400);
    }
    return next(createRepositoryError);
  }
  let message = success.github ? `Your new repo, ${success.github.name}, has been created:` : 'Your repo request has been submitted.';
  if (existingRepoId && success.github) {
    message = `Your repository ${success.github.name} is classified and the repo is now ready, unlocked, with your selected team permissions assigned.`;
  }
  const output = {
    ...success,
    title: existingRepoId ? 'Repository unlocked' : 'Repository created',
    message,
    url: null,
    messages: null,
  };
  if (success.github) {
    output.url = success.github.html_url;
  }
  output.messages = output['tasks'];
  delete output['tasks'];
  insights.trackEvent({
    name: 'ApiClientNewOrgRepoSuccessful',
    properties: {
      body: JSON.stringify(body),
      success: JSON.stringify(output),
    },
  });
  return res.json(output);
}

function translateTeams(body) {
  let admin = body.selectedAdminTeams;
  let write = body.selectedWriteTeams;
  let read = body.selectedReadTeams;
  let maintain = body.selectedMaintainTeams;

  // Remove teams with higher privileges already
  _.pullAll(maintain, admin);
  _.pullAll(maintain, write);
  _.pullAll(maintain, read);
  _.pullAll(write, admin);
  _.pullAll(read, admin);
  _.pullAll(read, write);

  body['ms.teams'] = {
    maintain,
    admin,
    push: write,
    pull: read,
  };

  delete body.selectedMaintainTeams;
  delete body.selectedAdminTeams;
  delete body.selectedWriteTeams;
  delete body.selectedReadTeams;
}

function translateValue(object, fromKey, toKey) {
  if (object[fromKey]) {
    object[toKey] = object[fromKey];
  }
  if (object[fromKey] !== undefined) {
    delete object[fromKey];
  }
}

export default router;
