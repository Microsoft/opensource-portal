//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import cors from 'cors';

import { getProviders } from '../transitional';

import { jsonError } from '../middleware';
import { IApiRequest } from '../middleware/apiReposAuth';

import apiExtension from './extension';
import apiClient from './client/';
import apiWebhook from './webhook';

import apiPeople from './people';

import AadApiAuthentication from '../middleware/apiAad';
import AzureDevOpsAuthenticationMiddleware from '../middleware/apiVstsAuth';
import ReposApiAuthentication from '../middleware/apiReposAuth';
import { CreateRepository, CreateRepositoryEntrypoint } from './createRepo';
import supportMultipleAuthProviders from '../middleware/supportMultipleAuthProviders';
import JsonErrorHandler from './jsonErrorHandler';
import getCompanySpecificDeployment from '../middleware/companySpecificDeployment';
import { ReposAppRequest } from '../interfaces';

const hardcodedApiVersions = [
  '2019-10-01',
  '2019-02-01',
  '2017-09-01',
  '2017-03-08',
  '2016-12-01',
];

router.use('/client', apiClient);
router.use('/webhook', apiWebhook);

router.use((req: IApiRequest, res, next) => {
  const apiVersion = (req.query['api-version'] || req.headers['api-version']) as string;
  if (!apiVersion) {
    return next(jsonError('This endpoint requires that an API Version be provided.', 422));
  }
  if (apiVersion.toLowerCase() === '2016-09-22_Preview'.toLowerCase()) {
    return next(jsonError('This endpoint no longer supports the original preview version. Please update your client to use a newer version such as ' + hardcodedApiVersions[0], 422));
  }
  if (hardcodedApiVersions.indexOf(apiVersion.toLowerCase()) < 0) {
    return next(jsonError('This endpoint does not support the API version you provided at this time.', 422));
  }
  req.apiVersion = apiVersion;
  return next();
});

//-----------------------------------------------------------------------------
// AUTHENTICATION: VSTS or repos
//-----------------------------------------------------------------------------
const multipleProviders = supportMultipleAuthProviders([
  AadApiAuthentication,
  ReposApiAuthentication,
  AzureDevOpsAuthenticationMiddleware,
]);

const aadAndCustomProviders = supportMultipleAuthProviders([
  AadApiAuthentication,
  ReposApiAuthentication,
]);

router.use('/people', cors(), multipleProviders, apiPeople);
router.use('/extension', cors(), multipleProviders, apiExtension);

//-----------------------------------------------------------------------------
// AUTHENTICATION: AAD or repos (specific to this app)
//-----------------------------------------------------------------------------
const dynamicStartupInstance = getCompanySpecificDeployment();
dynamicStartupInstance?.routes?.api?.index && dynamicStartupInstance?.routes?.api?.rootIndex(router);

router.use('/:org', aadAndCustomProviders);

router.use('/:org', function (req: IApiRequest, res, next) {
  const orgName = req.params.org;
  if (!req.apiKeyToken.organizationScopes) {
    return next(jsonError('There is a problem with the key configuration (no organization scopes)', 412));
  }
  // '*'' is authorized for all organizations in this configuration environment
  if (!req.apiKeyToken.hasOrganizationScope(orgName)) {
    return next(jsonError('The key is not authorized for this organization', 401));
  }
  if (!req.apiKeyToken.scopes) {
    return next(jsonError('There is a problem with the key configuration (no specific API scopes)', 412));
  }
  if (!req.apiKeyToken.hasScope('createRepo')) {
    return next(jsonError('The key is not authorized to use the repo create APIs', 401));
  }

  const providers = getProviders(req);
  const operations = providers.operations;
  let organization = null;
  try {
    organization = operations.getOrganization(orgName);
  } catch (ex) {
    return next(jsonError(ex, 400));
  }
  req.organization = organization;
  return next();
});

router.post('/:org/repos', asyncHandler(async function (req: ReposAppRequest, res, next) {
  const providers = getProviders(req);
  const organization = req.organization;
  const convergedObject = Object.assign({}, req.headers);
  req.insights.trackEvent({ name: 'ApiRepoCreateRequest', properties: convergedObject });
  Object.assign(convergedObject, req.body);
  delete convergedObject.access_token;
  delete convergedObject.authorization;
  const logic = providers.customizedNewRepositoryLogic;
  const customContext = logic?.createContext(req);
  /*
  removed approvals from primary method:

  // Validate approval types
  const msApprovalType = msProperties.approvalType;
  if (!msApprovalType) {
    throw jsonError(new Error('Missing corporate approval type information'), 422);
  }
  if (hardcodedApprovalTypes.indexOf(msApprovalType) < 0) {
    throw jsonError(new Error('The provided approval type is not supported'), 422);
  }
  // Validate specifics of what is in the approval
  switch (msApprovalType) {
    case 'NewReleaseReview':
    case 'ExistingReleaseReview':
      if (!msProperties.approvalUrl) {
        throw jsonError(new Error('Approval URL for the release review is required when using the release review approval type'), 422);
      }
      break;
    case 'SmallLibrariesToolsSamples':
      break;
    case 'Exempt':
      if (!msProperties.justification) {
        throw jsonError(new Error('Justification is required when using the exempted approval type'), 422);
      }
      break;
    default:
      throw jsonError(new Error('The requested approval type is not currently supported.'), 422);
  }

  */
  try {
    const repoCreateResponse = await CreateRepository(req, organization, logic, customContext, convergedObject, CreateRepositoryEntrypoint.Api);
    res.status(201);
    req.insights.trackEvent({
      name: 'ApiRepoCreateRequestSuccess', properties: {
        request: JSON.stringify(convergedObject),
        response: JSON.stringify(repoCreateResponse),
      }
    });
    return res.json(repoCreateResponse);
  } catch (error) {
    const data = { ...convergedObject };
    data.error = error.message;
    data.encodedError = JSON.stringify(error);
    req.insights.trackEvent({ name: 'ApiRepoCreateFailed', properties: data });
    return next(error);
  }
}));

router.use(JsonErrorHandler);

export default router;
