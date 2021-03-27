//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import axios from 'axios';
import asyncHandler from 'express-async-handler';

import { jsonError } from './jsonError';
import { IApiRequest } from './apiReposAuth';
import { PersonalAccessToken } from '../entities/token/token';
import { getProviders } from '../transitional';

// TODO: consider better caching
const localMemoryCacheVstsToAadId = new Map();

const vstsAuth = asyncHandler(async (req: IApiRequest, res, next) => {
  const config = getProviders(req).config;
  if (!config) {
    return next(new Error('Missing configuration for the application'));
  }
  if (!config.authentication || !config.authentication.vsts) {
    return next(
      new Error(
        'No VSTS authentication configuration available, VSTS authentication is not supported'
      )
    );
  }
  if (config.authentication.vsts.enabled !== true) {
    return next(
      new Error(
        'VSTS authentication is not enabled in the current configuration'
      )
    );
  }
  if (!config.authentication.vsts.vstsCollectionUrl) {
    return next(
      new Error(
        'VSTS collection URL is missing in the environment configuration'
      )
    );
  }
  const { graphProvider } = getProviders(req);
  const vstsCollectionUrl = config.authentication.vsts.vstsCollectionUrl;
  const connectionDataApi = `${vstsCollectionUrl}/_apis/connectiondata`;
  const authorizationHeader = req.headers.authorization;
  async function translateVstsUpnToAadId(upn: string) {
    let cached = localMemoryCacheVstsToAadId.get(upn);
    if (cached) {
      return cached;
    }
    const id = await graphProvider.getUserIdByUsername(upn);
    return id;
  }
  try {
    const response = await axios({
      url: connectionDataApi,
      headers: {
        Authorization: authorizationHeader,
        'X-TFS-FedAuthRedirect': 'Suppress',
      },
    });
    const body = response.data;
    if (!body.authenticatedUser || !body.authenticatedUser.isActive) {
      const error = jsonError(
        'The user is no longer active or authenticated',
        401
      );
      error['authErrorMessage'] = error.message;
      return next(error);
    }
    const displayName =
      body.authenticatedUser.providerDisplayName || 'Authenticated User';
    if (
      !body.authenticatedUser.properties ||
      !body.authenticatedUser.properties.Account
    ) {
      const error = jsonError(
        'Authenticated user information is not available from VSTS',
        401
      );
      error['authErrorMessage'] = error.message;
      return next(error);
    }
    if (
      body.authenticatedUser.properties.Account['$type'] !== 'System.String'
    ) {
      const error = jsonError(
        'Authenticated user type from VSTS is not supported',
        401
      );
      error['authErrorMessage'] = error.message;
      return next(error);
    }
    const upn = body.authenticatedUser.properties.Account['$value'];
    const id = await translateVstsUpnToAadId(upn);
    // IMPORTANT: for our use in the extension, apiKeyRow.owner is an AAD ID and is
    // the primary way to make sure things are good for now...
    const token = PersonalAccessToken.CreateFromAzureDevOpsTokenAuthorization({
      corporateId: id,
      source: 'vsts-pat',
      description: `Azure DevOps Personal Access Token for ${displayName}`,
      displayUsername: upn,
      scopes: 'extension,links',
    });
    req.apiKeyToken = token;
    req.apiKeyProviderName = 'vsts';
    return next();
  } catch (error) {
    const err = jsonError(
      `You are not authorized to access the resource via Azure DevOps`,
      401
    );
    err['authErrorMessage'] = error.message;
    err['skipLog'] = true;
    return next(err);
  }
});

export default vstsAuth;
