//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router, Response } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { getProviders } from '../transitional';
import { setIdentity } from '../middleware/business/authentication';
import { AddLinkToRequest } from '../middleware/links';
import { jsonError } from '../middleware';
import { apiContextMiddleware } from '../middleware/business/setContext';
import { ILocalExtensionKeyProvider } from '../entities/localExtensionKey';
import { LocalExtensionKey } from '../entities/localExtensionKey/localExtensionKey';
import { IApiRequest } from '../middleware/apiReposAuth';
import { PersonalAccessToken } from '../entities/token/token';

const thisApiScopeName = 'extension';

interface IExtensionResponse extends Response {
  localKey?: any;
}

interface IConnectionInformation {
  link?: any;
  operations?: any;
  auth?: any;
}

router.use(function (req: IApiRequest, res, next) {
  const token = req.apiKeyToken;
  if (!token.scopes) {
    return next(jsonError('The key is not authorized for specific APIs', 403));
  }
  if (!token.hasScope(thisApiScopeName)) {
    return next(jsonError('The key is not authorized to use the extension API', 403));
  }
  return next();
});

function overwriteUserContext(req: IApiRequest, res, next) {
  const token = req.apiKeyToken;
  const corporateId = token.corporateId;
  if (!corporateId) {
    return next(jsonError('No key owner', 403));
  }
  req.userContextOverwriteRequest = {
    user: {
      azure: {
        oid: corporateId,
      },
    },
  };
  return next();
}

// - - - Local middleware: use the retrieved API key row for identity  - - -
router.use(overwriteUserContext);
router.use(apiContextMiddleware);
// - - - Middleware: set the identities we have authenticated  - - -
router.use(setIdentity);
// - - - Middleware: resolve whether the corporate user has a link - - -
router.use(asyncHandler(AddLinkToRequest));
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

router.get('/', (req: IApiRequest, res) => {
  const { operations } = getProviders(req);

  // Basic info route, used to validate new users
  const apiContext = req.apiContext;

  const ghi = apiContext.getGitHubIdentity();
  const id = ghi ? ghi.id : null;
  const login = ghi ? ghi.username : null;

  const link = apiContext.link;

  // link display upn
  let displayUpn = link && link.corporateUsername ? link.corporateUsername : null;

  // vsts provider
  if (!displayUpn && req.apiKeyToken && req.apiKeyToken.displayUsername) {
    displayUpn = req.apiKeyToken.displayUsername;
  }

  const config = operations.config;

  const connectionInformation: IConnectionInformation = {};

  if (link) {
    // project into the connection info
    connectionInformation.link = {
      github: {
        id,
        login,
      },
      corporate: {
        preferredName: link.corporateDisplayName,
        userPrincipalName: link.corporateUsername,
        id: link.corporateId,
      },
    };
  }

  connectionInformation.operations = config.brand;

  // auth info
  if (req.apiKeyProviderName && req.apiKeyToken) {
    connectionInformation.auth = {
      provider: req.apiKeyProviderName,
      id: req.apiKeyToken.corporateId,
      upn: displayUpn,
    };
  }

  return res.json(connectionInformation);
});

router.get('/metadata', asyncHandler(getLocalEncryptionKeyMiddleware), (req: IApiRequest, res: IExtensionResponse) => {
  const apiContext = req.apiContext;

  const localKey = res.localKey;
  const { operations } = getProviders(req);
  const ghi = apiContext.getGitHubIdentity();
  const id = ghi ? ghi.id : null;
  const login = ghi ? ghi.username : null;
  const link = apiContext.link;
  const orgData = getSanitizedOrganizations(operations);
  const config = operations.config;

  const metadata = {
    extension: {
      localEncryptionKey: localKey,
    },
    operations: config.brand,
    serviceMessage: config.serviceMessage,
    reference: config.corporate.trainingResources ? config.corporate.trainingResources.footer : {},
    organizations: orgData,
    site: config.urls,
    link: undefined,
  };

  if (link) {
    metadata.link = {
      github: {
        id,
        login,
      },
      corporate: {
        preferredName: link.corporateDisplayName,
        userPrincipalName: link.corporateUsername,
        id: link.corporateId,
      },
    };
  }

  res.json(metadata);
});

function getSanitizedOrganizations(operations) {
  const value = {
    list: operations.organizationNames,
    settings: {},
  };
  value.list.map(function (organizationName) {
    const organization = operations.getOrganization(organizationName);
    const basics = {
      locked: organization.locked,
      createRepositoriesOnGitHub: organization.createRepositoriesOnGitHub,
      privateEngineering: organization.privateEngineering,
      externalMembersPermitted: organization.externalMembersPermitted,
      description: organization.description,
      priority: organization.priority,
      entities: organization.legalEntities,
      // broadAccessTeams: organization.broadAccessTeams,
      // systemTeamIds: organization.systemTeamIds,
    };
    value.settings[organizationName] = basics;
  });
  return value;
}

async function getLocalEncryptionKeyMiddleware(req: IApiRequest, res, next): Promise<void> {
  const providers = getProviders(req);
  const localExtensionKeyProvider = providers.localExtensionKeyProvider;
  const apiKeyToken = req.apiKeyToken;
  const insights = req.insights;
  try {
    const key = await getOrCreateLocalEncryptionKey(insights, localExtensionKeyProvider, apiKeyToken);
    if (!key) {
      throw new Error('No local extension key could be generated');
    }
    res.localKey = key;
  } catch (error) {
    return next(jsonError(error, 500));
  }
  return next();
}

async function getLocalEncryptionKey(localExtensionKeyProvider: ILocalExtensionKeyProvider, corporateId: string): Promise<string> {
  try {
    const localEncryptionKey = await localExtensionKeyProvider.getForCorporateId(corporateId);
    if (localEncryptionKey.isValidNow()) {
      return localEncryptionKey.localDataKey;
    }
    await localExtensionKeyProvider.delete(localEncryptionKey);
  } catch (error) {
    if (error && ((error.statusCode && error.statusCode === 404) || (error.status && error.status === 404))) {
      return null;
    }
    throw error;
  }
  return null;
}

async function createLocalEncryptionKey(insights, localExtensionKeyProvider: ILocalExtensionKeyProvider, corporateId: string): Promise<string> {
  const localEncryptionKey = LocalExtensionKey.CreateNewLocalExtensionKey(corporateId);
  await localExtensionKeyProvider.createNewForCorporateId(localEncryptionKey);
  insights.trackEvent({ name: 'ExtensionNewLocalKeyGenerated' });
  insights.trackMetric({ name: 'ExtensionNewLocalKeys', value: 1 });
  return localEncryptionKey.localDataKey;
}

async function getOrCreateLocalEncryptionKey(insights, localExtensionKeyProvider: ILocalExtensionKeyProvider, apiKeyToken: PersonalAccessToken): Promise<string> {
  const corporateId = apiKeyToken.corporateId; // apiKeyRow.RowKey || apiKeyRow.owner;
  if (!corporateId) {
    throw new Error('Owner identity required');
  }
  const localDataKey = await getLocalEncryptionKey(localExtensionKeyProvider, corporateId);
  if (localDataKey) {
    return localDataKey;
  }
  return await createLocalEncryptionKey(insights, localExtensionKeyProvider, corporateId);
}

router.use('*', (req, res, next) => {
  return next(jsonError('API not found', 404));
});

export default router;
