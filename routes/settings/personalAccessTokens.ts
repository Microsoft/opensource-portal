//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { getProviders } from '../../transitional';
import { PersonalAccessToken } from '../../entities/token/token';
import { ReposAppRequest } from '../../interfaces';

interface IPersonalAccessTokenForDisplay {
  active: boolean;
  expired: boolean;
  expires: string;
  identifier: string;
  description: string;
  apis: string[];
  tokenEntity: PersonalAccessToken;
}

export interface IRequestForSettingsPersonalAccessTokens extends ReposAppRequest {
  personalAccessTokens?: IPersonalAccessTokenForDisplay[];
}

const serviceName = 'repos-pat';
const tokenExpirationMs = 1000 * 60 * 60 * 24 * 365; // 365 days

function translateTableToEntities(personalAccessTokens: PersonalAccessToken[]): IPersonalAccessTokenForDisplay[] {
  return personalAccessTokens.map(pat => {
    // So that we do not share the hashed key with the user, we
    // build a hash of that and the timestamp to offer a single-version
    // tag to use for delete operations, etc.
    const displayToken: IPersonalAccessTokenForDisplay = {
      active: pat.active,
      expired: pat.isExpired(),
      expires: pat.expires ? pat.expires.toDateString() : null,
      description: pat.description,
      apis: pat.scopes ? pat.scopes.split(',') : [],
      identifier: pat.getIdentifier(),
      tokenEntity: pat,
    };
    return displayToken;
  });
}

function getPersonalAccessTokens(req: ReposAppRequest, res, next) {
  const providers = getProviders(req);
  const tokenProvider = providers.tokenProvider;
  const corporateId = req.individualContext.corporateIdentity.id;
  tokenProvider.queryTokensForCorporateId(corporateId).then(tokens => {
    req['personalAccessTokens'] = translateTableToEntities(tokens);
    return next();
  }).catch(error => {
    return next(error);
  });
}

function view(req: IRequestForSettingsPersonalAccessTokens, res) {
  const personalAccessTokens = req.personalAccessTokens;
  req.individualContext.webContext.render({
    view: 'settings/personalAccessTokens',
    title: 'Personal access tokens',
    state: {
      personalAccessTokens,
      newKey: res.newKey,
      isPreviewUser: true, //req.isPreviewUser,
    },
  });
}

router.use(getPersonalAccessTokens);

router.get('/', view);

function createToken(req: ReposAppRequest, res, next) {
  const providers = getProviders(req);
  const tokenProvider = providers.tokenProvider;
  const insights = req.insights;
  const description = req.body.description;
  if (!description) {
    return next(new Error('A description is required to create a new Personal Access Token'));
  }
  const corporateId = req.individualContext.corporateIdentity.id;
  const token = PersonalAccessToken.CreateNewToken();
  token.corporateId = corporateId;
  token.description = description;
  const now = new Date();
  token.expires = new Date(now.getTime() + tokenExpirationMs);
  token.source = serviceName;
  token.scopes = 'extension,links';
  insights.trackEvent({
    name: 'ReposCreateTokenStart',
    properties: {
      id: corporateId,
      description: description,
    },
  });
  tokenProvider.saveNewToken(token).then(ok => {
    insights.trackEvent({
      name: 'ReposCreateTokenFinish',
      properties: {
        id: corporateId,
        description: description,
      },
    });
    const newKey = token.getPrivateKey();
    getPersonalAccessTokens(req, res, () => {
        res.newKey = newKey;
        return view(req, res);
      });
  }).catch(insertError => {
    insights.trackEvent({
      name: 'ReposCreateTokenFailure',
      properties: {
        id: corporateId,
        description: description,
      },
    });
    return next(insertError);
  });
}

router.post('/create', createToken);
router.post('/extension', createToken);

router.post('/delete', asyncHandler(async (req: IRequestForSettingsPersonalAccessTokens, res, next) => {
  const providers = getProviders(req);
  const tokenProvider = providers.tokenProvider;
  const revokeAll = req.body.revokeAll === '1';
  const revokeIdentifier = req.body.revoke;
  const personalAccessTokens = req.personalAccessTokens;
  for (const pat of personalAccessTokens) {
    const token = pat.tokenEntity;
    if (revokeAll || pat.identifier === revokeIdentifier) {
      token.active = false;
      await tokenProvider.updateToken(token);
    }
  }
  return res.redirect('/settings/security/tokens');
}));

export default router;
