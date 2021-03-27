//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Auth provider for this application's own local key system.

// As an API authentication provider, the request is augmented with a
// "apiKeyRow" object that provides additional information found
// while validating the key.

import basicAuth from 'basic-auth';
import crypto from 'crypto';

import { jsonError } from './jsonError';
import { getProviders, IProviders, ReposAppRequest } from '../transitional';
import { PersonalAccessToken } from '../entities/token/token';

export const wrapErrorForImmediateUserError = (err: Error) => {
  (err as any).immediate = true;
  return err;
};

export interface IApiRequest extends ReposAppRequest {
  apiKeyToken: PersonalAccessToken;
  apiKeyProviderName: string;
  apiVersion?: string;

  userContextOverwriteRequest?: any; // refactor?
}

export default function ReposApiAuthentication(req: IApiRequest, res, next) {
  const user = basicAuth(req);
  const key = user ? user.pass || user.name : null;
  if (!key) {
    return next(jsonError('No key supplied', 400));
  }

  const sha1 = crypto.createHash('sha1');
  sha1.update(key);
  const hashValue = sha1.digest('hex');

  const providers = getProviders(req);
  const tokenProvider = providers.tokenProvider;

  const apiEventProperties = {
    keyHash: hashValue,
    apiVersion: req.apiVersion,
    url: req.originalUrl || req.url,
    failed: undefined,
    message: undefined,
    statusCode: undefined,
    warning: undefined,
  };
  tokenProvider
    .getToken(hashValue)
    .then((token: PersonalAccessToken) => {
      return after(null, token);
    })
    .catch((error) => {
      return after(error, null);
    });

  function after(tokenError: any, token: PersonalAccessToken) {
    const eventName = 'ApiRequest' + (tokenError ? 'Denied' : 'Approved');
    const warning = token?.warning;
    if (tokenError) {
      apiEventProperties.failed = true;
      apiEventProperties.message = tokenError.message;
      apiEventProperties.statusCode = tokenError.statusCode;
    }
    if (tokenError && warning) {
      apiEventProperties.warning = warning;
    }
    const insights = getProviders(req).insights;
    insights?.trackEvent({ name: eventName, properties: apiEventProperties });
    if (tokenError) {
      insights?.trackMetric({ name: 'ApiInvalidKey', value: 1 });
      tokenError.skipLog = true;
      return next(
        jsonError(
          tokenError.statusCode === 404
            ? 'Key not authorized'
            : tokenError.message,
          401
        )
      );
    }
    if (token.isRevoked()) {
      tokenError = jsonError(warning || 'Key revoked', 403);
      wrapErrorForImmediateUserError(tokenError);
      tokenError.authErrorMessage = tokenError.message;
      insights?.trackMetric({ name: 'ApiRevokedKeyAttempt', value: 1 });
      return next(tokenError);
    }
    if (token.isExpired()) {
      tokenError = jsonError(
        warning || 'A revoked key attempted to use an API',
        403
      );
      wrapErrorForImmediateUserError(tokenError);
      tokenError.authErrorMessage = tokenError.message;
      insights?.trackMetric({ name: 'ApiExpiredKeyAttempt', value: 1 });
      return next(tokenError);
    }
    insights?.trackMetric({ name: 'ApiRequest', value: 1 });
    req.apiKeyToken = token;
    req.apiKeyProviderName = 'repos';
    return next();
  }
}
