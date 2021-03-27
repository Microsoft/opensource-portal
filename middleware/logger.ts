//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import logger from 'morgan';

import { getProviders, ReposAppRequest } from '../transitional';

const encryptionMetadataKey = '_ClientEncryptionMetadata2';
const piiFormat =
  ':id :method :scrubbedUrl :status :response-time ms - :res[content-length] :encryptedSession :correlationId';
const format =
  ':method :scrubbedUrl :status :response-time ms - :res[content-length] :encryptedSession :correlationId';

logger.token('encryptedSession', function getUserId(req: ReposAppRequest) {
  const config = getProviders(req).config;
  if (req.session) {
    const sessionPassport = (req.session as any).passport;
    if (sessionPassport && sessionPassport.user) {
      const userType =
        config.authentication.scheme === 'aad' ? 'azure' : 'github';
      return sessionPassport.user[userType] &&
        sessionPassport.user[userType][encryptionMetadataKey] !== undefined
        ? 'encrypted'
        : 'plain';
    }
  }
});

logger.token('id', function getUserId(req: ReposAppRequest) {
  const config = getProviders(req).config;
  if (config) {
    const userType =
      config.authentication.scheme === 'aad' ? 'azure' : 'github';
    return req.user && req.user[userType] && req.user[userType].username
      ? req.user[userType].username
      : undefined;
  }
});

logger.token('correlationId', function getCorrelationId(req: ReposAppRequest) {
  return req.correlationId;
});

logger.token('scrubbedUrl', function getScrubbedUrl(req: ReposAppRequest) {
  return req.scrubbedUrl || req.originalUrl || req.url;
});

export default function createLogger(config) {
  return logger(
    config && config.debug && config.debug.showUsers === true
      ? piiFormat
      : format
  );
}
