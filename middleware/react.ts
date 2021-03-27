//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import fs from 'fs';
import path from 'path';

import appPackage from '../package.json';

import { getStaticBlobCacheFallback } from '../lib/staticBlobCacheFallback';
import { getProviders, ReposAppRequest } from '../transitional';

const staticReactPackageNameKey = 'static-react-package-name';
const staticClientPackageName = appPackage[staticReactPackageNameKey];

export function injectReactClient() {
  let indexPageContent = '';
  try {
    if (!staticClientPackageName) {
      throw new Error(
        `No property "${staticReactPackageNameKey}" in package.json to inject as a React client app`
      );
    }
    const staticModernReactApp = require(staticClientPackageName);
    const previewClientFolder = staticModernReactApp;
    if (typeof previewClientFolder !== 'string') {
      throw new Error(
        `The return value of the preview package ${staticClientPackageName} must be a string/path`
      );
    }
    indexPageContent = fs.readFileSync(
      path.join(previewClientFolder, 'client.html'),
      { encoding: 'utf8' }
    );
  } catch (hostClientError) {
    console.error(
      `The static client could not be loaded via package ${staticClientPackageName}. Note that index.html needs to be named client.html in build/.`
    );
    throw hostClientError;
  }
  return function injectedRoute(req: ReposAppRequest, res, next) {
    // special passthrough
    if (req.path.includes('/byClient')) {
      return next();
    }

    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    res.header('x-ms-repos-site', 'react');
    return res.send(indexPageContent);
  };
}

type CacheBuffer = {
  buffer: Buffer;
  contentType: string;
};
const localFallbackBlobCache = new Map<string, CacheBuffer>();

export async function TryFallbackToBlob(
  req: ReposAppRequest,
  res: express.Response
): Promise<boolean> {
  if (!req.path) {
    return false;
  }
  const providers = getProviders(req);
  const baseUrl = '/react' + req.originalUrl;
  if (localFallbackBlobCache.has(baseUrl)) {
    providers.insights.trackEvent({
      name: 'FallbackToBlob',
      properties: { baseUrl },
    });
    const entry = localFallbackBlobCache.get(baseUrl);
    if (entry.contentType) {
      res.contentType(entry.contentType);
    }
    res.send(entry.buffer);
    return true;
  }
  const fallbackBlob = await getStaticBlobCacheFallback(providers);
  const [buffer, contentType] = await fallbackBlob.get(baseUrl);
  if (buffer) {
    providers.insights.trackEvent({
      name: 'FallbackToBlob',
      properties: { baseUrl },
    });
    localFallbackBlobCache.set(baseUrl, { buffer, contentType });
    if (contentType) {
      res.contentType(contentType);
    }
    res.send(buffer);
    return true;
  }
  return false;
}
