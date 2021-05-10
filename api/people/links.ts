//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';

import { jsonError } from '../../middleware';
import { ICrossOrganizationMembersResult, MemberSearch, Operations } from '../../business';
import { ICorporateLink } from '../../interfaces';
import { IApiRequest } from '../../middleware/apiReposAuth';
import postLinkApi from './link';
import { ErrorHelper, getProviders } from '../../transitional';
import { wrapError } from '../../utils';

const router: Router = Router();

const unsupportedApiVersions = [
  '2016-12-01',
];

const extendedLinkApiVersions = [
  '2019-02-01',
];

router.use(function (req: IApiRequest, res, next) {
  const token = req.apiKeyToken;
  if (!token.scopes) {
    return next(jsonError('The key is not authorized for specific APIs', 401));
  }
  if (!token.hasScope('links') && !token.hasScope('link')) {
    return next(jsonError('The key is not authorized to use the links API', 401));
  }
  return next();
});

router.post('/', asyncHandler(postLinkApi));

router.get('/', asyncHandler(async (req: IApiRequest, res, next) => {
  const { operations } = getProviders(req);
  const skipOrganizations = req.query.showOrganizations !== undefined && !!req.query.showOrganizations;
  const showTimestamps = req.query.showTimestamps !== undefined && req.query.showTimestamps === 'true';
  const results = await getAllUsers(req.apiVersion, operations, skipOrganizations, showTimestamps);
  req.insights.trackMetric({ name: 'ApiRequestLinks', value: 1 });
  res.set('Content-Type', 'application/json');
  res.send(JSON.stringify(results, undefined, 2));
}));

router.get('/:linkid', asyncHandler(async (req: IApiRequest, res, next) => {
  if (unsupportedApiVersions.includes(req.apiVersion)) {
    return next(jsonError('This API is not supported by the API version you are using.', 400));
  }
  const linkid = req.params.linkid.toLowerCase();
  const { operations } = getProviders(req);
  const skipOrganizations = req.query.showOrganizations !== undefined && !!req.query.showOrganizations;
  const showTimestamps = req.query.showTimestamps !== undefined && req.query.showTimestamps === 'true';
  if (operations.providers.queryCache && operations.providers.queryCache.supportsOrganizationMembership) {
    // faster implementation
    const links = (await operations.providers.linkProvider.getAll()).filter(lid => lid['id'] === linkid);
    let link = links.length === 1 ? links[0] : null;
    if (!link) {
      return next(jsonError('Could not find the link', 404));
    }
    let entry = null;
    const thirdPartyId = link.thirdPartyId;
    try {
      entry = await getByThirdPartyId(thirdPartyId, req.apiVersion, operations, skipOrganizations, showTimestamps);
    } catch (error) {
      if (ErrorHelper.IsNotFound(error)) {
        return next(jsonError('Could not find the link', 404));
      } else {
        return next(jsonError(error, 500));
      }
    }
    req.insights.trackMetric({ name: 'ApiRequestLinkByLinkId', value: 1 });
    return res.json(entry);
  }
  const results = await getAllUsers(req.apiVersion, operations, skipOrganizations, showTimestamps, true);
  for (let i = 0; i < results.length; i++) {
    const entry = results[i];
    if (entry && entry.id === linkid) {
      req.insights.trackMetric({ name: 'ApiRequestLinkByLinkId', value: 1 });
      return res.json(entry);
    }
  }
  return next(jsonError('Could not find the link', 404));
}));

router.get('/github/:username', asyncHandler(async (req: IApiRequest, res, next) => {
  if (unsupportedApiVersions.includes(req.apiVersion)) {
    return next(jsonError('This API is not supported by the API version you are using.', 400));
  }
  const username = req.params.username.toLowerCase();
  const { operations } = getProviders(req);
  const skipOrganizations = req.query.showOrganizations !== undefined && !!req.query.showOrganizations;
  const showTimestamps = req.query.showTimestamps !== undefined && req.query.showTimestamps === 'true';
  if (operations.providers.queryCache && operations.providers.queryCache.supportsOrganizationMembership) {
    // faster implementation
    let account = null;
    try {
      account = await operations.getAccountByUsername(username);
    } catch (getAccountError) {
      if (ErrorHelper.IsNotFound(account)) {
        return next(jsonError('Could not find a link for the user', 404));
      }
      return next(jsonError(getAccountError, 500));
    }
    try {
      const entry = await getByThirdPartyId(String(account.id), req.apiVersion, operations, skipOrganizations, showTimestamps);
      req.insights.trackMetric({ name: 'ApiRequestLinkByGitHubUsername', value: 1 });
      return res.json(entry);
    } catch (entryError) {
      return next(jsonError(entryError, ErrorHelper.GetStatus(entryError) || 500));
    }
  }
  const results = await getAllUsers(req.apiVersion, operations, skipOrganizations, showTimestamps);
  for (let i = 0; i < results.length; i++) {
    const entry = results[i];
    if (entry && entry.github && entry.github.login.toLowerCase() === username) {
      req.insights.trackMetric({ name: 'ApiRequestLinkByGitHubUsername', value: 1 });
      return res.json(entry);
    }
  }
  return next(jsonError('Could not find a link for the user', 404));
}));

router.get('/aad/userPrincipalName/:upn', asyncHandler(async (req: IApiRequest, res, next) => {
  const upn = req.params.upn;
  const { operations } = getProviders(req);
  const skipOrganizations = req.query.showOrganizations !== undefined && !!req.query.showOrganizations;
  const showTimestamps = req.query.showTimestamps !== undefined && req.query.showTimestamps === 'true';
  if (operations.providers.queryCache && operations.providers.queryCache.supportsOrganizationMembership) {
    // faster implementation
    const links = await operations.providers.linkProvider.queryByCorporateUsername(upn);
    const r = [];
    for (const link of links) {
      const thirdPartyId = link.thirdPartyId;
      try {
        const entry = await getByThirdPartyId(thirdPartyId, req.apiVersion, operations, skipOrganizations, showTimestamps);
        if (entry) {
          r.push(entry);
        }
      } catch (partialIgnoreError) {
        if (!ErrorHelper.IsNotFound(partialIgnoreError)) {
          console.dir(partialIgnoreError);
        }
      }
    }
    req.insights.trackEvent({
      name: 'ApiRequestLinkByAadUpnResult',
      properties: {
        length: r.length.toString(),
        userPrincipalName: upn,
      },
    });
    return res.json(r);
  }
  const results = await getAllUsers(req.apiVersion, operations, skipOrganizations, showTimestamps);
  let r = [];
  for (let i = 0; i < results.length; i++) {
    const entry = results[i];
    if (entry && entry.aad && entry.aad.userPrincipalName === upn) {
      r.push(entry);
    }
  }
  req.insights.trackEvent({
    name: 'ApiRequestLinkByAadUpnResult',
    properties: {
      length: r.length.toString(),
      userPrincipalName: upn,
    },
  });
  if (r.length === 0) {
    return next(jsonError('Could not find a link for the user', 404));
  }
  req.insights.trackMetric({ name: 'ApiRequestLinkByAadUpn', value: 1 });
  return res.json(r);
}));

router.get('/aad/:id', asyncHandler(async (req: IApiRequest, res, next) => {
  if (req.apiVersion == '2016-12-01') {
    return next(jsonError('This API is not supported by the API version you are using.', 400));
  }
  const id = req.params.id;
  const skipOrganizations = req.query.showOrganizations !== undefined && !!req.query.showOrganizations;
  const showTimestamps = req.query.showTimestamps !== undefined && req.query.showTimestamps === 'true';
  const { operations } = getProviders(req);
  if (operations.providers.queryCache && operations.providers.queryCache.supportsOrganizationMembership) {
    // faster implementation
    const links = await operations.providers.linkProvider.queryByCorporateId(id);
    const r = [];
    for (const link of links) {
      const thirdPartyId = link.thirdPartyId;
      try {
        const entry = await getByThirdPartyId(thirdPartyId, req.apiVersion, operations, skipOrganizations, showTimestamps);
        if (entry) {
          r.push(entry);
        }
      } catch (partialIgnoreError) {
        if (!ErrorHelper.IsNotFound(partialIgnoreError)) {
          console.dir(partialIgnoreError);
        }
      }
    }
    req.insights.trackMetric({ name: 'ApiRequestLinkByAadId', value: 1 });
    return res.json(r);
  }
  const results = await getAllUsers(req.apiVersion, operations, skipOrganizations, showTimestamps);
  let r = [];
  for (let i = 0; i < results.length; i++) {
    const entry = results[i];
    if (entry && entry.aad && entry.aad.id === id) {
      r.push(entry);
    }
  }
  if (r.length === 0) {
    return next(jsonError('Could not find a link for the user', 404));
  }
  req.insights.trackMetric({ name: 'ApiRequestLinkByAadId', value: 1 });
  return res.json(r);
}));

async function getByThirdPartyId(thirdPartyId: string, apiVersion, operations: Operations, skipOrganizations: boolean, showTimestamps: boolean, showLinkIds?: boolean): Promise<any> {
  const providers = operations.providers;
  const { graphProvider } = providers;
  let link: ICorporateLink = null;
  try {
    link = await providers.linkProvider.getByThirdPartyId(thirdPartyId);
  } catch (linksError) {
    if (ErrorHelper.IsNotFound(linksError)) {
      throw jsonError(`${thirdPartyId} is not linked`, 404);
    } else {
      linksError = wrapError(linksError, 'There was a problem retrieving link information to display alongside the member.');
      throw jsonError(linksError, 500);
    }
  }
  const account = operations.getAccount(thirdPartyId);
  await account.getDetails();
  let orgMembershipNames: string[] = [];
  if (providers.queryCache && operations.providers.queryCache.supportsOrganizationMembership) {
    orgMembershipNames = (await providers.queryCache.userOrganizations(thirdPartyId)).map(org => org.organization.name);
  } else {
    // TODO: not implemented for performance reasons now
    throw ErrorHelper.NotImplemented();
  }
  const isExpandedView = extendedLinkApiVersions.includes(apiVersion);
  const entry = {
    github: {
      id: Number(link.thirdPartyId),
      login: account.login,
      organizations: undefined,
    },
    isServiceAccount: undefined,
    serviceAccountContact: undefined,
  };
  if (isExpandedView) {
    entry.github['avatar'] = account.avatar_url;
  }
  if (showLinkIds && link) {
    entry['id'] = link['id']; // not part of the interface
  }
  if (!skipOrganizations) {
    entry.github.organizations = orgMembershipNames;
  }
  // '2017-09-01' added 'isServiceAccount'; so '2016-12-01' & '2017-03-08' do not have it
  if (showTimestamps && link && link['created']) {
    entry['timestamp'] = link['created'];
  }
  if (link && link.isServiceAccount === true && apiVersion !== '2016-12-01' && apiVersion !== '2017-03-08') {
    entry.isServiceAccount = true;
    if (isExpandedView && link.isServiceAccount && link.serviceAccountMail) {
      entry.serviceAccountContact = link.serviceAccountMail;
    }
  }
  if (link?.corporateAlias || link?.corporateDisplayName || link?.corporateMailAddress || link?.corporateUsername) {
    const corporatePropertyName = apiVersion === '2016-12-01' ? 'corporate' : 'aad'; // This was renamed to be provider name-based
    entry[corporatePropertyName] = {
      alias: link?.corporateAlias,
      preferredName: link?.corporateDisplayName,
      userPrincipalName: link?.corporateUsername,
      emailAddress: link?.corporateMailAddress,
    };
    const corporateIdPropertyName = apiVersion === '2016-12-01' ? 'aadId' : 'id'; // Now just 'id'
    entry[corporatePropertyName][corporateIdPropertyName] = link.corporateId;
  }
  return entry;
}

async function getAllUsers(apiVersion, operations: Operations, skipOrganizations: boolean, showTimestamps: boolean, showLinkIds?: boolean): Promise<any[]> {
  let links: ICorporateLink[] = null;
  try {
    links = await operations.getLinks();
  } catch (linksError) {
    linksError = wrapError(linksError, 'There was a problem retrieving link information to display alongside members.');
    throw jsonError(linksError, 500);
  }
  let crossOrganizationMembers: ICrossOrganizationMembersResult;
  try {
    // TODO: this is a cross-org map!? validate return type...
    crossOrganizationMembers = await operations.getMembers();
  } catch (error) {
    error = wrapError(error, 'There was a problem getting the members list.');
    throw jsonError(error, 500);
  }
  const search = new MemberSearch({
    crossOrganizationMembers,
    type: 'linked',
    links,
    providers: operations.providers,
    pageSize: Number.MAX_SAFE_INTEGER,
  });
  try {
    await search.search(1);
    const sr = search.members;
    const isExpandedView = extendedLinkApiVersions.includes(apiVersion);
    const results = [];
    sr.forEach(member => {
      const entry = {
        github: {
          id: member['account'].id,
          login: member['account'].login,
          organizations: undefined,
        },
        isServiceAccount: undefined,
        serviceAccountContact: undefined,
      };
      if (isExpandedView) {
        entry.github['avatar'] = member['account'].avatar_url;
      }
      if (showLinkIds && member && member.link && member.link['id']) {
        entry['id'] = member.link['id'];
      }
      if (!skipOrganizations && member['orgs']) {
        entry.github.organizations = Object.getOwnPropertyNames(member['orgs']);
      }
      // '2017-09-01' added 'isServiceAccount'; so '2016-12-01' & '2017-03-08' do not have it
      const link = member.link as ICorporateLink;
      if (showTimestamps && link && link['created']) {
        entry['timestamp'] = link['created'];
      }
      if (link && link.isServiceAccount === true && apiVersion !== '2016-12-01' && apiVersion !== '2017-03-08') {
        entry.isServiceAccount = true;
        if (isExpandedView && link.isServiceAccount && link.serviceAccountMail) {
          entry.serviceAccountContact = link.serviceAccountMail;
        }
      }
      const corporate = member.link;
      if (corporate) {
        const corporatePropertyName = apiVersion === '2016-12-01' ? 'corporate' : 'aad'; // This was renamed to be provider name-based
        entry[corporatePropertyName] = {
          alias: corporate.corporateAlias,
          preferredName: corporate.corporateDisplayName,
          userPrincipalName: corporate.corporateUsername,
          emailAddress: corporate.corporateMailAddress,
        };
        const corporateIdPropertyName = apiVersion === '2016-12-01' ? 'aadId' : 'id'; // Now just 'id'
        entry[corporatePropertyName][corporateIdPropertyName] = corporate.corporateId;
      }
      results.push(entry);
    });
    return results;
  } catch (error) {
    throw jsonError(error, 400);
  }
}

export default router;
