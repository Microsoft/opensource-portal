//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  Organization,
  MemberSearch,
  ICrossOrganizationMembersResult,
  Operations,
} from '../../business';
import { getProviders, ReposAppRequest } from '../../transitional';
import LeakyLocalCache, { getLinksLightCache } from './leakyLocalCache';

// BAD PRACTICE: leaky local cache
// CONSIDER: use a better approach
const leakyLocalCachePeople = new LeakyLocalCache<
  boolean,
  ICrossOrganizationMembersResult
>();

export async function getPeopleAcrossOrganizations(operations: Operations) {
  const value = leakyLocalCachePeople.get(true);
  if (value) {
    return { crossOrganizationMembers: value };
  }
  const crossOrganizationMembers = await operations.getMembers();
  leakyLocalCachePeople.set(true, crossOrganizationMembers);
  return { crossOrganizationMembers };
}

export async function equivalentLegacyPeopleSearch(req: ReposAppRequest) {
  const { operations } = getProviders(req);
  const links = await getLinksLightCache(operations);
  const org = req.organization ? req.organization.name : null;
  const orgId = req.organization ? (req.organization as Organization).id : null;
  const { crossOrganizationMembers } = await getPeopleAcrossOrganizations(
    operations
  );
  const page = req.query.page_number ? Number(req.query.page_number) : 1;
  let phrase = req.query.q as string;
  let type = req.query.type as string;
  const validTypes = new Set([
    'linked',
    'active',
    'unlinked',
    'former',
    'serviceAccount',
    'unknownAccount',
    'owners',
  ]);
  if (!validTypes.has(type)) {
    type = null;
  }
  const filters = [];
  if (type) {
    filters.push({
      type: 'type',
      value: type,
      displayValue: type === 'former' ? 'formerly known' : type,
      displaySuffix: 'members',
    });
  }
  if (phrase) {
    filters.push({
      type: 'phrase',
      value: phrase,
      displayPrefix: 'matching',
    });
  }
  const search = new MemberSearch({
    phrase,
    type,
    pageSize: 1000000, // we'll slice it locally
    links,
    providers: operations.providers,
    orgId,
    crossOrganizationMembers,
    isOrganizationScoped: false,
  });
  await search.search(page, req.query.sort as string);
  return search;
}
