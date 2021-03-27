//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import _ from 'lodash';

import { ICorporateLink } from './corporateLink';
import { OrganizationMember } from './organizationMember';
import {
  ICrossOrganizationMembersResult,
  ICrossOrganizationMembershipBasics,
  ICrossOrganizationMembershipByOrganization,
} from './operations';
import { IProviders, RequestTeamMemberAddType } from '../transitional';
import { TeamMember } from './teamMember';

const earlyProfileFetchTypes = new Set(['former', 'active', 'unknownAccount']);

const defaultPageSize = 33; // GitHub.com seems to use a value around 33
const earlyFetchPageBreak = 200;

export interface IMemberSearchOptions {
  providers: IProviders;

  organizationMembers?: OrganizationMember[];
  crossOrganizationMembers?: ICrossOrganizationMembersResult;

  isOrganizationScoped?: boolean;
  links?: ICorporateLink[];
  pageSize?: number;
  phrase?: string;
  type?: string; // TODO: should be an enum eventually
  orgId?: string | number;
  team2AddType?: RequestTeamMemberAddType;
  teamMembers?: TeamMember[];
}

export class MemberSearch {
  #providers: IProviders;

  public members: OrganizationMember[];
  public links: ICorporateLink[];
  public teamMembers: any;
  public team2AddType: RequestTeamMemberAddType;
  public pageSize: number;
  public phrase: string;
  public type: string;
  public page: number;
  public sort: string;
  public totalPages: number;
  public totalItems: number;
  public pageFirstItem: number;
  public pageLastItem: number;

  private orgId: string;

  constructor(options: IMemberSearchOptions) {
    if (!options.crossOrganizationMembers && !options.organizationMembers) {
      throw new Error(
        'Options must include either crossOrganizationMembers or organizationMembers'
      );
    }
    if (options.crossOrganizationMembers && options.organizationMembers) {
      throw new Error(
        'Options cannot include both crossOrganizationMembers or organizationMembers'
      );
    }
    if (options.organizationMembers) {
      this.members = options.organizationMembers;
    } else if (options.crossOrganizationMembers) {
      // must be a Map from ID to object with { orgs, memberships, account }
      this.members = (Array.from(
        options.crossOrganizationMembers.values()
      ) as any) as OrganizationMember[];
    }
    translateMembers(this.members, options.isOrganizationScoped, options.links);
    this.links = options.links;
    this.#providers = options.providers as IProviders;
    this.teamMembers = options.teamMembers;
    this.team2AddType = options.team2AddType;
    this.pageSize = options.pageSize || defaultPageSize;

    this.phrase = options.phrase;
    this.type = options.type;

    if (options.orgId) {
      this.orgId = String(options.orgId);
    }
  }

  async search(page, sort?: string): Promise<void> {
    this.page = parseInt(page);
    this.sort = sort
      ? sort.charAt(0).toUpperCase() + sort.slice(1)
      : 'Alphabet';

    await this.filterOrganizationOwners();
    await this.filterByTeamMembers()
      .associateLinks()
      .getCorporateProfilesEarly(this.type);
    return this.filterByType(this.type)
      .filterByPhrase(this.phrase)
      .determinePages()
      ['sortBy' + this.sort]()
      .getPage(this.page)
      .sortOrganizations()
      .getCorporateProfiles();
  }

  async filterOrganizationOwners() {
    const { queryCache, organizationMemberCacheProvider } = this.#providers;
    if (
      this.type === 'owners' &&
      queryCache.supportsOrganizationMembership &&
      organizationMemberCacheProvider
    ) {
      if (!this.orgId) {
        throw new Error(
          'org owners view not available at the top root level currently'
        );
      }
      const allOwners = await organizationMemberCacheProvider.queryAllOrganizationOwners();
      const owners = new Set<string>();
      for (const owner of allOwners) {
        if (this.orgId && owner.organizationId === this.orgId) {
          owners.add(owner.userId);
        }
      }
      this.members = this.members.filter((member) =>
        owners.has(String(member.id))
      );
    }
    return this;
  }

  filterByTeamMembers() {
    // If this is a sub-team view, filter by members unless the
    // special "add a member" experience is present in this route.
    let teamSet = new Set();
    if (this.teamMembers) {
      for (let i = 0; i < this.teamMembers.length; i++) {
        const member = this.teamMembers[i];
        teamSet.add(member.id);
      }
      if (this.team2AddType) {
        for (let i = 0; i < this.members.length; i++) {
          const member = this.members[i];
          // NOTE: this is not officially part of the interface
          // TODO: cleanup isTeamMember approach
          member['isTeamMember'] = teamSet.has(member.id);
        }
      } else {
        this.members = this.members.filter((m) => {
          return teamSet.has(m.id);
        });
      }
    }
    return this;
  }

  async getCorporateProfilesEarly(type): Promise<MemberSearch> {
    // This will make a Redis call for every single member, if not cached,
    // so the early mode is only used in a specific type of view this early.
    // The default just resolves for a single page of people.
    if (
      this.pageSize > earlyFetchPageBreak ||
      earlyProfileFetchTypes.has(type)
    ) {
      return await this.getCorporateProfiles();
    }
    return this;
  }

  associateLinks() {
    const links = new Map();
    this.links.forEach((link) => {
      const id = parseInt(link.thirdPartyId, 10);
      links.set(id, link);
    });
    this.members.forEach((member) => {
      const link = links.get(member.id);
      if (link) {
        member.link = link;
      }
    });
    return this;
  }

  sortOrganizations() {
    this.members.forEach((m) => {
      const member = (m as any) as ICrossOrganizationMembershipByOrganization;
      if (member.orgs && member.orgs.length > 0) {
        member.orgs = _.sortBy(member.orgs, ['name']);
      }
    });
    return this;
  }

  async getCorporateProfiles(): Promise<MemberSearch> {
    // const projectLinkAsCorporateProfile = this.type !== 'former';
    // corporate.alias -> corporateMailAddress (?)
    // corporate.emailAddress -> corporateAlias
    return this;
  }

  determinePages() {
    this.totalPages = Math.ceil(this.members.length / this.pageSize);
    this.totalItems = this.members.length;
    return this;
  }

  getPage(page) {
    this.members = this.members.slice(
      (page - 1) * this.pageSize,
      (page - 1) * this.pageSize + this.pageSize
    );
    this.pageFirstItem = 1 + (page - 1) * this.pageSize;
    this.pageLastItem = this.pageFirstItem + this.members.length - 1;
    return this;
  }

  filterByPhrase(phrase) {
    if (phrase) {
      phrase = phrase.toLowerCase();
      this.members = this.members.filter((m) => {
        return memberMatchesPhrase(m, phrase);
      });
    }
    return this;
  }

  filterByType(type) {
    let filter = null;
    switch (type) {
      case 'linked':
        filter = (r: OrganizationMember) => {
          return r.link && r.link.thirdPartyId;
        };
        break;
      case 'unlinked':
        filter = (r: OrganizationMember) => {
          return !r.link;
        };
        break;
      case 'unknownAccount':
        filter = (r: OrganizationMember) => {
          return (
            r.link &&
            r.link.thirdPartyId &&
            (!r.link || !r.link.corporateUsername)
          );
        };
        break;
      case 'former':
        filter = (r: OrganizationMember) => {
          return (
            r.link &&
            r.link.thirdPartyId &&
            !r.link.isServiceAccount &&
            (!r.link || !r.link.corporateUsername)
          );
        };
        break;
      case 'active':
        filter = (r: OrganizationMember) => {
          return (
            r.link &&
            r.link.thirdPartyId &&
            r.link.corporateId &&
            !r.link.isServiceAccount &&
            r.link &&
            r.link.corporateUsername
          );
        };
        break;
      case 'serviceAccount':
        filter = (r: OrganizationMember) => {
          return r.link && r.link.isServiceAccount;
        };
        break;
    }
    if (filter) {
      this.members = this.members.filter(filter);
    }
    return this;
  }

  sortByAlphabet() {
    this.members.sort((a, b) => {
      const aAccountIdentity = a.login
        ? a.login.toLowerCase()
        : a['account'].login.toLowerCase();
      const bAccountIdentity = b.login
        ? b.login.toLowerCase()
        : b['account'].login.toLowerCase();
      if (aAccountIdentity > bAccountIdentity) return 1;
      if (aAccountIdentity < bAccountIdentity) return -1;
      return 0;
    });
    return this;
  }
}

function translateMembers(members, isOrganizationScoped, optionalLinks) {
  // Support showing
  const linkedNoOrg = new Map();
  if (!isOrganizationScoped && optionalLinks) {
    for (let i = 0; i < optionalLinks.length; i++) {
      const link = optionalLinks[i] as ICorporateLink;
      if (link && link.thirdPartyUsername && link.thirdPartyId) {
        const id = parseInt(link.thirdPartyId, 10);
        linkedNoOrg.set(id, link);
      }
    }
  }
  // A breaking change altered the projected format
  members.forEach((member) => {
    linkedNoOrg.delete(member.id);
    if (member.orgs && !member.account) {
      const orgNames = Object.getOwnPropertyNames(member.orgs);
      const firstOrganization = orgNames[0];
      member.account = member.orgs[firstOrganization];
    }
  });
  // Locate linked users with no org memberships
  if (linkedNoOrg.size) {
    const noOrgs = Array.from(linkedNoOrg.values());
    for (let i = 0; i < noOrgs.length; i++) {
      const n = noOrgs[i];
      const thirdPartyId =
        n.thirdPartyId /* new link objects */ ||
        n.ghid; /* old implementation */
      const thirdPartyUsername = (
        n.thirdPartyUsername ||
        n.ghu ||
        ''
      ).toLowerCase();
      const thirdPartyAvatar = n.thirdPartyAvatar || n.ghavatar;
      const id = parseInt(thirdPartyId, 10);
      const newMember = {
        account: {
          avatar_url: thirdPartyAvatar,
          id,
          login: thirdPartyUsername || null,
        },
        id,
        orgs: {},
      };
      // Create a member entry
      members.push(newMember);
    }
  }
}

function memberMatchesPhrase(member, phrase) {
  const link = member.link as ICorporateLink;
  let linkIdentity = link
    ? `${link.corporateUsername} ${link.corporateDisplayName} ${link.corporateId} ${link.thirdPartyUsername} ${link.thirdPartyId} ${link.corporateMailAddress} ${link.corporateAlias}`
    : '';
  let accountIdentity = member.login
    ? member.login.toLowerCase()
    : member.account.login.toLowerCase();
  let combined = (linkIdentity + ' ' + accountIdentity).toLowerCase();
  return combined.includes(phrase);
}
