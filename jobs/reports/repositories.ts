//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/* eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import _ from 'lodash';

import AutomaticTeamsWebhookProcessor from '../../webhooks/tasks/automaticTeams';

import moment from 'moment';

const querystring = require('querystring');

import {
  Repository,
  GitHubCollaboratorAffiliationQuery,
} from '../../business/repository';
import { requireJson, sleep } from '../../utils';
import { Operations } from '../../business/operations';
import { Organization } from '../../business/organization';
import { IReportsContext } from './task';
import { Collaborator } from '../../business/collaborator';
import { ICorporateLink } from '../../business/corporateLink';
import { Team } from '../../business/team';
import { TeamPermission } from '../../business/teamPermission';

const projectPlaceholder = '[project]\\';

const providerName = 'repositories';
const definitions = requireJson('jobs/reports/repositoryDefinitions.json');
const exemptRepositories = requireJson('jobs/reports/exemptRepositories.json');
const definitionsByName = {};
for (let i = 0; i < definitions.length; i++) {
  const definition = definitions[i];
  definitionsByName[definition.name] = definition;
}

const simpleDateFormat = 'l';

interface IReportsRepositoryContext {
  parent: any;

  definitionsUsed: Set<any>;
  issues: any;

  name: string;
  nameLowercase: string;

  repository: Repository;

  administrators?: any;
  actionableAdministrators?: any;

  countOfAdministratorCollaborators: number;
  countOfAdministratorTeams: number;

  administratorsByType?: any;

  recipients?: any;
  additionalRecipients?: any;
}

interface IReportEntry {
  rows?: any[];
  listItems?: any[];
}

interface IUserEntry {
  login: string;
  reasons: {
    memberships: Team[];
    directCollaborator: boolean;
    collaborator: boolean;
  };
}

interface IReportActionLink {
  color?: string;
  text: string;
  link?: string;
}

interface IBasicRepository {
  repoName: string;
  entityName: string;
  orgName: string;

  // approval metadata augmented
  approvalType: string | IReportActionLink;

  approvalJustification?: any;
  releaseReviewUrl?: string;
  approvalTypeId?: any;
  approvalLicense?: string;

  countOfAdministratorCollaborators: number | string;
  countOfAdministratorTeams: number | string;

  createdBy?: string;
  createdByUpn?: string;
  createdByLink?: string | IReportActionLink;

  pushed?: any;
  created?: any;
  updated?: any;
  recentActivity?: any;
  abandoned?: any;
  exemptionExpiresAt?: any;

  status?: string | any;
  ageInMonths?: string;

  administrators?: any;
  recipients?: any;
  additionalRecipients?: any;
}

interface ICampaignData {
  utm_source: string;
  utm_medium: string;
  utm_campaign: string;
  utm_content: string;

  go_github_prefix?: string;
  go_github?: string;
  go_github_query?: string;
}

interface IReportsReducedEntity {
  name: string;
  issues?: any[];
  recipients?: any[];
}

export function process(context: IReportsContext): Promise<IReportsContext> {
  return getRepos(context).then(iterateRepos);
}

async function getRepositoryAdministrators(
  repositoryContext: IReportsRepositoryContext
): Promise<Map<number, IUserEntry>> {
  const repository = repositoryContext.repository;
  const administrators = new Map<number, IUserEntry>();
  const cacheOptions = {
    backgroundRefresh: false, // immediate
    maxAgeSeconds: 60 * 60 * 24 * 3, // 3 days
  };
  const teams = await repository.getTeamPermissions(cacheOptions);
  const adminTeams = teams.filter((team) => team.permission === 'admin');
  repositoryContext.countOfAdministratorTeams = adminTeams.length;
  await teamMembers(cacheOptions, administrators, adminTeams);
  const directCollaborators = await getRepositoryDirectCollaborators(
    repository
  );
  const directAdminCollaborators = await justAdminCollaborators(
    directCollaborators
  );
  repositoryContext.countOfAdministratorCollaborators =
    directAdminCollaborators.length;
  await storeCollaborators(administrators, directAdminCollaborators);
  return administrators;
}

async function teamMembers(
  cacheOptions,
  administrators: Map<number, IUserEntry>,
  teamPermissions: TeamPermission[]
) {
  for (const teamPermission of teamPermissions) {
    const team = teamPermission.team;
    const members = await team.getMembers(cacheOptions);
    for (let i = 0; i < members.length; i++) {
      const member = members[i];
      const id = member.id;
      let entry = administrators.get(id);
      if (!entry) {
        entry = createUserEntry(member);
        administrators.set(id, entry);
      }
      entry.reasons.memberships.push(team);
    }
  }
}

function identifyActionableAdmins(
  repositoryContext: IReportsRepositoryContext,
  repository: Repository,
  administrators
) {
  const automaticTeams = new AutomaticTeamsWebhookProcessor();
  const { specialTeamIds } = automaticTeams.processOrgSpecialTeams(
    repository.organization
  );
  const actionableAdministrators = [];
  const adminIds = Array.from(administrators.keys());
  for (let i = 0; i < adminIds.length; i++) {
    const id = adminIds[i];
    const account = administrators.get(id);
    // Remove service accounts
    if (account.link && account.link.serviceAccount) {
      continue;
    }
    // Direct collaborators should always be involved
    if (account.reasons.directCollaborator) {
      actionableAdministrators.push(account);
      continue;
    }
    // Remove administrators who are only in special teams
    let realMemberships = 0;
    for (let j = 0; j < account.reasons.memberships.length; j++) {
      const team = account.reasons.memberships[j].id;
      if (!specialTeamIds.has(team)) {
        ++realMemberships;
      }
    }
    if (realMemberships) {
      actionableAdministrators.push(account);
    }
  }
  repositoryContext.actionableAdministrators = actionableAdministrators;
  return administrators;
}

async function iterateRepos(
  context: IReportsContext
): Promise<IReportsContext> {
  let repos = context.entities.repos;
  if (context.settings.slice) {
    let offset = 3000;
    let initial = repos.length > offset ? offset - context.settings.slice : 0;
    repos = repos.slice(initial, initial + context.settings.slice);
  }
  context.processing.repos = {
    remaining: repos.length,
  };
  for (const repo of repos) {
    try {
      await processRepository(context, repo);
    } catch (processRepositoryError) {
      console.dir(processRepositoryError);
    }
  }
  // Settled values are not reviewed, since many are just missing repos (404)
  // for repos that have already been deleted or otherwise moved
  context.repositoryData = _.sortBy(context.repositoryData, 'nameLowercase');
  return context;
}

async function getRepositoryDetails(repositoryContext) {
  const repository = repositoryContext.repository as Repository;
  const cacheOptions = {
    backgroundRefresh: false,
    maxAgeSeconds: 60 * 60 * 24, // 1 day
  };
  await repository.getDetails(cacheOptions);
  return repositoryContext;
}

function getIndividualUserLink(context: IReportsContext, id: number) {
  if (!context.linkData) {
    return Promise.reject(new Error('No link information has been loaded'));
  }
  return Promise.resolve(context.linkData.get(id));
}

async function gatherLinkData(
  repositoryContext: IReportsRepositoryContext,
  administrators: Map<number, any>
) {
  const keys = Array.from(administrators.keys());
  for (const id of keys) {
    const link = await getIndividualUserLink(repositoryContext.parent, id);
    const entry = administrators.get(id);
    entry.link = link;
  }
  return administrators;
}

async function processRepository(
  context: IReportsContext,
  repository: Repository
) {
  console.log(
    'repository: ' +
      context.processing.repos.remaining-- +
      ': ' +
      repository.full_name
  );
  const organization = repository.organization;
  // repo context
  const repositoryContext: IReportsRepositoryContext = {
    parent: context,
    definitionsUsed: new Set(),
    issues: {},
    name: repository.full_name,
    nameLowercase: repository.full_name.toLowerCase(),
    repository,
    countOfAdministratorCollaborators: 0,
    countOfAdministratorTeams: 0,
  };
  const campaignSettings = context.settings.campaign;
  function reposDirectLink(
    content,
    suffix?,
    alternateForRepoFullPath?
  ): string {
    const reposUrl = context.config.urls.repos;
    const q = getCampaignData(content);
    let fullPath = `${organization.name}/repos/${repository.name}`;
    if (suffix) {
      fullPath + '/' + suffix;
    }
    return (
      reposUrl +
      (alternateForRepoFullPath || fullPath) +
      '?' +
      querystring.stringify(q)
    );
  }
  function getCampaignData(content): ICampaignData {
    return {
      utm_source: campaignSettings.source,
      utm_medium: campaignSettings.medium,
      utm_campaign: campaignSettings.campaign,
      utm_content: content,
    };
  }
  function githubDirectLink(
    content,
    prefix?,
    suffix?,
    query?,
    alternateForRepoFullName?
  ) {
    const reposUrl = context.config.urls.repos;
    const repoFullName = repositoryContext.name; // full_name
    const q = getCampaignData(content);
    q.go_github = null;
    if (prefix) {
      q.go_github_prefix = prefix;
    }
    if (suffix) {
      q.go_github = suffix;
    }
    if (query) {
      q.go_github_query = query;
    }
    return (
      reposUrl +
      (alternateForRepoFullName || repoFullName) +
      '?' +
      querystring.stringify(q)
    );
  }
  if (!context.repositoryData) {
    context.repositoryData = [];
  }
  context.repositoryData.push(repositoryContext);
  await getRepositoryDetails(repositoryContext);
  const administrators = await getRepositoryAdministrators(repositoryContext);
  repositoryContext.administrators = administrators;
  await gatherLinkData(repositoryContext, administrators);
  await identifyActionableAdmins(repositoryContext, repository, administrators);
  await identityAdministratorsWithoutLinks(repositoryContext);
  const privateEngineering = organization.privateEngineering;
  const basicRepository: IBasicRepository = {
    repoName: repository.name,
    entityName: repository.full_name,
    orgName: organization.name,
    // Pre-populate; overwritten if and when an approval is found
    approvalType: {
      color: 'gray',
      text: 'Created on GitHub or unknown',
    },
    countOfAdministratorCollaborators:
      repositoryContext.countOfAdministratorCollaborators || '-',
    countOfAdministratorTeams:
      repositoryContext.countOfAdministratorTeams || '-',
  };
  await getNewRepoCreationInformation(
    context,
    repositoryContext,
    basicRepository
  );
  const publicPrivateStatus = {
    text: repository.private ? 'Private' : 'Public',
    color: repository.private ? 'red' : 'green',
  };
  basicRepository.status = publicPrivateStatus;
  // Recipients
  repositoryContext.recipients = [];
  const corporateAdministrators = [];
  if (repositoryContext.actionableAdministrators) {
    for (
      let y = 0;
      y < repositoryContext.actionableAdministrators.length;
      y++
    ) {
      const admin = repositoryContext.actionableAdministrators[y];
      if (admin && admin.link && admin.link.aadupn) {
        corporateAdministrators.push(admin.link.aadupn);
        if (!privateEngineering) {
          // Private engineering orgs do not send individuals nags on emails for now
          repositoryContext.recipients.push({
            type: 'upn',
            value: admin.link.aadupn,
            reasons: transformReasonsToArray(admin, repository.full_name),
          });
        }
      }
    }
  }
  // Send to org admins
  const orgName = repository.organization.name;
  const orgData = context.organizationData[orgName];
  for (
    let i = 0;
    orgData &&
    orgData.organizationContext &&
    orgData.organizationContext.recipients &&
    orgData.organizationContext.recipients.length &&
    i < orgData.organizationContext.recipients.length;
    i++
  ) {
    repositoryContext.recipients.push(
      orgData.organizationContext.recipients[i]
    );
  }
  // Basic administrators info
  basicRepository.administrators = 'None';
  if (corporateAdministrators.length > 0) {
    let caLink =
      'mailto:' +
      corporateAdministrators.join(';') +
      '?subject=' +
      repository.full_name;
    const peoplePlurality =
      corporateAdministrators.length > 1 ? 'people' : 'person';
    basicRepository.administrators = {
      link: caLink,
      text: `${corporateAdministrators.length} ${peoplePlurality}`,
    };
  }
  const actionEditCollaborators = {
    link: githubDirectLink(
      'editRepoPermissions',
      null,
      'settings/collaboration'
    ),
    text: 'Permissions',
  };
  const actionDelete = {
    link: githubDirectLink('repoDeleteOrTransfer', null, 'settings'),
    text: 'Consider deleting or transferring',
  };
  const actionView = {
    link: githubDirectLink('repoBrowse'),
    text: 'Open',
  };
  const actionShip = {
    link: githubDirectLink('repoShipIt', null, 'settings'),
    text: 'Ship it',
  };
  const actionViewInPortal = context.config.urls
    ? {
        link: reposDirectLink('repoDetails'),
        text: 'Details',
      }
    : null;
  if (
    repositoryContext.administratorsByType.linked.length === 0 ||
    repositoryContext.actionableAdministrators.length === 0
  ) {
    addEntityToIssueType(
      context,
      repositoryContext,
      'noRepositoryAdministrators',
      basicRepository,
      actionEditCollaborators,
      actionViewInPortal
    );
  }
  let createdAt = repository.created_at ? moment(repository.created_at) : null;
  if (createdAt) {
    basicRepository.created = createdAt.format(simpleDateFormat);
  }
  let updatedAt = repository.updated_at ? moment(repository.updated_at) : null;
  if (updatedAt) {
    basicRepository.updated = updatedAt.format(simpleDateFormat);
  }
  let pushedAt = repository.pushed_at ? moment(repository.pushed_at) : null;
  if (pushedAt) {
    basicRepository.pushed = pushedAt.format(simpleDateFormat);
  }
  let mostRecentActivityMoment = createdAt;
  let mostRecentActivity = 'Created';
  if (updatedAt && updatedAt.isAfter(mostRecentActivityMoment)) {
    mostRecentActivity = 'Updated';
    mostRecentActivityMoment = updatedAt;
  }
  if (pushedAt && pushedAt.isAfter(mostRecentActivityMoment)) {
    mostRecentActivity = 'Pushed';
    mostRecentActivityMoment = pushedAt;
  }
  const twoYearsAgo = moment().subtract(2, 'years');
  const oneYearAgo = moment().subtract(1, 'years');
  const nineMonthsAgo = moment().subtract(9, 'months');
  const thirtyDaysAgo = moment().subtract(30, 'days');
  const thisWeek = moment().subtract(7, 'days');
  const today = moment().subtract(1, 'days');
  const ageInMonths = today.diff(createdAt, 'months');
  if (ageInMonths > 0) {
    basicRepository.ageInMonths =
      ageInMonths === 1 ? '1 month' : ageInMonths + ' months';
  }
  const monthsSinceUpdates = today.diff(mostRecentActivityMoment, 'months');
  const timeAsString =
    monthsSinceUpdates + ' month' + (monthsSinceUpdates === 1 ? '' : 's');
  basicRepository.recentActivity =
    monthsSinceUpdates < 1
      ? 'Active'
      : `${timeAsString} (${mostRecentActivity})`;
  if (mostRecentActivityMoment.isBefore(nineMonthsAgo)) {
    basicRepository.abandoned = {
      text: `${monthsSinceUpdates} months`,
      color: 'red',
    };
  }
  if (
    exemptRepositories &&
    exemptRepositories[repository.id] &&
    exemptRepositories[repository.id].approved &&
    exemptRepositories[repository.id].days
  ) {
    const exemptionExpiresAt = moment(
      exemptRepositories[repository.id].approved
    )
      .add(exemptRepositories[repository.id].days, 'days')
      .subtract(2, 'weeks');
    if (moment().isAfter(exemptionExpiresAt)) {
      basicRepository.exemptionExpiresAt = exemptionExpiresAt.format(
        simpleDateFormat
      );
      addEntityToIssueType(
        context,
        repositoryContext,
        'expiringPrivateEngineeringExemptions',
        basicRepository,
        actionShip,
        actionDelete
      );
    }
  } else if (
    !repository.private &&
    mostRecentActivityMoment.isBefore(twoYearsAgo)
  ) {
    addEntityToIssueType(
      context,
      repositoryContext,
      'abandonedPublicRepositories',
      basicRepository,
      actionView,
      actionDelete
    );
  } else if (
    repository.private &&
    mostRecentActivityMoment.isBefore(twoYearsAgo)
  ) {
    addEntityToIssueType(
      context,
      repositoryContext,
      'twoYearOldPrivateRepositories',
      basicRepository,
      actionView,
      actionDelete
    );
  } else if (
    repository.private &&
    createdAt.isBefore(oneYearAgo) &&
    !privateEngineering
  ) {
    addEntityToIssueType(
      context,
      repositoryContext,
      'oneYearOldPrivateRepositories',
      basicRepository,
      actionView,
      actionDelete
    );
  } else if (
    repository.private &&
    createdAt.isBefore(thirtyDaysAgo) &&
    !privateEngineering
  ) {
    addEntityToIssueType(
      context,
      repositoryContext,
      'privateRepositoriesLessThanOneYear',
      basicRepository,
      actionShip,
      actionDelete
    );
  } else if (createdAt.isAfter(thisWeek) && !privateEngineering) {
    // New public and private repos
    const repositoryForManagerAndLawyer = shallowCloneWithAdditionalRecipients(
      basicRepository,
      repositoryContext.additionalRecipients
    );
    if (createdAt.isAfter(today)) {
      addEntityToIssueType(
        context,
        repositoryContext,
        'NewReposToday',
        repositoryForManagerAndLawyer,
        actionView,
        actionViewInPortal
      );
    }
    // Always include in the weekly summary
    addEntityToIssueType(
      context,
      repositoryContext,
      'NewReposWeek',
      repositoryForManagerAndLawyer,
      actionView,
      actionViewInPortal
    );
  }
  // Alert on too many administrators, excluding private engineering organizations at this time
  // NOTE: commenting out the "too many" notice for September 2017
  //if (!privateEngineering && repositoryContext.actionableAdministrators.length > context.settings.tooManyRepoAdministrators) {
  //addEntityToIssueType(context, repositoryContext, 'repositoryTooManyAdministrators', basicRepository, actionViewInPortal, actionEditCollaborators);
  //}
  if (context.settings.repoDelayAfter) {
    await sleep(context.settings.repoDelayAfter);
  }
}

function shallowCloneWithAdditionalRecipients(
  basicRepository: IBasicRepository,
  additionalRecipients
) {
  const clone = Object.assign({}, basicRepository);
  if (additionalRecipients && additionalRecipients.length) {
    clone.additionalRecipients = additionalRecipients;
  }
  return clone;
}

async function getNewRepoCreationInformation(
  context: IReportsContext,
  repositoryContext: IReportsRepositoryContext,
  basicRepository: IBasicRepository
): Promise<void> {
  const repository = repositoryContext.repository;
  const thisWeek = moment().subtract(7, 'days');
  let createdAt = repository.created_at ? moment(repository.created_at) : null;
  let isBrandNew = createdAt.isAfter(thisWeek);
  const repositoryMetadataProvider =
    context.providers.repositoryMetadataProvider;
  if (!isBrandNew || !repositoryMetadataProvider) {
    return;
  }
  const releaseTypeMapping =
    context.config &&
    context.config.github &&
    context.config.github.approvalTypes &&
    context.config.github.approvalTypes.fields
      ? context.config.github.approvalTypes.fields.approvalIdsToReleaseType
      : null;
  let approval = null;
  try {
    approval = await repositoryMetadataProvider.getRepositoryMetadata(
      repository.id
    );
  } catch (approvalGetError) {
    return;
  }
  if (!approval) {
    return;
  }
  if (
    approval.repositoryId ==
      repositoryContext.repository
        .id /* not strict equal, data client IDs are strings vs GitHub responses use numbers */ ||
    (approval.organizationName &&
      approval.organizationName.toLowerCase() ===
        repositoryContext.repository.organization.name.toLowerCase())
  ) {
    basicRepository.approvalLicense = approval.initialLicense;
    basicRepository.approvalJustification = approval.releaseReviewJustification;
    if (approval.releaseReviewType && releaseTypeMapping) {
      const approvalTypes = Object.getOwnPropertyNames(releaseTypeMapping);
      for (let j = 0; j < approvalTypes.length; j++) {
        const id = approvalTypes[j];
        const title = releaseTypeMapping[id];
        if (approval.projectType === id) {
          basicRepository.approvalTypeId = approval.projectType; // ?
          // Hard-coded specific to show justification text or approval links
          if (
            (id === 'NewReleaseReview' || id === 'ExistingReleaseReview') &&
            approval.releaseReviewUrl
          ) {
            basicRepository.approvalType = {
              text: title,
              link: approval.releaseReviewUrl,
            };
          } else if (id !== 'Exempt') {
            basicRepository.approvalType = title;
          } else {
            basicRepository.approvalType = `${title}: ${approval.releaseReviewJustification}`;
          }
        }
      }
    }
    if (!basicRepository.approvalType) {
      basicRepository.approvalType = approval.projectType; // Fallback if it's not configured in the system
    }
    const createdBy = approval.createdByThirdPartyUsername;
    if (!createdBy) {
      return;
    } else {
      basicRepository.createdBy = createdBy;
      const id = await getIdFromUsername(
        context,
        repositoryContext.repository.organization,
        createdBy
      );
      const link = await getIndividualUserLink(context, id);
      basicRepository.createdBy =
        link.corporateDisplayName || basicRepository.createdBy;
      basicRepository.createdByUpn = link.corporateUsername;
      basicRepository.createdByLink = basicRepository.createdByUpn
        ? {
            link: `mailto:${basicRepository.createdByUpn}`,
            text: basicRepository.createdBy,
          }
        : basicRepository.createdBy;
      if (link.corporateUsername) {
        await augmentWithAdditionalRecipients(context, repositoryContext, link);
      }
    }
  }
}

async function augmentWithAdditionalRecipients(
  context: IReportsContext,
  repositoryContext,
  createdByLink: ICorporateLink
): Promise<IReportsContext> {
  if (!createdByLink || !createdByLink.corporateUsername) {
    return context;
  }
  if (createdByLink.isServiceAccount) {
    // Service accounts do not have legal contacts
    return context;
  }
  const upn = createdByLink.corporateUsername;
  const createdByName = createdByLink.corporateDisplayName || upn;
  const operations = context.operations;
  const {
    corporateContactProvider,
    mailAddressProvider,
  } = operations.providers;
  // Only if the provider supports both advanced Microsoft-specific functions for now
  if (!corporateContactProvider) {
    return context;
  }
  const fullRepoName = repositoryContext.repository.full_name;
  let additional = [];
  try {
    const contacts = await corporateContactProvider.lookupContacts(upn);
    if (contacts && contacts.managerUsername) {
      const managerName =
        contacts.managerDisplayName || contacts.managerUsername;
      additional.push({
        type: 'upn',
        value: contacts.managerUsername,
        reasons: [
          `${managerName} is the manager of ${createdByName} who created a new repository ${fullRepoName}`,
        ],
      });
    }
    if (contacts && contacts.openSourceContact) {
      let lc = contacts.openSourceContact;
      let isVstsTeam = false;
      // TODO: validate if this is ever even a case in the new reviewer model
      if (lc && lc.startsWith(projectPlaceholder)) {
        isVstsTeam = true;
        lc = lc.replace(projectPlaceholder, '[Reviews]\\');
      }
      let legalReason = `${lc} is the legal contact assigned to ${createdByName} who created a new repository ${fullRepoName}`;
      additional.push({
        type: isVstsTeam ? 'vststeam' : 'upn',
        value: lc,
        reasons: [legalReason],
      });
    }
  } catch (managerInformationError) {
    console.dir(managerInformationError);
  }
  if (additional.length) {
    repositoryContext.additionalRecipients = additional;
  }
  return context;
}

async function getIdFromUsername(
  context,
  organization: Organization,
  username: string
): Promise<number> {
  // Depends on this being a current member of an org
  const operations = context.operations as Operations;
  const account = await operations.getAccountByUsername(username);
  return account.id;
}

function addEntityToIssueType(
  context,
  repositoryContext,
  type,
  entity,
  optionalAction1,
  optionalAction2
) {
  const definition = definitionsByName[type];
  if (!definition) {
    throw new Error(`No defined issue type ${type}`);
  }
  let hadActions = true && entity.actions;
  const entityClone = Object.assign({}, entity);
  if (hadActions) {
    delete entityClone.actions;
  }
  if (!entityClone.actions && optionalAction1) {
    entityClone.actions = { actions: [] };
  }
  if (optionalAction1) {
    entityClone.actions.actions.push(optionalAction1);
  }
  if (optionalAction2) {
    entityClone.actions.actions.push(optionalAction2);
  }
  // Track that the definition was used provider-wide and per-entity
  repositoryContext.definitionsUsed.add(type);
  if (!context.visitedDefinitions[providerName]) {
    context.visitedDefinitions[providerName] = new Set();
  }
  context.visitedDefinitions[providerName].add(type);
  const placeholder = repositoryContext.issues;
  let propertyName = null;
  if (!placeholder[type]) {
    const entry: IReportEntry = {};
    placeholder[type] = entry;
    if (definition.hasTable && definition.hasList) {
      throw new Error(
        'Definitions cannot have both tables and lists at this time'
      );
    }
    if (definition.hasTable) {
      entry.rows = [];
    }
    if (definition.hasList) {
      entry.listItems = [];
    }
  }
  if (definition.hasTable && definition.hasList) {
    throw new Error(
      'Definitions cannot have both tables and lists at this time'
    );
  }
  let listPropertiesName = null;
  if (definition.hasTable) {
    propertyName = 'rows';
    listPropertiesName = 'table';
  }
  if (definition.hasList) {
    propertyName = 'listItems';
    listPropertiesName = 'list';
  }
  if (!propertyName) {
    throw new Error('No definition items collection available');
  }
  const dest = placeholder[type][propertyName];
  dest.push(entityClone);
  const listProperties = definition[listPropertiesName];
  if (listProperties && (listProperties.groupBy || listProperties.sortBy)) {
    const sortBy = [dest];
    if (listProperties.groupBy) {
      sortBy.push(listProperties.groupBy);
    }
    if (listProperties.sortBy) {
      sortBy.push(listProperties.sortBy);
    }
    const after = _.sortBy.apply(null, sortBy);
    placeholder[type][propertyName] = after;
  }
}

async function identityAdministratorsWithoutLinks(
  repositoryContext: IReportsRepositoryContext
) {
  const actionableAdministrators = repositoryContext.actionableAdministrators;
  const administratorsByType = {
    linked: actionableAdministrators.filter((admin) => {
      return admin.link;
    }),
    unlinked: actionableAdministrators.filter((admin) => {
      return !admin.link;
    }),
  };
  repositoryContext.administratorsByType = administratorsByType;
  return repositoryContext;
}

function justAdminCollaborators(collaborators: Collaborator[]): Collaborator[] {
  return collaborators.filter((collaborator) => collaborator.permissions.admin);
}

function getRepositoryDirectCollaborators(repository: Repository) {
  const directCollaboratorOptions = {
    affiliation: GitHubCollaboratorAffiliationQuery.Direct,
    backgroundRefresh: true,
    maxAgeSeconds: 60 * 60 * 24, // 1 day allowed
  };
  return repository.getCollaborators(directCollaboratorOptions);
}

async function getRepos(context): Promise<any> {
  const operations = context.operations as Operations;
  const repos = await operations.getRepos();
  context.entities.repos = repos.sort((a, b) => {
    return a.full_name.localeCompare(b.full_name, 'en', {
      sensitivity: 'base',
    });
  });
  return context;
}

function createUserEntry(basics): IUserEntry {
  return {
    login: basics.login,
    reasons: {
      memberships: [],
      directCollaborator: false,
      collaborator: false,
    },
  };
}

function transformReasonsToArray(userEntry, repositoryName) {
  const reasons = [];
  // For efficiency reasons, direct collaborator wins over team memberships
  if (userEntry.reasons.directCollaborator) {
    reasons.push(`Administrator of the ${repositoryName} repository`);
  } else {
    for (let i = 0; i < userEntry.reasons.memberships.length; i++) {
      const team = userEntry.reasons.memberships[i];
      reasons.push(
        `Member of the ${team.name} team with administrator rights to the ${repositoryName} repository`
      );
    }
  }

  if (!reasons.length) {
    reasons.push(`Unknown reason related to the ${repositoryName}`);
  }
  return reasons;
}

async function storeCollaborators(
  administrators: Map<number, IUserEntry>,
  collaborators: Collaborator[]
) {
  for (const collaborator of collaborators) {
    const id = collaborator.id;
    let entry = administrators.get(id);
    if (!entry) {
      entry = createUserEntry(collaborator);
      administrators.set(id, entry);
    }
    entry.reasons.collaborator = true;
  }
}

export async function build(context: IReportsContext) {
  return context;
}

export async function consolidate(context: IReportsContext) {
  // For any used definitions of a provider entity instance, add it to the generic report
  const consolidated = {
    definitions: [],
    entities: [],
  };
  for (let i = 0; i < definitions.length; i++) {
    const definition = definitions[i];
    if (
      !context.visitedDefinitions ||
      !context.visitedDefinitions[providerName]
    ) {
      return context;
    }
    if (context.visitedDefinitions[providerName].has(definition.name)) {
      consolidated.definitions.push(definition);
    }
  }
  // Entities
  const sorted = _.sortBy(context.repositoryData, 'nameLowercase'); // 'entityName'); // full_name groups by org name AND repo name naturally
  for (let i = 0; i < sorted.length; i++) {
    const fullEntity = sorted[i];
    const reducedEntity: IReportsReducedEntity = {
      name: fullEntity.name,
    };
    const contextDirectProperties = ['issues', 'recipients'];
    cloneProperties(fullEntity, contextDirectProperties, reducedEntity);
    // Only store in the consolidated report if there are recipients for the entity
    const issueCounter = Object.getOwnPropertyNames(reducedEntity.issues);
    if (
      issueCounter &&
      issueCounter.length &&
      reducedEntity &&
      reducedEntity.recipients &&
      reducedEntity.recipients.length > 0
    ) {
      consolidated.entities.push(reducedEntity);
    } else if (issueCounter && issueCounter.length) {
      console.warn(
        `There are no recipients to receive ${reducedEntity.name} reports with active issues`
      );
    }
  }
  context.consolidated[providerName] = consolidated;
  return context;
}

function cloneProperties(source, properties, target) {
  for (let j = 0; j < properties.length; j++) {
    const property = properties[j];
    if (source[property]) {
      target[property] = source[property];
    }
  }
}
