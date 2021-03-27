//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/* eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import _ from 'lodash';
import moment from 'moment';

import { requireJson, sleep } from '../../utils';
import { IReportsContext } from './task';
import { Team } from '../../business/team';
import { TeamMember } from '../../business/teamMember';
import { ErrorHelper } from '../../transitional';

interface IReportsTeamsReducedEntity {
  name: string;
  issues?: any;
  recipients?: any[];
}

interface IMapWithOrganization extends Array<any> {
  orgs?: any;
}

export interface IReportsTeamContext {
  parent: any;

  definitionsUsed: Set<any>;
  issues: any;

  name: string;
  nameLowercase: string;

  team: Team;

  recipients?: any[];
  additionalRecipients?: any[];
  maintainers?: any[];
  teamMaintainers?: any[];
  maintainersByType?: {
    linked: any[];
    unlinked: any[];
  };
}

interface IReportsBasicTeam {
  teamName: string;
  teamSlug: string;
  entityName: string;
  orgName: string;
  id: string;

  maintainers?: any;

  created?: any;
  updated?: any;

  ageInMonths?: any;
  recentActivity?: any;
}

interface IReportsBasicTeamsEntry {
  rows?: any[];
  listItems?: any[];
}

const providerName = 'teams';
const definitions = requireJson('jobs/reports/teamDefinitions.json');
const definitionsByName = {};
for (let i = 0; i < definitions.length; i++) {
  const definition = definitions[i];
  definitionsByName[definition.name] = definition;
}

const simpleDateFormat = 'l';

export async function process(
  context: IReportsContext
): Promise<IReportsContext> {
  await getTeams(context);
  await iterateTeams(context);
  return context;
}

async function identityMaintainersWithoutLinks(
  teamContext: IReportsTeamContext,
  maintainers
): Promise<IReportsTeamContext> {
  const maintainersByType = {
    linked: maintainers.filter((admin) => {
      return admin.link;
    }),
    unlinked: maintainers.filter((admin) => {
      return !admin.link;
    }),
  };
  teamContext.maintainersByType = maintainersByType;
  return teamContext;
}

async function identifyTeamMaintainers(
  teamContext: IReportsTeamContext,
  team: Team
): Promise<TeamMember[]> {
  const cacheOptions = {
    backgroundRefresh: false, // immediate
    maxAgeSeconds: 60 * 60 * 5, // 5 hours
  };
  const maintainers = await team.getMaintainers(cacheOptions);
  teamContext.teamMaintainers = maintainers;
  return maintainers;
}

async function iterateTeams(context: IReportsContext) {
  let teams = context.entities.teams;
  context.processing.teams = {
    remaining: teams.length,
  };
  for (const team of teams) {
    try {
      await processTeam(context, team);
    } catch (processTeamError) {
      // Settled values are not reviewed, since many are just missing teams (404)
      // for teams that have already been deleted or otherwise moved
      console.dir(processTeamError);
    }
  }
  context.teamData = _.sortBy(context.teamData, 'nameLowercase');
  return context;
}

async function getTeamDetails(
  teamContext: IReportsTeamContext
): Promise<IReportsTeamContext> {
  const team = teamContext.team as Team;
  const cacheOptions = {
    backgroundRefresh: false,
    maxAgeSeconds: 60 * 60 * 24, // 1 day
  };
  await team.getDetails(cacheOptions);
  return teamContext;
}

async function gatherLinkData(
  teamContext: IReportsTeamContext,
  maintainers: TeamMember[]
) {
  for (const maintainer of maintainers) {
    try {
      await maintainer.resolveDirectLink();
    } catch (ignoreDirectResolveError) {
      console.dir(ignoreDirectResolveError);
    }
  }
  return maintainers;
}

async function processTeam(context: IReportsContext, team: Team) {
  console.log(
    'team: ' +
      context.processing.teams.remaining-- +
      ': ' +
      team.organization.name +
      '/' +
      team.name
  );
  try {
    const teamContext: IReportsTeamContext = {
      parent: context,
      definitionsUsed: new Set(),
      issues: {},
      name: team.name,
      nameLowercase: team.name.toLowerCase(),
      team: team,
    };
    if (!context.teamData) {
      context.teamData = [];
    }
    context.teamData.push(teamContext);
    await getTeamDetails(teamContext);
    const maintainers = await identifyTeamMaintainers(teamContext, team);
    await gatherLinkData(teamContext, maintainers);
    await identityMaintainersWithoutLinks(teamContext, maintainers);
    const organization = team.organization;
    // We do not provide reports for private engineering orgs for now
    const privateEngineering = organization.privateEngineering;
    // Some organizations, such as the .NET Foundation, will allow external community members to
    // be organization members
    const externalMembersPermitted = organization.externalMembersPermitted;
    const slug = team.slug;
    const basicTeam: IReportsBasicTeam = {
      teamName: team.name,
      teamSlug: slug,
      entityName: team.name,
      orgName: organization.name,
      id: team.id.toString(),
    };
    const orgName = team.organization.name;
    // Recipients
    teamContext.recipients = [];
    const corporateAdministrators = [];
    if (teamContext.maintainers) {
      for (let y = 0; y < teamContext.maintainers.length; y++) {
        const admin = teamContext.maintainers[y];
        if (admin && admin.link && admin.link.aadupn) {
          corporateAdministrators.push(admin.link.aadupn);
          if (!privateEngineering) {
            // Private engineering orgs do not send individuals nags on emails for now
            teamContext.recipients.push({
              type: 'upn',
              value: admin.link.aadupn,
              reasons: transformReasonsToArray(admin, team.name, orgName),
            });
          }
        }
      }
    }
    // Send to org admins
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
      teamContext.recipients.push(orgData.organizationContext.recipients[i]);
    }
    basicTeam.maintainers =
      teamContext.maintainers && teamContext.maintainers.length
        ? teamContext.maintainers.length.toString()
        : 'None';
    const actionPromoteMembers = {
      link: `https://github.com/orgs/${orgName}/teams/${slug}/members`,
      text: 'Promote members',
    };
    const actionRemoveMembers = {
      link: `https://github.com/orgs/${orgName}/teams/${slug}/members`,
      text: 'Remove members',
    };
    const actionDelete = {
      link: `https://github.com/orgs/${orgName}/teams/${slug}/edit`,
      text: 'Consider deleting',
    };
    const actionView = {
      link: `https://github.com/orgs/${orgName}/teams/${slug}/members`,
      text: 'Open',
    };
    const actionViewInPortal = context.config.urls
      ? {
          link: `${context.config.urls.repos}${organization.name}/teams/${slug}`,
          text: 'Manage team',
        }
      : null;
    let createdAt = team.created_at ? moment(team.created_at) : null;
    if (createdAt) {
      basicTeam.created = createdAt.format(simpleDateFormat);
    }
    let updatedAt = team.updated_at ? moment(team.updated_at) : null;
    if (updatedAt) {
      basicTeam.updated = updatedAt.format(simpleDateFormat);
    }
    let mostRecentActivityMoment = createdAt;
    let mostRecentActivity = 'Created';
    if (updatedAt && updatedAt.isAfter(mostRecentActivityMoment)) {
      mostRecentActivity = 'Updated';
      mostRecentActivityMoment = updatedAt;
    }
    // Completely empty standard teams (we exclude system teams that may be used for specific portal permissions)
    const systemTeamIds = team.organization.systemTeamIds;
    const isSystemTeam = systemTeamIds.includes(team.id);
    if (team.members_count == 0 && team.repos_count == 0 && !isSystemTeam) {
      addEntityToIssueType(
        context,
        teamContext,
        'emptyTeams',
        basicTeam,
        actionDelete
      );
    } else {
      // Member or maintainer issues
      if (!isSystemTeam && team.members_count == 0) {
        addEntityToIssueType(
          context,
          teamContext,
          'noTeamMembers',
          basicTeam,
          actionDelete,
          actionViewInPortal
        );
      } else if (!isSystemTeam && teamContext.teamMaintainers.length === 0) {
        addEntityToIssueType(
          context,
          teamContext,
          'noTeamMaintainers',
          basicTeam,
          actionPromoteMembers,
          actionViewInPortal
        );
      } else if (teamContext.maintainersByType.unlinked.length > 0) {
        const logins = [];
        for (
          let z = 0;
          z < teamContext.maintainersByType.unlinked.length;
          z++
        ) {
          const unlinkedEntry = teamContext.maintainersByType.unlinked[z];
          logins.push(unlinkedEntry.login);
        }
        const specialEntity = Object.assign(
          {
            logins: logins.join(', '),
          },
          basicTeam
        );
        const reportName = externalMembersPermitted
          ? 'unlinkedMaintainersWhenAllowed'
          : 'unlinkedMaintainers';
        // TODO: use the operations e-mail
        addEntityToIssueType(
          context,
          teamContext,
          reportName,
          specialEntity,
          actionRemoveMembers,
          {
            link: `mailto:opensource@microsoft.com?subject=Reporting a former employee related to the ${orgName} ${team.name} team`,
            text: 'Former employee?',
          }
        );
      }
      // No repositories
      if (team.repos_count <= 0 && !isSystemTeam) {
        addEntityToIssueType(
          context,
          teamContext,
          'TeamsWithoutRepositories',
          basicTeam,
          actionViewInPortal
        );
      }
    }
    const thisWeek = moment().subtract(7, 'days');
    const today = moment().subtract(1, 'days');
    const ageInMonths = today.diff(createdAt, 'months');
    if (ageInMonths > 0) {
      basicTeam.ageInMonths =
        ageInMonths === 1 ? '1 month' : ageInMonths + ' months';
    }
    const monthsSinceUpdates = today.diff(mostRecentActivityMoment, 'months');
    const timeAsString =
      monthsSinceUpdates + ' month' + (monthsSinceUpdates === 1 ? '' : 's');
    basicTeam.recentActivity =
      monthsSinceUpdates < 1
        ? 'Active'
        : `${timeAsString} (${mostRecentActivity})`;
    if (createdAt.isAfter(thisWeek) && !privateEngineering) {
      // New public and private repos
      const teamForManagerAndLawyer = shallowCloneWithAdditionalRecipients(
        basicTeam,
        teamContext.additionalRecipients
      );
      if (createdAt.isAfter(today)) {
        addEntityToIssueType(
          context,
          teamContext,
          'NewTeamsToday',
          teamForManagerAndLawyer,
          actionView,
          actionViewInPortal
        );
      }
      // Always include in the weekly summary
      addEntityToIssueType(
        context,
        teamContext,
        'NewTeamsWeek',
        teamForManagerAndLawyer,
        actionView,
        actionViewInPortal
      );
    }
    if (context.settings.teamDelayAfter) {
      await sleep(context.settings.teamDelayAfter);
    }
    return context;
  } catch (problem) {
    if (ErrorHelper.IsNotFound(problem)) {
      // Missing teams are OK to not spew too many errors about...
    } else {
      console.dir(problem);
    }
    throw problem;
  }
}

function shallowCloneWithAdditionalRecipients(basicTeam, additionalRecipients) {
  const clone = Object.assign({}, basicTeam);
  if (additionalRecipients && additionalRecipients.length) {
    clone.additionalRecipients = additionalRecipients;
  }
  return clone;
}

function addEntityToIssueType(
  context,
  teamContext,
  type,
  entity,
  optionalAction1?,
  optionalAction2?
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
  teamContext.definitionsUsed.add(type);
  if (!context.visitedDefinitions[providerName]) {
    context.visitedDefinitions[providerName] = new Set();
  }
  context.visitedDefinitions[providerName].add(type);

  const placeholder = teamContext.issues;
  let propertyName = null;
  if (!placeholder[type]) {
    const entry: IReportsBasicTeamsEntry = {};
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

function flattenTeamsMap(teamsMap, operations) {
  const asArray = Array.from(teamsMap.values()) as IMapWithOrganization;
  const values = [];
  for (let i = 0; i < asArray.length; i++) {
    const byOrg = asArray[i].orgs;
    const organizationNames = Object.getOwnPropertyNames(byOrg);
    if (organizationNames.length !== 1) {
      throw new Error(
        'Expected just a single organization for the team while preparing the list of teams to report on'
      );
    }
    const orgName = organizationNames[0];
    const organization = operations.getOrganization(orgName);
    const entity = byOrg[orgName];
    const team = organization.teamFromEntity(entity);
    values.push(team);
  }
  return values;
}

async function getTeams(context: IReportsContext): Promise<IReportsContext> {
  const operations = context.operations;
  const teams = await operations.getCrossOrganizationTeams();
  const asArray = flattenTeamsMap(teams, operations);
  context.entities.teams = asArray.sort((a, b) => {
    const aFullName = `${a.organization.name}/${a.name}`;
    const bFullName = `${b.organization.name}/${b.name}`;
    return aFullName.localeCompare(bFullName, 'en', { sensitivity: 'base' });
  });
  return context;
}

function transformReasonsToArray(userEntry, teamName, orgName) {
  const reasons = [];
  // For efficiency reasons, direct collaborator wins over team memberships
  if (userEntry.reasons.directCollaborator) {
    reasons.push(`Administrator of the ${teamName} team`);
  } else {
    for (let i = 0; i < userEntry.reasons.memberships.length; i++) {
      const team = userEntry.reasons.memberships[i];
      reasons.push(
        `Maintainer of the ${team.name} team in the ${orgName} GitHub organization`
      );
    }
  }

  if (!reasons.length) {
    reasons.push(
      `Unknown reason related to the ${teamName} team in the ${orgName} GitHub organization`
    );
  }
  return reasons;
}

export async function build(context: IReportsContext) {
  return context;
}

export async function consolidate(
  context: IReportsContext
): Promise<IReportsContext> {
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
  const sorted = _.sortBy(context.teamData, 'nameLowercase'); // 'entityName'); // name groups by org name AND repo name naturally
  for (let i = 0; i < sorted.length; i++) {
    const fullEntity = sorted[i];
    const reducedEntity: IReportsTeamsReducedEntity = {
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
