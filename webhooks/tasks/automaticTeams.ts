//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

const teamTypes = ['read', 'write', 'admin'];
const defaultLargeAdminTeamSize = 250;

import { WebhookProcessor } from '../organizationProcessor';
import { Operations } from '../../business';
import { Organization } from '../../business';

import RenderHtmlMail from '../../lib/emailRender';
import { IMailProvider } from '../../lib/mailProvider';
import getCompanySpecificDeployment from '../../middleware/companySpecificDeployment';

interface IAutomaticTeamsMail {
  to: string;
  cc?: string;
  subject: string;
  category?: string[];
  content?: string;
}

interface ICustomDataEventName {
  content?: string;
  receipt: string;
  eventName?: string;
}

export default class AutomaticTeamsWebhookProcessor implements WebhookProcessor {
  processOrgSpecialTeams(organization: Organization) {
    const specialTeams = organization.specialRepositoryPermissionTeams;
    let specials = [];
    let specialTeamIds = new Set<number>();
    let specialTeamLevels = new Map<number, string>();
    teamTypes.forEach(specialTeam => {
      if (specialTeams[specialTeam] && specialTeams[specialTeam].length) {
        specials.push(specialTeam);
        for (let i = 0; i < specialTeams[specialTeam].length; i++) {
          const teamId = specialTeams[specialTeam][i];
          specialTeamIds.add(teamId);
          specialTeamLevels.set(teamId, translateSpecialToGitHub(specialTeam));
        }
      }
    });
    return { specialTeams, specials, specialTeamIds, specialTeamLevels };
  }

  filter(data: any) {
    const eventType = data.properties.event;
    const eventAction = data.body.action;
    // Someone added a team to the repo
    if (eventType === 'team' && ['add_repository', 'added_to_repository'].includes(eventAction)) {
      return true;
    }
    // Someone removed a team from the repo
    if (eventType === 'team' && eventAction === 'removed_from_repository') {
      return true;
    }
    // Team permission level changed
    if (eventType === 'team' && eventAction === 'edited') {
      return true;
    }
    // A new repo may need the teams
    if (eventType === 'repository' && eventAction === 'created') {
      return true;
    }
    return false;
  }

  async run(operations: Operations, organization: Organization, data: any): Promise<boolean> {
    const eventType = data.properties.event;
    const eventAction = data.body.action;
    const { specialTeamIds, specialTeamLevels } = this.processOrgSpecialTeams(organization);
    const preventLargeTeamPermissions = organization.preventLargeTeamPermissions;
    const repositoryBody = data.body.repository;
    const newPermissions = repositoryBody ? repositoryBody.permissions : null;
    const whoChangedIt = (data.body && data.body.sender ? data.body.sender.login : null) as string;
    const whoChangedItId = whoChangedIt ? data.body.sender.id : null;

    // New repository
    if (eventType === 'repository' && eventAction === 'created') {
      for (const teamId of specialTeamIds) {
        const necessaryPermission = specialTeamLevels.get(teamId);
        await setTeamPermission(operations, organization, repositoryBody, teamId, necessaryPermission, `a new repository was created by username ${whoChangedIt}, setting automatic permissions`);
      }
    } else if (eventType === 'team') {
      const teamBody = data.body.team;
      const teamId = teamBody.id;
      const teamName = teamBody.name;
      // Enforce required special team permissions
      if (specialTeamIds.has(teamId)) {
        const necessaryPermission = specialTeamLevels.get(teamId);
        if (!necessaryPermission) {
          throw new Error(`No ideal permission level found for the team ${teamId}.`);
        }
        if (eventAction === 'removed_from_repository') {
          // Someone removed the entire team
          await setTeamPermission(operations, organization, repositoryBody, teamId, necessaryPermission, `the team and its permission were removed by the username ${whoChangedIt}`);
        } else if (eventAction === 'edited') {
          // The team no longer has the appropriate permission level
          if (newPermissions[necessaryPermission] !== true) {
            await setTeamPermission(operations, organization, repositoryBody, teamId, necessaryPermission, `the permission was downgraded by the username ${whoChangedIt}`);
          }
        }
        return true;
      }

      // Prevent granting large teams access
      if (preventLargeTeamPermissions) {
        const teamSize = await getTeamSize(organization, teamId);
        // Special thanks to the GitHub API team. The added_to_repository event did not
        // include the 'permissions' information. Fixed and deployed by GitHub on
        // 6/13/17. Thank you for helping us simplify our code!
        if (['added_to_repository', 'edited'].includes(eventAction) && newPermissions) {
          const specificReason = teamTooLargeForPurpose(teamId, newPermissions.admin, newPermissions.push, organization, teamSize, preventLargeTeamPermissions);
          if (specificReason && !operations.isSystemAccountByUsername(whoChangedIt)) {
            await revertLargePermissionChange(operations, organization, repositoryBody, teamId, teamName, whoChangedIt, whoChangedItId, specificReason);
          }
        }
        return true;
      }
    }

    return true;
  }
}

function teamTooLargeForPurpose(teamId, isAdmin, isPush, organization, teamSize, preventLargeTeamPermissions) {
  const broadAccessTeams = organization.broadAccessTeams;
  let isBroadAccessTeam = broadAccessTeams && broadAccessTeams.includes(teamId);
  if (isBroadAccessTeam && (isAdmin || isPush)) {
    return 'The team is a very broad access team and does not allow push (write) or admin access to prevent widespread escalation of privileges and spamming thousands of people';
  }
  let teamSizeLimitAdmin = defaultLargeAdminTeamSize;
  let teamSizeLimitType = 'default limit';
  if (preventLargeTeamPermissions && preventLargeTeamPermissions.maximumAdministrators) {
    teamSizeLimitAdmin = preventLargeTeamPermissions.maximumAdministrators;
    teamSizeLimitType = `administrator team limit in the ${organization.name} organization`;
  }
  if (isAdmin && teamSize >= teamSizeLimitAdmin) {
    return `The team has ${teamSize} members which surpasses the ${teamSizeLimitAdmin} ${teamSizeLimitType}`;
  }

  return false;
}

function translateSpecialToGitHub(ourTerm) {
  switch (ourTerm) {
  case 'admin':
    return 'admin';
  case 'write':
    return 'push';
  case 'read':
    return 'pull';
  }
  throw new Error(`Unknown team type ${ourTerm}`);
}

export async function getTeamSize(organization: Organization, teamId): Promise<number> {
  const team = organization.team(teamId);
  await team.getDetails();
  return team.members_count || 0;
}

async function revertLargePermissionChange(operations: Operations, organization: Organization, repositoryBody, teamId, teamName: string, whoChangedIt, whoChangedItId: string, specificReason?: string) {
  specificReason = specificReason ? ': ' + specificReason : '';
  const blockReason = `the permission was upgraded by ${whoChangedIt} but a large team permission prevention feature has reverted the change${specificReason}`;
  console.log(blockReason);
  const insights = operations.insights;
  insights.trackMetric({ name: 'JobAutomaticTeamsLargeTeamPermissionBlock', value: 1 });
  insights.trackEvent({
    name: 'JobAutomaticTeamsLargeTeamPermissionBlocked',
    properties: {
      specificReason: specificReason,
      teamId: teamId,
      organization: organization.name,
      repository: repositoryBody.name,
      whoChangedIt: whoChangedIt,
      whoChangedItId: whoChangedItId,
    },
  });
  const successfulAndOk = await setTeamPermission(operations, organization, repositoryBody, teamId, 'pull', blockReason);
  if (successfulAndOk) {
    const owner = repositoryBody.owner.login.toLowerCase(); // We do not want to notify for each fork, if the permissions bubble to the fork
    if (owner === organization.name.toLowerCase()) {
      await largeTeamPermissionPreventionWarningMail(operations, organization, repositoryBody, teamId, teamName, blockReason, whoChangedIt, whoChangedItId);
    }
  }
}

async function largeTeamPermissionPreventionWarningMail(operations: Operations, organization: Organization, repositoryBody, teamId, teamName, reason, whoChangedIt, whoChangedItId): Promise<void> {
  // System accounts should not need notifications
  const mailProvider = operations.providers.mailProvider;
  const insights = operations.providers.insights;
  if (!mailProvider || operations.isSystemAccountByUsername(whoChangedIt)) {
    return;
  }
  const senderMember = organization.member(whoChangedItId);
  const mailAddress = await senderMember.getMailAddress();
  if (!mailAddress) {
    return;
  }
  const basedir = operations.config.typescript.appDirectory;
  const operationsMail = operations.getOperationsMailAddress();
  const companySpecific = getCompanySpecificDeployment();
  const largeTeamProtectionDetailsLink = companySpecific?.strings?.largeTeamProtectionDetailsLink;
  const config = operations.config;
  await sendEmail(config, insights, basedir, mailProvider, mailAddress, operationsMail, {
    repository: repositoryBody,
    whoChangedIt,
    teamName,
    reason,
    companyName: config.brand.companyName,
    largeTeamProtectionDetailsLink,
  });
}

async function sendEmail(config, insights, basedir, mailProvider: IMailProvider, to, operationsMail: string, body) {
  body.reason = `You are receiving this e-mail because you changed the permissions on the ${body.teamName} GitHub team, triggering this action.`;
  body.headline = 'Team permission change reverted';
  body.notification = 'warning';
  body.app = `${config.brand.companyName} GitHub`;
  const mail: IAutomaticTeamsMail = {
    to,
    cc: operationsMail,
    subject: `Team permission change for ${body.repository.full_name} repository reverted`,
    category: ['error', 'repos'],
  };
  let mailContent = null;
  try {
    mailContent = await RenderHtmlMail(basedir, 'largeTeamProtected', body);
  } catch(renderError) {
    insights.trackException({
      exception: renderError,
      properties: {
        content: body,
        eventName: 'JobAutomaticTeamsLargeTeamPermissionBlockMailRenderFailure',
      },
    });
    throw renderError;
  }
  mail.content = mailContent;
  let customData: ICustomDataEventName = {
    content: body,
    receipt: '',
  };
  try {
    const mailResult = await mailProvider.sendMail(mail);
    customData.receipt = mailResult;
  } catch (mailError) {
    customData.eventName = 'JobAutomaticTeamsLargeTeamPermissionBlockMailFailure';
    insights.trackException({ exception: mailError, properties: customData });
    throw mailError;
  }
  insights.trackEvent({ name: 'JobAutomaticTeamsLargeTeamPermissionBlockMailSuccess', properties: customData });
}

async function setTeamPermission(operations: Operations, organization: Organization, repositoryBody: any, teamId, necessaryPermission, reason): Promise<boolean> {
  const { customizedTeamPermissionsWebhookLogic } = operations.providers;
  const repoName = repositoryBody.name;
  const repoId = repositoryBody?.id;
  const orgName = organization.name;
  const repository = organization.repository(repoName, { id: repoId });
  if (customizedTeamPermissionsWebhookLogic) {
    const shouldSkipEnforcement = await customizedTeamPermissionsWebhookLogic.shouldSkipEnforcement(repository);
    if (shouldSkipEnforcement) {
      console.log(`Customized logic for team permissions: skipping enforcement for repository ${repository.id}`);
      return false;
    }
  }
  const description = `setting permission level ${necessaryPermission} for the team with ID ${teamId} on the repository ${repoName} inside the ${orgName} GitHub org because ${reason}`;
  const insights = operations.insights;
  let error = null;
  try {
    await repository.setTeamPermission(teamId, necessaryPermission);
  } catch (setError) {
    error = setError;
  }
  const eventRoot = 'AutomaticRepoPermissionSet';
  const eventName = eventRoot + error ? 'Success' : 'Failure';
  if (error) {
    error.description = description;
    console.warn(`${eventName} ${description}`);
  } else {
    console.log(`${eventName} ${description}`);
  }
  if (insights) {
    insights.trackEvent({
      name: eventName,
      properties: {
        success: !!error,
        reason: reason,
        description: description,
      },
    });
  }
  if (error) {
    throw error;
  }
  return true;
}
