//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import moment from 'moment';

import { IReposJob } from '../../app';

// Organization invitations cleanup: remove any invitations that are older than a
// set period of time from the organization.

const defaultMaximumInvitationAgeDays = 4;

export default async function cleanup({ providers }: IReposJob): Promise<void> {
  const insights = providers.insights;
  let maximumInvitationAgeDays = defaultMaximumInvitationAgeDays;
  const { config, operations } = providers;
  if (
    config.github &&
    config.github.jobs &&
    config.github.jobs.cleanup &&
    config.github.jobs.cleanup.maximumInvitationAgeDays
  ) {
    maximumInvitationAgeDays =
      config.github.jobs.cleanup.maximumInvitationAgeDays;
  }
  const maximumAgeMoment = moment().subtract(maximumInvitationAgeDays, 'days');
  const organizations = operations.getOrganizations();
  let removedInvitations = 0;
  for (let organization of organizations) {
    let invitations: any[];
    try {
      invitations = await organization.getMembershipInvitations();
    } catch (getInvitationsError) {
      insights.trackException({ exception: getInvitationsError });
      console.dir(getInvitationsError);
      continue;
    }
    if (!invitations || invitations.length === 0) {
      continue;
    }
    const invitationsToRemove = [];
    let emailInvitations = 0;
    for (let i = 0; i < invitations.length; i++) {
      const invite = invitations[i];
      const createdAt = moment(invite.created_at);
      if (createdAt.isBefore(maximumAgeMoment)) {
        if (invite.login) {
          invitationsToRemove.push(invite.login);
        } else {
          ++emailInvitations;
          console.warn(
            `An e-mail based invitation to ${invite.email} cannot be automatically canceled`
          );
        }
        const data = {
          createdAt: createdAt.format(),
          invitedAgo: createdAt.fromNow(),
          login: invite.login,
          inviter:
            invite && invite.inviter && invite.inviter.login
              ? invite.inviter.login
              : undefined,
          role: invite.role,
          emailInvited: invite.email,
        };
        const eventName = invite.login
          ? 'JobOrganizationInviteCleanupInvitationNeeded'
          : 'JobOrganizationInviteCleanupInvitationNotUser';
        insights.trackEvent({
          name: eventName,
          properties: data,
        });
      }
    }
    console.log(
      `Organization ${organization.name} has ${invitationsToRemove.length} expired invitations out of ${invitations.length} total invitations pending`
    );
    if (emailInvitations) {
      console.warn(
        `Organization ${organization.name} has ${emailInvitations} e-mail based invitations that cannot be canceled through this job`
      );
    }
    for (let login of invitationsToRemove) {
      try {
        await organization.removeMember(login);
      } catch (removeError) {
        insights.trackException({ exception: removeError });
        insights.trackEvent({
          name: 'JobOrganizationInvitationsCleanupInvitationFailed',
          properties: {
            login: login,
            message: removeError.message,
          },
        });
      }
    }
  }
  console.log(
    `Job finishing. Removed ${removedInvitations} expired invitations.`
  );
  insights.trackMetric({
    name: 'JobOrganizationInvitationsExpired',
    value: removedInvitations,
  });
}
