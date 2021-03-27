//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Team, GitHubTeamRole, ITeamMembershipRoleState } from '../business';
import { Operations } from '../business';
import { IndividualContext } from '../user';
import { ErrorHelper, NoCacheNoBackground } from '../transitional';
import { addArrayToSet } from '../utils';
import { IMail } from '../lib/mailProvider';
import { OrganizationMembershipState } from '../business';

interface ISelfServiceAllowedResult {
  currentMaintainerCount: number;
  currentMemberCount: number;
}

export interface ISelfServiceTeamMemberToMaintainerUpgradeOptions {
  operations: Operations;
  team: Team;
}

export default class SelfServiceTeamMemberToMaintainerUpgrades {
  #team: Team;
  #operations: Operations;

  constructor(options: ISelfServiceTeamMemberToMaintainerUpgradeOptions) {
    if (!options.operations || !options.team) {
      throw new Error('options.operations and options.team are required');
    }
    this.#operations = options.operations;
    this.#team = options.team;
  }

  isSelfServiceUpgradeEnabled(): boolean {
    return this.#operations.allowSelfServiceTeamMemberToMaintainerUpgrades();
  }

  maximimumAllowedMembers(): number {
    const value = this.#operations.config.github?.teams
      ?.maximumMembersToAllowUpgrade;
    return value ? Number(value) : 0;
  }

  maximumAllowedMaintainers(): number {
    const value = this.#operations.config.github?.teams
      ?.maximumMaintainersToAllowUpgrade;
    return value ? Number(value) : 0;
  }

  async isTeamEligible(
    cacheOk?: boolean
  ): Promise<ISelfServiceAllowedResult | string> {
    const cacheOptions = cacheOk ? {} : NoCacheNoBackground;
    const team = this.#team;
    const maintainersCount = (await team.getMaintainers(cacheOptions)).length;
    if (maintainersCount > this.maximumAllowedMaintainers()) {
      return `There are currently ${maintainersCount} maintainers of the team. Self-service upgrade is only available if there are ${this.maximumAllowedMaintainers()} or fewer maintainers.`;
    }
    const membersCount = (await team.getMembers(cacheOptions)).length;
    if (membersCount > this.maximimumAllowedMembers()) {
      return `There are currently ${membersCount} members of the team. Self-service upgrade is only available if there are ${this.maximimumAllowedMembers()} or fewer members.`;
    }
    return {
      currentMaintainerCount: maintainersCount,
      currentMemberCount: membersCount,
    };
  }

  async isUserTeamMember(login: string): Promise<boolean> {
    const membership = (await this.#team.getMembership(
      login,
      NoCacheNoBackground
    )) as ITeamMembershipRoleState;
    return (
      membership.state === OrganizationMembershipState.Active &&
      membership.role === GitHubTeamRole.Member
    );
  }

  async validateUserCanSelfServicePromote(
    individualContext: IndividualContext
  ): Promise<ISelfServiceAllowedResult> {
    if (!this.isSelfServiceUpgradeEnabled()) {
      throw new Error('Self-service upgrade is not available');
    }
    if (
      !individualContext.corporateIdentity ||
      !individualContext.corporateIdentity.id ||
      !individualContext.getGitHubIdentity().id
    ) {
      throw new Error('The authenticated user is not properly linked');
    }
    const login = individualContext.getGitHubIdentity().username;
    const teamEligiblityResult:
      | ISelfServiceAllowedResult
      | string = await this.isTeamEligible();
    if (typeof teamEligiblityResult === 'string') {
      throw new Error(teamEligiblityResult as string);
    }
    const userIsTeamMember = await this.isUserTeamMember(login);
    if (!userIsTeamMember) {
      throw new Error(
        'The user is not a member of the team and so cannot be upgraded'
      );
    }
    return teamEligiblityResult as ISelfServiceAllowedResult;
  }

  async upgrade(individualContext: IndividualContext): Promise<void> {
    const team = this.#team;
    const operations = this.#operations;
    await team.getDetails();
    const canUpgradeDetails = await this.validateUserCanSelfServicePromote(
      individualContext
    );
    const login = individualContext.getGitHubIdentity().username;
    const { queryCache } = operations.providers;
    try {
      await team.addMaintainer(login);
      if (queryCache && queryCache.supportsTeamMembership) {
        try {
          await queryCache.addOrUpdateTeamMember(
            String(team.organization.id),
            String(team.id),
            String(individualContext.getGitHubIdentity().id),
            GitHubTeamRole.Maintainer,
            login,
            individualContext.getGitHubIdentity().avatar
          );
        } catch (ignoreQueryCacheUpdateError) {
          console.log('ignoreQueryCacheUpdateError:');
          console.warn(ignoreQueryCacheUpdateError);
        }
      }
    } catch (upgradeError) {
      throw ErrorHelper.WrapError(
        upgradeError,
        `Self-service team maintainer upgrade for the GitHub account ${login} in the ${team.organization.name} org team ${team.name} failed: ${upgradeError}`
      );
    }
    // Refresh for display
    try {
      await team.getMaintainers(NoCacheNoBackground);
    } catch (ignoreTeamRefreshErrors) {
      console.log('ignoreTeamRefreshErrors:');
      console.warn(ignoreTeamRefreshErrors);
    }
    // Send a notification mail to many folks... with this logic:
    // If the Team Maintainer count before the upgrade was > 0, just notify the current team maintainers.
    // Otherwise, notify all Team Members of the upgrade.
    try {
      const maintainers = await team.getMaintainers();
      const idsToNotify = new Set<string>(
        maintainers.map((maintainer) => String(maintainer.id))
      );
      let notifyDescription = `All of the Team Maintainers for the ${team.name} team are being notified in this mail.`;
      if (canUpgradeDetails.currentMaintainerCount === 0) {
        const members = await team.getMembers();
        addArrayToSet(
          idsToNotify,
          members.map((member) => String(member.id))
        );
        notifyDescription = `All of the Team Members and Team Maintainers for the ${team.name} team are being notified in this mail.`;
      }
      const links = await operations.getLinksFromThirdPartyIds(
        Array.from(idsToNotify.values())
      );
      const thirdPartyLoginToLink = new Map();
      links.map((userLink) => {
        const login = userLink.thirdPartyUsername.toLowerCase();
        thirdPartyLoginToLink.set(login, userLink);
      });
      const mailAddresses = await operations.getMailAddressesFromCorporateUsernames(
        links.map((link) => link.corporateUsername)
      );
      const opsAddress = operations.getOperationsMailAddress();
      const companyName = operations.config.brand.companyName;
      const identifierForRequester =
        individualContext.corporateIdentity.displayName ||
        individualContext.corporateIdentity.username;
      const mail: IMail = {
        to: mailAddresses,
        cc: opsAddress ? [opsAddress] : null,
        subject: `${team.organization.name}/${team.name}: GitHub Team Member ${identifierForRequester} upgraded themselves to Team Maintainer`,
        content: await operations.emailRender(
          'teamMemberSelfServiceMaintainerUpgrade',
          {
            reason: `This is a required operational notification: ${identifierForRequester} used a self-service permission upgrade feature. ${notifyDescription}`,
            headline: `Team maintainer upgrade`,
            notification: 'information',
            app: `${companyName} GitHub`,
            team,
            link: individualContext.link,
            organization: team.organization,
            identifierForRequester,
            maintainers,
            thirdPartyLoginToLink,
          }
        ),
      };
      await this.#operations.sendMail(mail);
    } catch (mailError) {
      console.log('mailError:');
      console.warn(mailError);
    }
    const insights = this.#operations.insights;
    if (insights) {
      insights.trackMetric({
        name: 'TeamSelfServiceMemberToMaintainerUpgrades',
        value: 1,
      });
    }
  }
}
