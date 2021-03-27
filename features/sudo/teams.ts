//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { OrganizationSudo } from '.';
import {
  GitHubTeamRole,
  ICorporateLink,
  ITeamMembershipRoleState,
  Organization,
} from '../../business';
import { ErrorHelper, IProviders } from '../../transitional';

export class OrganizationSudoGitHubTeams extends OrganizationSudo {
  constructor(providers: IProviders, organization: Organization) {
    super(providers, organization);
  }

  async isSudoer(githubLogin: string, link?: ICorporateLink): Promise<boolean> {
    if (this.isSudoEnvironmentOff()) {
      return false;
    }

    const organization = this.organization;
    const sudoerTeam = organization.sudoersTeam;
    if (!sudoerTeam) {
      return false;
    }

    let membership: GitHubTeamRole = null;
    try {
      const response = await sudoerTeam.getMembershipEfficiently(githubLogin);
      if (response && (response as ITeamMembershipRoleState).role) {
        membership = (response as ITeamMembershipRoleState).role;
      }
    } catch (getMembershipError) {
      if (ErrorHelper.IsNotFound(getMembershipError)) {
        return false;
      }
      throw getMembershipError;
    }
    const isKnownMembership =
      membership === GitHubTeamRole.Member ||
      membership === GitHubTeamRole.Maintainer;
    if (membership && isKnownMembership) {
      return isKnownMembership;
    } else if (membership) {
      throw new Error(
        `Cannot determine sudo status for ${githubLogin}, unrecognized membership type: ${membership}`
      );
    } else {
      return false;
    }
  }
}
