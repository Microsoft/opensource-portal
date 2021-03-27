//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IProviders } from '../transitional';

/*eslint no-console: ["error", { allow: ["log", "dir"] }] */

// ----------------------------------------------------------------------------
// Onboarding helper
// ----------------------------------------------------------------------------
// This file is only used when an organization has its "onboarding" value set.
// It helps present a dev or devops person with the mapping of team IDs to team
// names for the organization. This is not actually a middleware route but
// rather a configuration/app initialization method just stored here to keep it
// out of the way.
// ----------------------------------------------------------------------------
export default async function Onboard(app, config) {
  const { operations } = app.settings.providers as IProviders;
  const onboardingOrganizations = config.github.organizations?.onboarding || [];
  for (const orgEntry of onboardingOrganizations) {
    if (orgEntry && orgEntry.name && orgEntry.ownerToken) {
      let s = 'Organization Onboarding Helper for "' + orgEntry.name + '":\n';
      for (var key in orgEntry) {
        s += '- ' + key + ': ';
        s += orgEntry[key] !== undefined ? 'value set' : 'undefined';
        s += '\n';
      }
      const organization = operations.getOnboardingOrganization(orgEntry.name);
      try {
        const teams = await organization.getTeams();
        s +=
          'Here is a mapping from team ID to team slug (based on the name),\nto help with selecting the team IDs needed to run the portal\nsuch as the repo approvers and sudoers teams.\n\n';
        for (let j = 0; j < teams.length; j++) {
          const team = teams[j];
          s += team.id + ': ' + team.slug + '\n';
        }
      } catch (error) {
        console.log(
          `Error retrieving teams for the organization ${orgEntry.name}`
        );
        console.dir(error);
      }
      console.log(s);
    } else {
      console.log(
        'An org requires that its NAME and TOKEN configuration parameters are set before onboarding can begin.'
      );
    }
  }
  console.log('This concludes the execution of the onboarding helper.');
}
