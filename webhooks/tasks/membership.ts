//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { WebhookProcessor } from '../organizationProcessor';
import { Operations, Organization } from '../../business';
import { IProviders, GitHubTeamRole } from '../../interfaces';

export default class MembershipWebhookProcessor implements WebhookProcessor {
  filter(data: any) {
    let eventType = data.properties.event;
    return eventType === 'membership';
  }

  async run(operations: Operations, organization: Organization, data: any): Promise<any> {
    const providers = operations.providers as IProviders;
    const queryCache = providers.queryCache;
    const event = data.body;
    const organizationId = event.organization.id as number;
    if (!operations.isOrganizationManagedById(organizationId)) {
      console.log(`skipping organization ID ${organizationId} which is not directly managed: ${event.organization.login}`);
      return true;
    }
    if (event.action === 'added' && event.scope === 'team') {
      const userIdAsString = event.member.id.toString();
      const userLogin = event.member.login;
      const organizationIdAsString = event.organization.id.toString();
      const teamIdAsString = event.team.id.toString();
      const avatar = event.member.avatar_url;
      console.log(`team member added: login=${userLogin} id=${userIdAsString} team id=${teamIdAsString} slug=${event.team.slug} org: ${event.organization.login}`);
      try {
        if (queryCache && queryCache.supportsTeamMembership) {
          await queryCache.addOrUpdateTeamMember(organizationIdAsString, teamIdAsString, userIdAsString, GitHubTeamRole.Member, userLogin, avatar);
        }
      } catch (queryCacheError) {
        console.dir(queryCacheError);
      }
    } else if (event.action === 'removed' && event.scope === 'team') {
      const userIdAsString = event.member.id.toString();
      const userLogin = event.member.login;
      const organizationIdAsString = event.organization.id.toString();
      const teamIdAsString = event.team.id.toString();
      console.log(`REMOVED: team member: login=${userLogin} id=${userIdAsString} team id=${teamIdAsString} slug=${event.team.slug} org: ${event.organization.login}`);
      try {
        if (queryCache && queryCache.supportsTeamMembership) {
          await queryCache.removeTeamMember(organizationIdAsString, teamIdAsString, userIdAsString);
        }
      } catch (queryCacheError) {
        console.dir(queryCacheError);
      }
    } else {
      console.log('unprocessed team membership event');
      console.dir(data);
    }

    // update the team in question
    /*
    const immediateRefreshOptions = {
      backgroundRefresh: false,
      maxAgeSeconds: 0,
    };
    */
    // console.log(`refreshing members in the team ${data.body.team.name} ${data.body.team.id} list`);
    // const team = organization.team(data.body.team.id);
    // await team.getDetails();
    // team.getMembers({
    //   backgroundRefresh: false,
    //   maxAgeSeconds: 0.1,
    // }, (getMembersError, members) => {
    //   let num = '';
    //   if (!getMembersError && members && members.length) {
    //     num = members.length;
    //   }
    //   console.log(`refreshed ${num} team members, getting maintainers`);
    //   team.getMembers({
    //     role: 'maintainer',
    //     backgroundRefresh: false,
    //     maxAgeSeconds: 0.1,
    //   }, (getMaintainersError, maintainers) => {
    //     let num2 = '';
    //     if (!getMaintainersError && maintainers && maintainers.length) {
    //       num2 = members.length;
    //     }
    //     console.log(`refreshed ${num2} team maintainers`);
    //   });
    // });
    return true;
  }
}
