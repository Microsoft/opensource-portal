//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// ORGANIZATION membership and ownership

/*eslint no-console: ["error", { allow: ["dir", "log", "warn"] }] */

import { Operations } from '../../business';
import {
  Organization,
  OrganizationMembershipRole,
  OrganizationMembershipState,
} from '../../business';
import { IProviders, NoCacheNoBackground } from '../../transitional';
import { WebhookProcessor } from '../organizationProcessor';

// NOTE: unfortunately role changes from admin->member or member->admin do not fire GitHub hooks

function getRoleFromString(role: string): OrganizationMembershipRole {
  switch (role) {
    case 'admin':
      return OrganizationMembershipRole.Admin;
    case 'member':
      return OrganizationMembershipRole.Member;
    default:
      throw new Error(`OrganizationMembershipRole ${role} is not supported`);
  }
}

export default class OrganizationWebhookProcessor implements WebhookProcessor {
  filter(data: any) {
    let eventType = data.properties.event;
    return eventType === 'organization';
  }

  async run(
    operations: Operations,
    organization: Organization,
    data: any
  ): Promise<boolean> {
    const providers = operations.providers as IProviders;
    const queryCache = providers.queryCache;

    const event = data.body;
    let refresh = false;
    if (event.action === 'member_invited') {
      if (
        !event.invitation ||
        !event.invitation.inviter ||
        !event.invitation.inviter.login
      ) {
        // should no longer be an issue per GitHub in September 2019
      }
      console.log(
        `org member invite by ${event.invitation.inviter.login}; ghu ${event.invitation.login} role ${event.invitation.role} ghid ${event.invitation.id} org: ${event.organization.login}`
      );
    } else if (event.action === 'member_added') {
      console.log(
        `org member added; ghu ${event.membership.user.login} role ${event.membership.role} state ${event.membership.state} ghid ${event.membership.user.id} org: ${event.organization.login}`
      );
      if (
        event.membership.state === 'active' ||
        event.membership.state === 'pending'
      ) {
        // triple-check the state; GitHub is sending new memberships are PENDING and not ACTIVE now.
        const login = event.membership.user.login;
        const liveMembership = await organization.getMembership(
          login,
          NoCacheNoBackground
        );
        let state = null;
        if (liveMembership) {
          console.log(`live membership: state=${liveMembership.state}`);
          state = liveMembership.state;
        }
        if (state === OrganizationMembershipState.Active) {
          const userIdAsString = event.membership.user.id.toString();
          const organizationIdAsString = event.organization.id.toString();
          try {
            if (queryCache && queryCache.supportsOrganizationMembership) {
              const role = getRoleFromString(event.membership.role);
              await queryCache.addOrUpdateOrganizationMember(
                organizationIdAsString,
                role,
                userIdAsString
              );
              console.log(
                `OK: query cache added orgid=${organizationIdAsString}, userid=${userIdAsString}, role=${role}`
              );
            } else {
              console.warn('the organization does not use the query cache');
            }
          } catch (queryCacheError) {
            console.dir(queryCacheError);
          }
        } else {
          console.log(`Live state is still not right to insert: ${state}`);
        }
      }
      refresh = true;
    } else if (event.action === 'member_removed') {
      console.log(
        `org member REMOVED; ghu ${event.membership.user.login} role ${event.membership.role} state ${event.membership.state} ghid ${event.membership.user.id} org: ${event.organization.login}`
      );
      const userIdAsString = event.membership.user.id.toString();
      const organizationIdAsString = event.organization.id.toString();
      try {
        if (queryCache && queryCache.supportsOrganizationMembership) {
          await queryCache.removeOrganizationMember(
            organizationIdAsString,
            userIdAsString
          );
          console.log(
            `OK: query cache removed orgid=${organizationIdAsString}, userid=${userIdAsString}`
          );
        } else {
          console.warn('the organization does not use the query cache');
        }
      } catch (queryCacheError) {
        console.dir(queryCacheError);
      }
      refresh = true;
    } else {
      console.log('Unsupported org event:');
      console.dir(data);
    }
    if (refresh) {
      // const orgName = organization.name;
      // console.log(`refreshing ${orgName} org members list`);
      // const immediateRefreshOptions = {
      //   backgroundRefresh: false,
      //   maxAgeSeconds: 0.01,
      // };
      // return organization.getMembers(immediateRefreshOptions).then(ok => {
      //   console.log(`refreshed membership list for the org ${orgName}, will refresh x-org immediately`);
      //   return operations.getMembers(immediateRefreshOptions).then(done => {
      //     console.log('refreshed x-org memberships');
      //   });
      // }).catch(error => {
      //   // ignore error
      //   return callback();
      // });
    }

    return true;
  }
}
