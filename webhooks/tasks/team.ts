//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// TEAM changes

import moment from 'moment';

import { WebhookProcessor } from '../organizationProcessor';
import { Operations } from '../../business';
import { Organization } from '../../business';
import { permissionsObjectToValue } from '../../transitional';

// When teams are added or removed on GitHub, refresh the organization's list of
// teams as well as the cross-organization view of the teams.

// TODO: connect to query cache
// TODO: consider whether to slowly kick off Redis cache updates, too

export default class TeamWebhookProcessor implements WebhookProcessor {
  filter(data: any) {
    let eventType = data.properties.event;
    return eventType === 'team';
  }

  async run(operations: Operations, organization: Organization, data: any): Promise<boolean> {
    const queryCache = operations.providers.queryCache;
    const event = data.body;
    let refresh = false;
    let expectedAfterRefresh = false;
    const teamId = event.team.id;
    const teamIdAsString = event.team.id.toString();
    const organizationIdAsString = event.organization.id.toString();
    let addOrUpdate = false;
    if (event.action === 'created') {
      console.log(`team created: ${event.team.name} in organization ${event.organization.login} by ${event.sender.login}`);
      refresh = true;
      expectedAfterRefresh = true;
      addOrUpdate = true;
    } else if (event.action === 'deleted') {
      console.log(`team DELETED: ${event.team.name} in organization ${event.organization.login} by ${event.sender.login}`);
      refresh = true;
      try {
        if (organizationIdAsString === organization.id.toString() && queryCache && queryCache.supportsTeams) {
          await queryCache.removeOrganizationTeam(organizationIdAsString, teamIdAsString);
        }
      } catch (queryCacheError) {
        console.dir(queryCacheError);
      }
    } else if (event.action === 'edited') {
      addOrUpdate = true;
      if (event.changes && event.changes.repository && event.changes.repository.permissions && event.changes.repository.permissions.from && queryCache && queryCache.supportsTeamPermissions) {
        const oldRepositoryPermissionLevel = permissionsObjectToValue(event.changes.repository.permissions.from);
        const newRepositoryPermissionLevel = permissionsObjectToValue(event.repository.permissions);
        console.log(`team ${event.team.name} permission level for repo ${event.repository.name} changed from ${oldRepositoryPermissionLevel} to ${newRepositoryPermissionLevel}`);
        const isPrivate = event.repository.private as boolean;
        const repoName = event.repository.name as string;
        const orgId = event.repository.owner.id as number;
        if (operations.isOrganizationManagedById(orgId)) {
          await queryCache.addOrUpdateTeamsPermission(
            organizationIdAsString,
            event.repository.id.toString(),
            isPrivate,
            repoName,
            event.team.id.toString(),
            newRepositoryPermissionLevel);
        }
      }
    } else if (event.action === 'added_to_repository') {
      console.log(`team got permission to repo: ${event.team.name} for repo ${event.repository.name} in organization ${event.organization.login} by ${event.sender.login}`);
      if (queryCache && queryCache.supportsTeamPermissions) {
        const isPrivate = event.repository.private as boolean;
        const repoName = event.repository.name as string;
        const orgId = event.repository.owner.id as number;
        if (operations.isOrganizationManagedById(orgId)) {
          await queryCache.addOrUpdateTeamsPermission(
            organizationIdAsString,
            event.repository.id.toString(),
            isPrivate,
            repoName,
            event.team.id.toString(),
            permissionsObjectToValue(event.repository.permissions)); // equiv to event.team.permission as GitHubRepositoryPermission
        }
      }
    } else if (event.action === 'removed_from_repository') {
      console.log(`team lost permission to repo: ${event.team.name} for repo ${event.repository.name} in organization ${event.organization.login} by ${event.sender.login}`);
      if (queryCache && queryCache.supportsTeamPermissions) {
        await queryCache.removeRepositoryTeam(organizationIdAsString, event.team.id.toString(), event.repository.id.toString());
      }
    } else {
      console.log('other team condition:');
      console.dir(data);
    }

    if (addOrUpdate) {
      try {
        if (organizationIdAsString === organization.id.toString() && queryCache && queryCache.supportsTeams) {
          await queryCache.addOrUpdateTeam(organizationIdAsString, teamIdAsString, event.team);
        }
      } catch (queryCacheError) {
        console.dir(queryCacheError);
      }
    }

    if (refresh) {
      const startingRefresh = moment();
      // organization.getTeams(immediateRefreshOptions, () => {
      //   console.log('refreshing teams list after add or remove operations');
      //   const now = moment();
      //   const elapsedSeconds = Math.ceil(moment.duration(now.diff(startingRefresh)).asSeconds());
      //   console.log(`elapsed seconds since kicked off the refresh: ${elapsedSeconds}`);
      //   const crossOrgRefreshOptions = {
      //     backgroundRefresh: false,
      //     maxAgeSeconds: elapsedSeconds || 15,
      //   };
      //   operations.getTeams(null, crossOrgRefreshOptions, (crossOrgRefreshError, allTeams) => {
      //     if (crossOrgRefreshError) {
      //       console.log('cross-org team refresh encountered an error:');
      //       console.dir(crossOrgRefreshError);
      //     } else {
      //       console.log(`refreshed cross-org teams list with ${elapsedSeconds} seconds buffer`);
      //       if (expectedAfterRefresh && allTeams.has(teamId)) {
      //         console.log('Verified that the team ' + teamId + ' was present in the cross-org result');
      //       }
      //     }
      //   });
      // });
    }

    // Immediately, to help delete the ticket
    return true;
  }
}
