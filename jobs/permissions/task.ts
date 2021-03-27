//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

import { shuffle } from 'lodash';

import { TeamPermission } from '../../business/teamPermission';
import { IReposJob, IReposJobResult } from '../../app';
import AutomaticTeamsWebhookProcessor from '../../webhooks/tasks/automaticTeams';
import {
  GitHubRepositoryPermission,
  RepositoryLockdownState,
  RepositoryMetadataEntity,
} from '../../entities/repositoryMetadata/repositoryMetadata';
import { sleep } from '../../utils';
import { ErrorHelper } from '../../transitional';

// Permissions processing: visit all repos and make sure that any designated read, write, admin
// teams for the organization are present on every repo. This job is designed to be run relatively
// regularly but is not looking to answer "the truth" - it will use the cache of repos and other
// assets to not abuse GitHub and its API exhaustively. Over time repos will converge to having
// the right permissions.
//
// If a repository is "compliance locked", the system teams are not enforced until the lock is removed.

const maxParallelism = 1;

const delayBetweenSeconds = 1;

export default async function permissionsRun({
  providers,
}: IReposJob): Promise<IReposJobResult> {
  const { operations, repositoryMetadataProvider } = providers;
  for (const organization of shuffle(
    Array.from(operations.organizations.values())
  )) {
    console.log(
      `Reviewing permissions for all repos in ${organization.name}...`
    );
    try {
      const repos = await organization.getRepositories();
      console.log(`Repos in the ${organization.name} org: ${repos.length}`);
      let z = 0;
      const automaticTeams = new AutomaticTeamsWebhookProcessor();
      for (let repo of repos) {
        console.log(`${repo.organization.name}/${repo.name}`);
        sleep(1000 * delayBetweenSeconds);
        const cacheOptions = {
          maxAgeSeconds: 10 * 60 /* 10m */,
          backgroundRefresh: false,
        };
        ++z;
        if (z % 250 === 1) {
          console.log('. ' + z);
        }
        const {
          specialTeamIds,
          specialTeamLevels,
        } = automaticTeams.processOrgSpecialTeams(repo.organization);
        let permissions: TeamPermission[] = null;
        try {
          permissions = await repo.getTeamPermissions(cacheOptions);
        } catch (getError) {
          if (getError.status == /* loose */ 404) {
            console.log(`Repo gone: ${repo.organization.name}/${repo.name}`);
          } else {
            console.log(
              `There was a problem getting the permissions for the repo ${repo.name} from ${repo.organization.name}`
            );
            console.dir(getError);
          }
          return;
        }
        const { customizedTeamPermissionsWebhookLogic } = providers;
        if (customizedTeamPermissionsWebhookLogic) {
          const shouldSkipEnforcement = await customizedTeamPermissionsWebhookLogic.shouldSkipEnforcement(
            repo
          );
          if (shouldSkipEnforcement) {
            console.log(
              `Customized logic for team permissions: skipping enforcement`
            );
            return;
          }
        }
        const currentPermissions = new Map<
          number,
          GitHubRepositoryPermission
        >();
        permissions.forEach((entry) => {
          currentPermissions.set(Number(entry.team.id), entry.permission);
        });
        const teamsToSet = new Set<number>();
        specialTeamIds.forEach((specialTeamId) => {
          if (!currentPermissions.has(specialTeamId)) {
            teamsToSet.add(specialTeamId);
          } else if (
            isAtLeastPermissionLevel(
              currentPermissions.get(specialTeamId),
              specialTeamLevels.get(specialTeamId)
            )
          ) {
            // The team permission is already acceptable
          } else {
            console.log(
              `Permission level for ${specialTeamId} is not good enough, expected ${specialTeamLevels.get(
                specialTeamId
              )} but currently ${currentPermissions.get(specialTeamId)}`
            );
            teamsToSet.add(specialTeamId);
          }
        });
        const setArray = Array.from(teamsToSet.values());
        for (let teamId of setArray) {
          const newPermission = specialTeamLevels.get(teamId);
          console.log(
            `adding ${teamId} team with permission ${newPermission} to the repo ${repo.name}`
          );
          try {
            await repo.setTeamPermission(
              teamId,
              newPermission as GitHubRepositoryPermission
            );
          } catch (error) {
            if (ErrorHelper.IsNotFound(error)) {
              console.log(
                `the team ID ${teamId} could not be found when setting to repo ${repo.name} in org ${organization.name} and should likely be removed from config...`
              );
            } else {
              console.log(`${repo.name}`);
              console.dir(error);
              throw error;
            }
          }
        }
      }
    } catch (processOrganizationError) {
      console.dir(processOrganizationError);
      console.log(
        `moving past ${organization.name} processing due to error...`
      );
    }
  }
  return {};
}

function isAtLeastPermissionLevel(value, expected) {
  if (value !== 'admin' && value !== 'push' && value !== 'pull') {
    throw new Error(
      `The permission type ${value} is not understood by isAtLeastPermissionLevel`
    );
  }
  if (value === expected) {
    return true;
  }
  // Admin always wins
  if (value === 'admin') {
    return true;
  } else if (expected === 'admin') {
    return false;
  }
  if (expected === 'write' && value === expected) {
    return true;
  }
  if (expected === 'read') {
    return true;
  }
  return false;
}
