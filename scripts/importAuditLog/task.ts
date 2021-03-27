//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import fs from 'fs';
import csv from 'csv-parser';
import moment from 'moment';

import { IReposJob } from '../../app';
import { AuditLogRecord } from '../../entities/auditLogRecord/auditLogRecord';
import { sleep } from '../../utils';
import { Operations } from '../../business/operations';
import { Organization } from '../../business/organization';
import { AuditLogSource } from '../../entities/auditLogRecord';

const restingBeforeRequestMs = 200;

const veryOldCacheOK = {
  maxAgeSeconds: 60 * 60 * 6 /* 6 hours old */,
};

async function tryGetUserId(
  operations: Operations,
  usernamesToId: Map<string, boolean | number>,
  username: string
): Promise<string> {
  const knownUsernameToId = usernamesToId.get(username);
  if (knownUsernameToId === undefined) {
    try {
      await sleep(restingBeforeRequestMs);
      const account = await operations.getAccountByUsername(username);
      usernamesToId.set(username, account.id);
      return account.id.toString();
    } catch (error) {
      if (error.status === 404) {
        usernamesToId.set(username, false);
      }
    }
  } else if (knownUsernameToId === false) {
    return null;
  } else {
    return knownUsernameToId.toString();
  }
}

async function tryGetRepositoryId(
  operations: Operations,
  reposToIds: Map<string, boolean | number>,
  organization: Organization,
  repoName: string
): Promise<string> {
  const knownToId = reposToIds.get(repoName);
  if (knownToId === undefined) {
    try {
      await sleep(restingBeforeRequestMs);
      const repository = organization.repository(repoName);
      const repoDetails = await repository.getDetails();
      if (repoDetails.headers['x-ratelimit-remaining']) {
        const rl = Number(repoDetails.headers['x-ratelimit-remaining']);
        if (rl < 5000) {
          console.log('Slowing down and resting a minute...');
          await sleep(1000 * 60);
        }
      }
      reposToIds.set(repoName, repoDetails.id);
      return repoDetails.id.toString();
    } catch (error) {
      if (error.status === 404) {
        reposToIds.set(repoName, false);
      }
    }
  } else if (knownToId === false) {
    return null;
  } else {
    return knownToId.toString();
  }
}

async function tryGetTeamId(
  teamsToIds: Map<string, boolean | number>,
  organization: Organization,
  teamName: string
): Promise<string> {
  const knownToId = teamsToIds.get(teamName);
  if (knownToId === undefined) {
    try {
      await sleep(restingBeforeRequestMs);
      const team = await organization.getTeamFromName(teamName, veryOldCacheOK);
      teamsToIds.set(teamName, team.id);
      return team.id.toString();
    } catch (error) {
      if (error.status === 404) {
        teamsToIds.set(teamName, false);
      } else if (error.status === 301 && error.team) {
        teamsToIds.set(teamName, error.team.id);
        return error.team.id.toString();
      }
    }
  } else if (knownToId === false) {
    return null;
  } else {
    return knownToId.toString();
  }
}

export default async function migration({ providers }: IReposJob) {
  const { linkProvider, operations } = providers;

  const rows = await parseCsv(process.argv[2]);
  console.log(`parsed ${rows.length} rows`);

  const usernamesToId = new Map<string, boolean | number>();
  const allLinks = await linkProvider.getAll();
  for (const link of allLinks) {
    usernamesToId.set(
      link.thirdPartyUsername.toLowerCase(),
      Number(link.thirdPartyId)
    );
  }

  const reposToId = new Map<string, boolean | number>();
  const teamToId = new Map<string, boolean | number>();
  const actionNames = new Set<string>();
  let i = 0;
  for (const row of rows) {
    ++i;
    const remainder = { ...row };

    // created_at: epoch
    const created = new Date(parseInt(row.created_at, 10));
    const createdAsMoment = moment(created);
    delete remainder.created_at;

    const action = row.action;
    if (!actionNames.has(action)) {
      console.log(`new action: ${action}`);
    }
    actionNames.add(action);
    delete remainder.action;

    console.log(
      `${i}: ${row.action} ${row.actor} ${row.user} ${row.org} ${
        row.repo
      } ${createdAsMoment.fromNow()}`
    );
    const auditRecord = new AuditLogRecord();
    auditRecord.recordId = '?';
    auditRecord.recordSource = AuditLogSource.AuditLogImport;

    // actor: username
    const actorUsername = row.actor ? row.actor.toLowerCase() : null;
    delete remainder.actor;
    let actorId = null;
    if (actorUsername) {
      actorId = await tryGetUserId(operations, usernamesToId, actorUsername);
    }

    // user
    const userUsername = row.user ? row.user.toLowerCase() : null;
    delete remainder.user;
    let userId = null;
    if (userUsername) {
      userId = await tryGetUserId(operations, usernamesToId, userUsername);
    }

    // org
    const orgName = row.org ? row.org.toLowerCase() : null;
    delete remainder.org;
    let organization = operations.getOrganization(orgName);

    // repo
    let repoName = row.repo ? row.repo.toLowerCase() : null;
    delete remainder.repo;
    let repoId = null;
    if (repoName) {
      const i = repoName.indexOf('/');
      if (i >= 0) {
        const organizationSubstring = repoName.substr(0, i);
        if (organizationSubstring.toLowerCase() !== organization.name) {
          console.log(
            `different organization name, may be a fork: ${organizationSubstring}`
          );
        }
        repoName = repoName.substr(i + 1);
        repoId = await tryGetRepositoryId(
          operations,
          reposToId,
          organization,
          repoName
        );
      }
    }

    // data.team
    let teamName = row['data.team'] ? row['data.team'].toLowerCase() : null;
    delete remainder['data.team'];
    let teamId = null;
    if (teamName) {
      const i = teamName.indexOf('/');
      if (i >= 0) {
        const organizationSubstring = teamName.substr(0, i);
        if (organizationSubstring.toLowerCase() !== organization.name) {
          console.log('!!');
        }
        teamName = teamName.substr(i + 1);
        teamId = await tryGetTeamId(teamToId, organization, teamName);
      }
    }

    // data.target_login
    if (row['data.target_login']) {
      console.log(row['data.target_login']);
    }
    delete remainder['data.target_login'];

    // data.hook_id
    if (row['data.hook_id']) {
      console.log(row['data.hook_id']);
    }
    delete remainder['data.hook_id'];

    // data.events
    if (row['data.events']) {
      console.log(row['data.events']);
    }
    delete remainder['data.events'];

    // data.events_were
    if (row['data.events_were']) {
      console.log(row['data.events_were']);
    }
    delete remainder['data.events_were'];

    // data.old_user
    const oldUser = row['data.old_user']
      ? row['data.old_user'].toLowerCase()
      : null;
    let oldUserId = null;
    if (oldUser) {
      oldUserId = await tryGetUserId(operations, usernamesToId, oldUser);
    }
    delete remainder['data.old_user'];

    const remainderKeys = Object.getOwnPropertyNames(remainder);

    console.dir({
      actorUsername,
      actorId,
      userUsername,
      userId,
      orgName,
      orgId: organization.id,
      repoName,
      repoId,
      oldUser,
      oldUserId,
      teamName,
      teamId,
      remainder: remainderKeys.length ? remainder : null,
    });

    if (remainderKeys.length) {
      console.dir(remainder);
      console.log();
    }

    // recordId
    // recordSource
    // action
    // additionalData
    // repositoryId, name
    // organizationId, name
    // created
    // inserted
    // actorUsername, id
    // userUsername, Id
    // incomingUsername, Id
    // teamName, id
    // corporateId, corporateUsername
  }
}

function parseCsv(filename: string): Promise<any[]> {
  return new Promise((resolve, reject) => {
    let rows = [];
    fs.createReadStream(filename)
      .pipe(csv())
      .on('data', (row) => {
        rows.push(row);
      })
      .on('end', () => {
        resolve(rows);
      });
  });
}
