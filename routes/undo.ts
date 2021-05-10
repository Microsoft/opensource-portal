//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { Operations, Repository } from '../business';
import { ErrorHelper, getProviders } from '../transitional';
import { AuditLogRecord } from '../entities/auditLogRecord/auditLogRecord';
import { daysInMilliseconds } from '../utils';
import { AuditEvents } from '../entities/auditLogRecord';
import { GitHubRepositoryPermission } from '../entities/repositoryMetadata/repositoryMetadata';
import { IGitHubIdentity, IndividualContext } from '../user';
import { IMail } from '../lib/mailProvider';
import { ReposAppRequest, UserAlertType } from '../interfaces';

const validDaysBeforeNow = 21;

interface IHaveUndoCandiates extends ReposAppRequest {
  undoOperations: IUndoEntry[];
}

interface IUndoOutcome {
  message?: string;
  warnings?: string[];
  localUrl?: string;
}

interface IUndoEntry {
  title?: string;
  operation: AuditLogRecord;
  supported: boolean;
  notSupportedMessage?: string;
  eventDescription?: string;
  undoDescription?: string;
  undo?: () => Promise<IUndoOutcome>;
}

function filterToRecentEvents(eventList: AuditLogRecord[], now: Date, validAfter: Date) {
  return eventList.filter(entry => {
    if (!entry.created) {
      return false;
    }
    const created = new Date(entry.created);
    if (created > now) {
      return false;
    }
    if (created < validAfter) {
      return false;
    }
    return true;
  });
}

function actorIdentityFromRecord(record: AuditLogRecord): string {
  if (record.actorUsername && record.actorCorporateUsername) {
    return `GitHub account ${record.actorUsername} linked to corporate user ${record.actorCorporateUsername}`;
  } else if (record.actorUsername) {
    return `GitHub account ${record.actorUsername}`;
  } else {
    return 'Unknown account';
  }
}

function repositoryIdentityFromRecord(record: AuditLogRecord): string {
  if (record.organizationName && record.repositoryName) {
    return `GitHub repo ${record.repositoryName} in the GitHub organization ${record.organizationName}`;
  } else if (record.repositoryName) {
    return `GitHub repo ${record.repositoryName}`;
  } else {
    return 'Unknown repo';
  }
}

function teamIdentityFromRecord(record: AuditLogRecord): string {
  if (record.organizationName && record.teamName) {
    return `GitHub team ${record.teamName} in the GitHub organization ${record.organizationName}`;
  } else if (record.teamName) {
    return `GitHub team ${record.teamName}`;
  } else {
    return 'unknown team name';
  }
}

function finalizeUndoEvents(operations: Operations, identity: IGitHubIdentity, records: AuditLogRecord[]): IUndoEntry[] {
  return records.map(record => {
    const entry: IUndoEntry = {
      operation: record,
      supported: false,
    };
    switch(record.action) {
      // Dropped collaborator admin on a repository
      case AuditEvents.Repository.EditMember: {
        entry.title = 'Repository permission edited';
        if (record.additionalData?.changes?.permission?.from === 'admin') {
          entry.title = 'Repo collaborator admin permission removed';
          entry.supported = true;
          entry.eventDescription = `${actorIdentityFromRecord(record)} edited their own permission to the ${repositoryIdentityFromRecord(record)}, removing their ability to administer settings and configure the repo.`;
          entry.undoDescription = 'Undo will restore the ability of the GitHub user to administer the repository.';
          entry.undo = undoRepoCollaboratorAdminRepoPermission.bind(null, operations, identity, entry);
        }
        break;
      }
      // Dropped own membership in a repo while an admin
      case AuditEvents.Repository.RemoveMember:
      case 'member.removed': {
        entry.title = 'Removed collaborator from repo';
        entry.title = 'Repo collaborator removed';
        entry.supported = true;
        entry.eventDescription = `${actorIdentityFromRecord(record)} removed themselves from the ${repositoryIdentityFromRecord(record)} while an admin, removing their ability to administer settings and configure the repo.`;
        entry.undoDescription = 'Undo will restore the ability of the GitHub user to administer the repository.';
        entry.undo = undoRepoCollaboratorAdminRepoPermission.bind(null, operations, identity, entry);
        break;
      };
      case AuditEvents.Team.Edited: {
        entry.title = 'Team permission edited'; // They dropped admin on a repository the team supports
        if (record.additionalData?.changes?.repository?.permissions?.from?.admin === true) {
          entry.title = 'Repo admin permission removed from a team';
          entry.supported = true;
          entry.eventDescription = `${actorIdentityFromRecord(record)} edited the ${teamIdentityFromRecord(record)} permission, removing the team's ability to administer settings and configure the repo.`;
          entry.undoDescription = 'Undo will restore the ability of the team to administer the repository.';
          entry.undo = undoTeamAdminRepoPermission.bind(null, operations, entry);
        }
        break;
      }
      default: {
        break;
      }
    }
    return entry;
  });
}

function undoRepoCollaboratorAdminRepoPermission(operations: Operations, identity: IGitHubIdentity, entry: IUndoEntry) {
  return undoRepoCollaboratorAdminRepoPermissionAsync(operations, identity, entry);
}

async function undoRepoCollaboratorAdminRepoPermissionAsync(operations: Operations, identity: IGitHubIdentity, entry: IUndoEntry): Promise<IUndoOutcome> {
  const operation = entry.operation;
  if (!operation) {
    throw new Error('No operation');
  }
  if (operation.action !== AuditEvents.Repository.EditMember && operation.action !== 'member.removed' && operation.action !== AuditEvents.Repository.RemoveMember) {
    throw new Error('Unsupported action');
  }
  if (operation.action === AuditEvents.Repository.EditMember && operation.additionalData?.changes?.permission?.from !== 'admin') {
    throw new Error('This action record is not of the correct format for restoring administrative collaborator permissions');
  }
  if (!operation.organizationId) {
    throw new Error('No organization ID stored in the record');
  }
  const organization = operations.getOrganizationById(Number(operation.organizationId));
  if (!operation.repositoryId) {
    throw new Error('No repository ID');
  }
  let repository: Repository = null;
  try {
    repository = await organization.getRepositoryById(Number(operation.repositoryId));
  } catch (getRepositoryError) {
    if (ErrorHelper.IsNotFound(getRepositoryError)) {
      throw new Error(`The repository ${operation.repositoryName} could not be retrieved by ID ${operation.repositoryId}. Has this repository been deleted?`);
    }
    throw getRepositoryError;
  }
  // Documenting decision:
  // Users sometimes are forcefully removed when they go on leave, have other
  // account issues, etc. By allowing recovery to _different_ user accounts,
  // we can reduce the friction on this operation. We should still warn...
  const warnings = [];
  if (operation.userId != /* loose */ identity.id || operation.userUsername.toLowerCase() !== identity.username.toLowerCase()) {
    warnings.push(`The authentication GitHub user identity, login=${identity.username} and id=${identity.id}, is different from the previous operation user, login=${operation.userUsername} id=${operation.userId}.`);
  }
  // Restore the permission
  try {
    await repository.addCollaborator(identity.username, GitHubRepositoryPermission.Admin);
  } catch (restoreError) {
    throw restoreError;
  }
  return {
    message: `The repo ${repository.name} in the ${organization.name} GitHub org has re-authorized ${identity.username} to administer the repository.`,
    localUrl: `${repository.baseUrl}permissions`,
    warnings,
  };
}

function undoTeamAdminRepoPermission(operations: Operations, entry: IUndoEntry) {
  return undoTeamAdminRepoPermissionAsync(operations, entry);
}

async function undoTeamAdminRepoPermissionAsync(operations: Operations, entry: IUndoEntry): Promise<IUndoOutcome> {
  const operation = entry.operation;
  if (!operation) {
    throw new Error('No operation');
  }
  if (operation.action !== AuditEvents.Team.Edited) {
    throw new Error('Unsupported action');
  }
  if (operation.additionalData?.changes?.repository?.permissions?.from?.admin !== true) {
    throw new Error('This action record is not of the correct format for restoring administrative permissions');
  }
  if (!operation.organizationId) {
    throw new Error('No organization ID stored in the record');
  }
  const organization = operations.getOrganizationById(Number(operation.organizationId));
  if (!operation.repositoryId) {
    throw new Error('No repository ID');
  }
  let repository: Repository = null;
  try {
    repository = await organization.getRepositoryById(Number(operation.repositoryId));
  } catch (getRepositoryError) {
    if (ErrorHelper.IsNotFound(getRepositoryError)) {
      throw new Error(`The repository ${operation.repositoryName} could not be retrieved by ID ${operation.repositoryId}. Has this repository been deleted?`);
    }
    throw getRepositoryError;
  }
  if (!operation.teamId) {
    throw new Error('No team ID');
  }
  const teamId = Number(operation.teamId);
  const team = organization.team(teamId);
  try {
    await team.getDetails();
  } catch (getTeamError) {
    if (ErrorHelper.IsNotFound(getTeamError)) {
      throw new Error(`The team ${operation.teamName} could not be retrieved by ID ${operation.teamId}. Has the team been deleted?`);
    }
    throw getTeamError;
  }
  // Restore the permission
  try {
    await repository.setTeamPermission(teamId, GitHubRepositoryPermission.Admin);
  } catch (restoreError) {
    throw restoreError;
  }
  return {
    message: `The team ${team.name} in the ${organization.name} GitHub org has been given back administrative access to the repository ${repository.full_name}.`,
    localUrl: `${repository.baseUrl}permissions`,
  };
}

router.use(asyncHandler(async function (req: IHaveUndoCandiates, res, next) {
  const { operations } = getProviders(req);
  if (!operations.allowUndoSystem) {
    res.status(404);
    return next(new Error('This feature is unavailable in this application instance'));
  }
  const auditLogRecordProvider = operations.providers.auditLogRecordProvider;
  if (!auditLogRecordProvider) {
    return next(new Error('Undo capability is not available for this site'));
  }
  const ghi = req.individualContext.getGitHubIdentity();
  if (!ghi || !ghi.username) {
    return next(new Error('GitHub identity required'));
  }
  try {
    const now = new Date();
    const before = new Date(now.getTime() - daysInMilliseconds(validDaysBeforeNow));
    const undoResults = await auditLogRecordProvider.queryAuditLogForThirdPartyIdUndoOperations(ghi.id.toString());
    const candidateUndoOperations = filterToRecentEvents(undoResults, now, before);
    req.undoOperations = finalizeUndoEvents(operations, ghi, candidateUndoOperations);
  } catch (error) {
    return next(error);
  }
  return next();
}));

router.post('/', asyncHandler(async (req: IHaveUndoCandiates, res, next) => {
  const { operations } = getProviders(req);
  const insights = operations.insights;
  const link = req.individualContext.link;
  const githubId = req.individualContext.getGitHubIdentity().id;
  const undoOperations = req.undoOperations;
  const recordId = req.body.id;
  if (!recordId) {
    res.status(400);
    return next(new Error('Missing event ID'));
  }
  const matchingEvents = undoOperations.filter(op => op.operation?.recordId === recordId);
  if (!matchingEvents.length) {
    res.status(400);
    return next(new Error('Not a valid candidate event'));
  }
  const record = matchingEvents[0];
  const operation = record.operation;
  let isOK = false;
  if (operation.userCorporateId) {
    if (operation.userCorporateId !== link.corporateId) {
      res.status(400);
      return next(new Error('This event cannot be undone for your account due to its linked corporate ID'));
    }
    isOK = true;
  } else if (operation.actorId) {
    if (operation.actorId !== githubId) {
      res.status(400);
      return next(new Error('This event cannot be undone for your account due to its GitHub account ID'));
    }
    isOK = true;
  }
  if (!isOK) {
    res.status(400);
    return next(new Error('Processing cannot continue'));
  }
  if (!record.supported) {
    return next(new Error(`This action type (${operation.action}) is not currently supported for undo operations. Record ID: ${operation.recordId}`));
  }
  if (!record.undo) {
    return next(new Error('This operation is not currently supported for undo'));
  }
  try {
    const result = await record.undo();
    req.individualContext.webContext.saveUserAlert(result.message || 'OK', 'Undo operation completed', UserAlertType.Success);
    if (result.warnings && result.warnings.length) {
      req.individualContext.webContext.saveUserAlert(result.warnings.join('; '), 'Operation warnings', UserAlertType.Warning);
    }
    insights?.trackMetric({ name: 'UndoOperations', value: 1 });
    nextTickAsyncSendMail(operations, req.individualContext, record, result);
    return res.redirect(result.localUrl || '/');
  } catch (undoError) {
    insights?.trackException({ exception: undoError });
    return next(undoError);
  }
}));

router.get('/', asyncHandler(async (req: IHaveUndoCandiates, res, next) => {
  const { operations } = getProviders(req);
  const insights = operations.insights;
  insights?.trackMetric({ name: 'UndoPageViews', value: 1 });
  return req.individualContext.webContext.render({
    view: 'undo',
    title: 'Undo',
    state: {
      undoOperations: req.undoOperations,
    },
  });
}));

function nextTickAsyncSendMail(operations: Operations, context: IndividualContext, undoEntry: IUndoEntry, undoOutcome: IUndoOutcome) {
  const insights = operations.insights;
  process.nextTick(() => {
    sendUndoMailNotification(operations, context, undoEntry, undoOutcome).then(ok => {
      insights?.trackEvent({ name: 'UndoMailSent', properties: { recordId: undoEntry.operation.recordId } });
    }).catch(error => {
      insights?.trackException({ exception: error });
      insights?.trackEvent({ name: 'UndoMailSendFailed', properties: { recordId: undoEntry.operation.recordId } });
    });
  });
}

async function sendUndoMailNotification(operations: Operations, context: IndividualContext, undoEntry: IUndoEntry, undoOutcome: IUndoOutcome) {
  const operationsMails = [ operations.getOperationsMailAddress() ];
  const ghi = context.getGitHubIdentity();
  const link = context.link;
  const details = {
    thirdPartyUsername: ghi.username,
    undoEntry,
    undoOutcome,
    operation: undoEntry.operation,
    link,
    mailAddress: null,
  };
  if (operationsMails) {
    try {
      const mailToOperations: IMail = {
        to: operationsMails,
        subject: `GitHub undo operation completed for ${ghi.username}`,
        content: await operations.emailRender('undo', {
          reason: (`A user just used the undo function on the site. As the operations contact for this system, you are receiving this e-mail.
                    This mail was sent to: ${operationsMails.join(', ')}`),
          headline: 'Undo operation complete',
          notification: 'information',
          app: `${operations.config.brand.companyName} GitHub`,
          isMailToOperations: true,
          isMailToUser: false,
          details,
        }),
      };
      await operations.sendMail(mailToOperations);
    } catch (mailIssue) {
      console.dir(mailIssue);
    }
  }
  if (!link) {
    throw new Error('No link for the individual context, no mail can be sent')
  }
  try {
    const mailAddress = await operations.getMailAddressFromCorporateUsername(link.corporateUsername);
    if (mailAddress) {
      details.mailAddress = mailAddress;
      const companyName = operations.config.brand.companyName;
      const mailToCreator: IMail = {
        to: mailAddress,
        subject: `GitHub undo operation completed for ${mailAddress}`,
        content: await operations.emailRender('undo', {
          reason: (`You just used the undo feature on the GitHub management site. This mail confirms the operation.
                    This mail was sent to: ${mailAddress} and also the GitHub administrators for the system.`),
          headline: 'Undo',
          notification: 'information',
          app: `${companyName} GitHub`,
          isMailToUser: true,
          isMailToOperations: false,
          details,
          operationsMail: operationsMails.join(','),
        }),
      };
      await operations.sendMail(mailToCreator);
    }
  } catch (noLinkOrEmail) {
    console.dir(noLinkOrEmail);
  }
}

export default router;
