//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// AUDIT LOG: this capability is offline since the original data store was not ideal.
// This should be rewritten to use the entity concept, and integrate with the newer
// GitHub Enterprise Cloud capability of using GraphQL to hit the official audit log
// for organizations, and also to import JSON-based audit export files.

import { WebhookProcessor } from '../organizationProcessor';
import { Operations } from '../../business';
import { Organization } from '../../business';
import { AuditLogRecord } from '../../entities/auditLogRecord/auditLogRecord';
import { MapWebhookEventsToAuditEvents, AuditLogSource } from '../../entities/auditLogRecord';

const eventTypes = new Set([
  'membership',
  'member',
  'organization',
  'repository',
  'team',
]);

async function runAsync(operations: Operations, organization: Organization, data: any) {
  const { auditLogRecordProvider } = operations.providers;
  if (!auditLogRecordProvider) {
    return;
  }
  const { body, properties } = data;
  const fullEventName = `${properties.event}.${body.action}`;
  const mappedEventValue = MapWebhookEventsToAuditEvents[fullEventName];
  if (!mappedEventValue) {
    console.log(`unsupported audit log event: ${fullEventName}`);
    return;
  }
  const record = new AuditLogRecord();
  record.recordSource = AuditLogSource.Webhook;
  // UUID for now: record.recordId =
  record.action = fullEventName;
  record.created = properties.started || new Date();
  record.additionalData = {};
  let undoCandidate = false;
  if (body.changes) {
    const changes = body.changes;
     record.additionalData.changes = changes;
     if (changes?.repository?.permissions?.from?.admin === true) {
      undoCandidate = true;
     } else if (changes?.permission?.from === 'admin') {
      undoCandidate = true;
     }
  }
  if (body.scope === 'team' && body.action === 'removed' && body.member) {
    undoCandidate = true;
  } else if (body.event === 'team' && body.action === 'removed_from_repository') {
    undoCandidate = true;
  } else if (body.event === 'membership' && body.action === 'removed') {
    undoCandidate = true;
  } else if (fullEventName === 'member.removed') {
    undoCandidate = true;
  }
  if (undoCandidate) {
    record.additionalData.undoCandidate = undoCandidate;
  }
  if (body.scope) {
    record.additionalData.scope = body.scope;
  }
  if (properties.delivery) {
    record.additionalData.delivery = properties.delivery;
  }
  if (body.organization) {
    record.organizationName = body.organization.login;
    record.organizationId = body.organization.id;
  }
  if (body.repository) {
    record.repositoryName = body.repository.name;
    record.repositoryId = body.repository.id;
  }
  if (body.sender) {
    record.actorId = body.sender.id;
    record.actorUsername = body.sender.login;
    // do we have a link for the actor?
  }
  if (body.user || body.member) {
    record.userId = body.user?.id || body.member?.id;
    record.userUsername = body.user?.login || body.member?.login;
    // TODO: corporate link?
  }
  if (body.team) {
    record.teamId = body.team.id;
    record.teamName = body.team.name;
  }
  if (body.membership) {
  //   document.membership = {
  //     state: body.membership.state,
  //     role: body.membership.role,
  //     user: {
  //       id: body.membership.user.id,
  //       login: body.membership.user.login,
  //     },
  }
  record.inserted = new Date();
  // console.dir(record);
  await auditLogRecordProvider.insertRecord(record);
}

export default class AuditLogRecorderWebhookProcessor implements WebhookProcessor {
  filter(data: any) {
    let eventType = data.properties.event;
    const has = eventTypes.has(eventType);
    if (!has) {
      console.log(`audit log does not support event type: ${eventType}`);
    }
    return has;
  }

  async run(operations: Operations, organization: Organization, data: any): Promise<boolean> {
    const result = await runAsync(operations, organization, data);
    return true;
  }
}
