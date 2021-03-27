//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IAuditLogRecordProvider,
  IAuditLogRecordProviderCreateOptions,
  AuditLogRecordProvider,
} from './auditLogRecordProvider';

export async function createAndInitializeAuditLogRecordProviderInstance(
  options?: IAuditLogRecordProviderCreateOptions
): Promise<IAuditLogRecordProvider> {
  const provider = new AuditLogRecordProvider(options);
  await provider.initialize();
  return provider;
}

export enum AuditLogSource {
  Webhook = 'webhook',
  AuditLogImport = 'import',
}

export const AuditEvents = {
  Repository: {
    Create: 'repo.create',
    AccessChange: 'repo.access',
    AddMember: 'repo.add_member',
    UpdateMember: 'repo.update_member',
    EditMember: 'member.edited',
    RemoveMember: 'repo.remove_member',
    Destroy: 'repo.destroy',
    Rename: 'repo.rename',
    Transfer: 'repo.transfer',
  },
  Team: {
    Create: 'team.create',
    AddRepository: 'team.add_repository',
    RemoveRepository: 'team.remove_repository',
    AddMember: 'team.add_member',
    RemoveMember: 'team.remove_member',
    Destroy: 'team.destroy',
    Edited: 'team.edited',
    UpdatePermission: 'team.update_permission', // is this changing an individual collaborator?
    UpdateRepositoryPermission: 'team.update_repository_permission',
  },
  Organization: {
    InviteMember: 'org.invite_member',
    UpdateMember: 'org.update_member',
    RemoveMember: 'org.remove_member',
    RemoveOutsideCollaborator: 'org.remove_outside_collaborator',
    CancelInvitation: 'org.cancel_invitation',
  },
};

export const MapWebhookEventsToAuditEvents = {
  'team.added_to_repository': AuditEvents.Team.AddRepository,
  'team.edited': AuditEvents.Team.Edited,
  'membership.added': AuditEvents.Team.AddMember,
  'membership.removed': AuditEvents.Team.RemoveMember,
  'member.added': AuditEvents.Repository.AddMember,
  'member.removed': AuditEvents.Repository.RemoveMember,
  'member.edited': AuditEvents.Repository.EditMember,
  'team.created': AuditEvents.Team.Create,
};

// Events not in this index:
// ---
// billing.change_email
// account.plan_change
// payment_method.create
// hook.create
// hook.config_changed
// hook.destroy
// hook.events_changed
// payment_method.update
// org.audit_log_export
// org.block_user
// org.unblock_user
// protected_branch.create
// oauth_application.destroy
// oauth_application.create
// protected_branch.rejected_ref_update
// protected_branch.destroy
// protected_branch.required_status_override
// required_status_check.create
// issue_comment.update
// issue_comment.destroy
// repo.pages_https_redirect_disabled
// repo.pages_cname
// org.add_billing_manager
// repo.pages_source
// migration.download
// migration.create
// migration.destroy_file
