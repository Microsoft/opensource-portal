//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// const debugCacheOptimization = require('debug')('oss-cache-optimization');

export enum FieldType {
  Keep, // keep the property value, if present
  Drop, // drop it
  Entity, // a sub-entity
}

export enum ResponseBodyType {
  Entity,
  Array,
}

const entityPropertiesToKeep = new Map<GitHubResponseType, Set<string>>();
const entityPropertiesToDrop = new Map<GitHubResponseType, Set<string>>();
const entityPropertiesSubsets = new Map<
  GitHubResponseType,
  Map<string, GitHubResponseType>
>();
const apiToEntityType = new Map<string, GitHubResponseType>();
const apiToEntityResponseType = new Map<string, ResponseBodyType>();

export enum GitHubResponseType {
  Headers = 'Headers',
  Organization = 'Organization',
  OrganizationDetails = 'OrganizationDetails',
  OrganizationUserMembership = 'OrganizationUserMembership',
  Repository = 'Repository',
  UserOrOrganization = 'UserOrOrganization',
  Collaborator = 'Collaborator',
  UserPermissionLevel = 'UserPermissionLevel',
  Team = 'Team',
  UserDetail = 'UserDetail',
}

function RegisterEntity(entityType: GitHubResponseType, properties: any[]) {
  if (!entityPropertiesToKeep.has(entityType)) {
    entityPropertiesToKeep.set(entityType, new Set());
  }
  if (!entityPropertiesToDrop.has(entityType)) {
    entityPropertiesToDrop.set(entityType, new Set());
  }
  if (!entityPropertiesSubsets.has(entityType)) {
    entityPropertiesSubsets.set(entityType, new Map());
  }
  // debugCacheOptimization(`Registering subset entity ${entityType}`);
  for (let i = 0; i < properties.length; i++) {
    let pair: [string, FieldType, GitHubResponseType?];
    pair = properties[i];
    const fieldName = pair[0];
    if (pair[1] === FieldType.Drop || pair[1] === FieldType.Keep) {
      const targetMap =
        pair[1] === FieldType.Drop
          ? entityPropertiesToDrop
          : entityPropertiesToKeep;
      const targetName = pair[1] === FieldType.Drop ? 'drop' : 'keep';
      const container = targetMap.get(entityType);
      container.add(fieldName);
      // debugCacheOptimization(`Entity ${entityType} ${targetName.toUpperCase()} Property ${fieldName}`);
    } else if (pair[1] === FieldType.Entity) {
      const mappedPropertyType = pair[2];
      if (!mappedPropertyType) {
        throw new Error(
          'Entity subtypes must have defined new subtypes to register'
        );
      }
      const container = entityPropertiesSubsets.get(entityType);
      container.set(fieldName, mappedPropertyType);
    } else {
      throw new Error(`Not supported/implemented field type ${pair[1]}`);
    }
  }
}

function RegisterEndpoint(
  endpoint: string,
  entityType: GitHubResponseType,
  responseType: ResponseBodyType = ResponseBodyType.Entity
) {
  if (apiToEntityType.has(endpoint)) {
    throw new Error(
      `Endpoint ${endpoint} has already registered a GitHub entity type. No duplicates.`
    );
  }
  apiToEntityType.set(endpoint, entityType);
  apiToEntityResponseType.set(endpoint, responseType);
}

RegisterEndpoint('repos.get', GitHubResponseType.Repository);
RegisterEndpoint(
  'repos.getForOrg',
  GitHubResponseType.Repository,
  ResponseBodyType.Array
);
RegisterEndpoint(
  'repos.listForOrg',
  GitHubResponseType.Repository,
  ResponseBodyType.Array
); // new  repos.getForOrg
RegisterEndpoint(
  'repos.getCollaborators',
  GitHubResponseType.Collaborator,
  ResponseBodyType.Array
);
RegisterEndpoint(
  'repos.listCollaborators',
  GitHubResponseType.Collaborator,
  ResponseBodyType.Array
); // new repos.getCollaborators
RegisterEndpoint(
  'repos.reviewUserPermissionLevel',
  GitHubResponseType.UserPermissionLevel
);
RegisterEndpoint(
  'repos.getCollaboratorPermissionLevel',
  GitHubResponseType.UserPermissionLevel
); // replaces repos.reviewUserPermissionLevel
RegisterEndpoint(
  'repos.listTeams',
  GitHubResponseType.Team,
  ResponseBodyType.Array
); // new, replaces ?
RegisterEndpoint('orgs.get', GitHubResponseType.OrganizationDetails);
RegisterEndpoint(
  'orgs.getMembers',
  GitHubResponseType.UserOrOrganization,
  ResponseBodyType.Array
);
RegisterEndpoint(
  'orgs.listMembers',
  GitHubResponseType.UserOrOrganization,
  ResponseBodyType.Array
); // new orgs.getMembers
RegisterEndpoint(
  'orgs.getOrgMembership',
  GitHubResponseType.OrganizationUserMembership
);
RegisterEndpoint(
  'orgs.getMembership',
  GitHubResponseType.OrganizationUserMembership
); // replaces orgs.getOrgMembership
RegisterEndpoint('orgs.getTeam', GitHubResponseType.Team);
RegisterEndpoint(
  'orgs.getTeams',
  GitHubResponseType.Team,
  ResponseBodyType.Array
);
RegisterEndpoint(
  'orgs.getTeamMembers',
  GitHubResponseType.UserOrOrganization,
  ResponseBodyType.Array
);
RegisterEndpoint(
  'orgs.getTeamRepos',
  GitHubResponseType.Repository,
  ResponseBodyType.Array
);
RegisterEndpoint('teams.list', GitHubResponseType.Team, ResponseBodyType.Array); // new orgs.getTeams
RegisterEndpoint(
  'teams.listMembersLegacy',
  GitHubResponseType.UserOrOrganization,
  ResponseBodyType.Array
); // new orgs.getTeamMembers
RegisterEndpoint(
  'teams.listMembersInOrg',
  GitHubResponseType.UserOrOrganization,
  ResponseBodyType.Array
); // new orgs.getTeamMembers
// teams.listReposInOrg
RegisterEndpoint('users.getById', GitHubResponseType.UserDetail);

RegisterEntity(GitHubResponseType.UserPermissionLevel, [
  ['permission', FieldType.Keep],
  ['user', FieldType.Entity, GitHubResponseType.UserOrOrganization],
]);

RegisterEntity(GitHubResponseType.OrganizationUserMembership, [
  ['url', FieldType.Drop],
  ['state', FieldType.Keep],
  ['role', FieldType.Keep],
  ['organization_url', FieldType.Drop],
  ['user', FieldType.Entity, GitHubResponseType.UserOrOrganization],
  ['organization', FieldType.Entity, GitHubResponseType.OrganizationDetails],
]);

RegisterEntity(GitHubResponseType.Headers, [
  ['access-control-allow-origin', FieldType.Drop],
  ['access-control-expose-headers', FieldType.Drop],
  ['cache-control', FieldType.Drop],
  ['connection', FieldType.Drop],
  ['content-encoding', FieldType.Drop],
  ['content-length', FieldType.Drop],
  ['content-security-policy', FieldType.Drop],
  ['content-type', FieldType.Keep],
  ['date', FieldType.Keep],
  ['etag', FieldType.Keep],
  ['last-modified', FieldType.Keep],
  ['link', FieldType.Keep],
  ['referrer-policy', FieldType.Drop],
  ['server', FieldType.Drop],
  ['status', FieldType.Keep],
  ['strict-transport-security', FieldType.Drop],
  ['transfer-encoding', FieldType.Drop],
  ['vary', FieldType.Drop],
  ['x-accepted-oauth-scopes', FieldType.Drop],
  ['x-content-type-options', FieldType.Drop],
  ['x-frame-options', FieldType.Drop],
  ['x-oauth-client-id', FieldType.Keep],
  ['x-github-media-type', FieldType.Keep],
  ['x-github-request-id', FieldType.Keep],
  ['x-oauth-scopes', FieldType.Drop],
  ['x-ratelimit-limit', FieldType.Drop],
  ['x-ratelimit-remaining', FieldType.Keep],
  ['x-ratelimit-reset', FieldType.Keep],
  ['x-xss-protection', FieldType.Drop],
  ['statusActual', FieldType.Keep],
]);

RegisterEntity(GitHubResponseType.Collaborator, [
  ['permissions', FieldType.Keep], // key to collaborators
  ['login', FieldType.Keep],
  ['id', FieldType.Keep],
  ['node_id', FieldType.Keep],
  ['avatar_url', FieldType.Keep],
  ['gravatar_id', FieldType.Drop],
  ['url', FieldType.Drop],
  ['html_url', FieldType.Drop],
  ['followers_url', FieldType.Drop],
  ['following_url', FieldType.Drop],
  ['gists_url', FieldType.Drop],
  ['starred_url', FieldType.Drop],
  ['subscriptions_url', FieldType.Drop],
  ['organizations_url', FieldType.Drop],
  ['repos_url', FieldType.Drop],
  ['events_url', FieldType.Drop],
  ['received_events_url', FieldType.Drop],
  ['type', FieldType.Keep],
  ['site_admin', FieldType.Drop],
]);

RegisterEntity(GitHubResponseType.Repository, [
  ['id', FieldType.Keep],
  ['node_id', FieldType.Keep],
  ['name', FieldType.Keep],
  ['full_name', FieldType.Keep],
  ['private', FieldType.Keep],
  ['owner', FieldType.Entity, GitHubResponseType.UserOrOrganization],
  ['organization', FieldType.Entity, GitHubResponseType.Organization],
  ['html_url', FieldType.Keep], // explicitly used in Repository.ts
  ['description', FieldType.Keep],
  ['fork', FieldType.Keep],
  ['url', FieldType.Keep], // ? probably want to keep for now
  ['forks_url', FieldType.Drop],
  ['keys_url', FieldType.Drop],
  ['collaborators_url', FieldType.Drop],
  ['teams_url', FieldType.Drop],
  ['hooks_url', FieldType.Drop],
  ['issue_events_url', FieldType.Drop],
  ['events_url', FieldType.Drop],
  ['assignees_url', FieldType.Drop],
  ['branches_url', FieldType.Drop],
  ['tags_url', FieldType.Drop],
  ['blobs_url', FieldType.Drop],
  ['git_tags_url', FieldType.Drop],
  ['git_refs_url', FieldType.Drop],
  ['trees_url', FieldType.Drop],
  ['statuses_url', FieldType.Drop],
  ['languages_url', FieldType.Drop],
  ['stargazers_url', FieldType.Drop],
  ['contributors_url', FieldType.Drop],
  ['subscribers_url', FieldType.Drop],
  ['subscription_url', FieldType.Drop],
  ['commits_url', FieldType.Drop],
  ['git_commits_url', FieldType.Drop],
  ['comments_url', FieldType.Drop],
  ['issue_comment_url', FieldType.Drop],
  ['contents_url', FieldType.Drop],
  ['compare_url', FieldType.Drop],
  ['merges_url', FieldType.Drop],
  ['archive_url', FieldType.Drop],
  ['downloads_url', FieldType.Drop],
  ['issues_url', FieldType.Drop],
  ['pulls_url', FieldType.Drop],
  ['milestones_url', FieldType.Drop],
  ['notifications_url', FieldType.Drop],
  ['labels_url', FieldType.Drop],
  ['releases_url', FieldType.Drop],
  ['deployments_url', FieldType.Drop],
  ['created_at', FieldType.Keep],
  ['updated_at', FieldType.Keep],
  ['pushed_at', FieldType.Keep],
  ['git_url', FieldType.Keep],
  ['ssh_url', FieldType.Keep],
  ['clone_url', FieldType.Keep],
  ['svn_url', FieldType.Drop],
  ['mirror_url', FieldType.Drop],
  ['archived', FieldType.Keep],
  ['disabled', FieldType.Keep],
  ['open_issues_count', FieldType.Keep],
  ['license', FieldType.Keep],
  ['forks', FieldType.Keep],
  ['open_issues', FieldType.Keep],
  ['watchers', FieldType.Keep],
  ['default_branch', FieldType.Keep],
  ['allow_squash_merge', FieldType.Keep],
  ['allow_merge_commit', FieldType.Keep],
  ['allow_rebase_merge', FieldType.Keep],
  ['homepage', FieldType.Keep],
  ['size', FieldType.Keep],
  ['stargazers_count', FieldType.Keep],
  ['watchers_count', FieldType.Keep],
  ['language', FieldType.Keep],
  ['has_issues', FieldType.Keep],
  ['has_projects', FieldType.Keep],
  ['has_downloads', FieldType.Keep],
  ['has_wiki', FieldType.Keep],
  ['has_pages', FieldType.Keep],
  ['forks_count', FieldType.Keep],
  ['network_count', FieldType.Keep],
  ['subscribers_count', FieldType.Keep],
  ['permissions', FieldType.Keep], // is useful when used for team repos listing
  ['parent', FieldType.Entity, GitHubResponseType.Repository],
]);

RegisterEntity(GitHubResponseType.UserDetail, [
  ['login', FieldType.Keep],
  ['id', FieldType.Keep],
  ['node_id', FieldType.Keep],
  ['avatar_url', FieldType.Keep],
  ['created_at', FieldType.Keep],
  ['updated_at', FieldType.Keep],
  ['bio', FieldType.Keep],
  ['blog', FieldType.Keep],
  ['location', FieldType.Keep],
  ['name', FieldType.Keep],
  ['company', FieldType.Keep],
  ['email', FieldType.Keep],
  ['followers', FieldType.Keep],
  ['hireable', FieldType.Keep],
  ['following', FieldType.Keep],
  ['public_gists', FieldType.Keep],
  ['public_repos', FieldType.Keep],
  ['gravatar_id', FieldType.Drop],
  ['followers_url', FieldType.Drop],
  ['following_url', FieldType.Drop],
  ['gists_url', FieldType.Drop],
  ['url', FieldType.Drop],
  ['html_url', FieldType.Drop],
  ['followers_url', FieldType.Drop],
  ['following_url', FieldType.Drop],
  ['starred_url', FieldType.Drop],
  ['subscriptions_url', FieldType.Drop],
  ['organizations_url', FieldType.Drop],
  ['repos_url', FieldType.Drop],
  ['events_url', FieldType.Drop],
  ['received_events_url', FieldType.Drop],
  ['type', FieldType.Keep],
  ['site_admin', FieldType.Drop],
]);
RegisterEntity(GitHubResponseType.UserOrOrganization, [
  ['login', FieldType.Keep],
  ['id', FieldType.Keep],
  ['node_id', FieldType.Keep],
  ['avatar_url', FieldType.Keep],
  ['gravatar_id', FieldType.Drop],
  ['url', FieldType.Keep], // ? do we ever actually use it
  ['html_url', FieldType.Drop],
  ['followers_url', FieldType.Drop],
  ['following_url', FieldType.Drop],
  ['gists_url', FieldType.Drop],
  ['starred_url', FieldType.Drop],
  ['subscriptions_url', FieldType.Drop],
  ['organizations_url', FieldType.Drop],
  ['repos_url', FieldType.Drop],
  ['events_url', FieldType.Drop],
  ['received_events_url', FieldType.Drop],
  ['type', FieldType.Keep],
  ['site_admin', FieldType.Drop],
]);

RegisterEntity(GitHubResponseType.Team, [
  ['name', FieldType.Keep],
  ['id', FieldType.Keep],
  ['node_id', FieldType.Keep],
  ['slug', FieldType.Keep],
  ['description', FieldType.Keep],
  ['privacy', FieldType.Keep],
  ['url', FieldType.Keep],
  ['html_url', FieldType.Keep],
  ['members_url', FieldType.Drop],
  ['repositories_url', FieldType.Drop],
  ['permission', FieldType.Keep],
  ['created_at', FieldType.Keep],
  ['updated_at', FieldType.Keep],
  ['members_count', FieldType.Keep],
  ['repos_count', FieldType.Keep],
  ['organization', FieldType.Entity, GitHubResponseType.OrganizationDetails],
  ['parent', FieldType.Entity, GitHubResponseType.Team],
]);

RegisterEntity(GitHubResponseType.Organization, [
  ['login', FieldType.Keep],
  ['id', FieldType.Keep],
  ['node_id', FieldType.Keep],
  ['avatar_url', FieldType.Keep],
  ['gravatar_id', FieldType.Drop],
  ['url', FieldType.Keep], // ? do we ever actually use it
  ['html_url', FieldType.Drop],
  ['hooks_url', FieldType.Drop],
  ['followers_url', FieldType.Drop],
  ['following_url', FieldType.Drop],
  ['gists_url', FieldType.Drop],
  ['starred_url', FieldType.Drop],
  ['subscriptions_url', FieldType.Drop],
  ['organizations_url', FieldType.Drop],
  ['repos_url', FieldType.Drop],
  ['events_url', FieldType.Drop],
  ['received_events_url', FieldType.Drop],
  ['type', FieldType.Keep],
  ['site_admin', FieldType.Drop],
]);

RegisterEntity(GitHubResponseType.OrganizationDetails, [
  ['login', FieldType.Keep],
  ['id', FieldType.Keep],
  ['node_id', FieldType.Keep],
  ['url', FieldType.Keep],
  ['repos_url', FieldType.Drop],
  ['events_url', FieldType.Drop],
  ['hooks_url', FieldType.Drop],
  ['issues_url', FieldType.Drop],
  ['members_url', FieldType.Drop],
  ['public_members_url', FieldType.Drop],
  ['avatar_url', FieldType.Keep],
  ['description', FieldType.Keep],
  ['name', FieldType.Keep],
  ['company', FieldType.Keep],
  ['blog', FieldType.Keep],
  ['location', FieldType.Keep],
  ['email', FieldType.Keep],
  ['is_verified', FieldType.Keep],
  ['has_organization_projects', FieldType.Keep],
  ['has_repository_projects', FieldType.Keep],
  ['public_repos', FieldType.Keep],
  ['public_gists', FieldType.Keep],
  ['followers', FieldType.Keep],
  ['following', FieldType.Keep],
  ['html_url', FieldType.Keep],
  ['created_at', FieldType.Keep],
  ['updated_at', FieldType.Keep],
  ['type', FieldType.Keep],
  ['total_private_repos', FieldType.Keep],
  ['owned_private_repos', FieldType.Keep],
  ['private_gists', FieldType.Keep],
  ['disk_usage', FieldType.Keep],
  ['collaborators', FieldType.Keep],
  ['billing_email', FieldType.Keep],
  ['default_repository_permission', FieldType.Keep],
  ['members_can_create_repositories', FieldType.Keep],
  ['two_factor_requirement_enabled', FieldType.Keep],
  ['plan', FieldType.Keep],
]);

export interface IGitHubEntityDefinitions {
  entityPropertiesToKeep: Map<GitHubResponseType, Set<string>>;
  entityPropertiesToDrop: Map<GitHubResponseType, Set<string>>;
  entityPropertiesSubsets: Map<
    GitHubResponseType,
    Map<string, GitHubResponseType>
  >;
  apiToEntityType: Map<string, GitHubResponseType>;
  apiToEntityResponseType: Map<string, ResponseBodyType>;
}

export function getEntityDefinitions(): IGitHubEntityDefinitions {
  return {
    entityPropertiesToKeep,
    entityPropertiesToDrop,
    entityPropertiesSubsets,
    apiToEntityType,
    apiToEntityResponseType,
  };
}
