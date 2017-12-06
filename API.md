# API

There is an initial API implementation available to help partner teams
with their scenarios.

The API requires an approved API key to be used. It is designed to be run in a web
service or application scenario where the API key can be secured. If it is necessary
to allow others to use this API outside of a web service or cloud situation, this ask
and scenario will need to be resourced and built out.

The API can work with most GitHub client libraries as long as you can change the
base path and also set headers.

## API Versioning

The following API versions are currently supported:

- "2016-12-01": supports creating repositories and also creating links
- "2017-03-08": updates the shape of link responses to `aad.id` and `corporate` is renamed `aad`

An API version must be provided as either a _header_ value or in the _query string_.

- Valid requests MUST have a supported API version present
- A request MUST provide the version in either a query string or a header parameter
- A request MAY provide the API version in a header called `api-version`
- A request MAY provide the API version in a query string parameter named `api-version`

## Headers and parameters

Please provide:

- `content-type` of `application/json`
- `authorization` header using basic auth (see below)
- `api-version` header, if providing the API version through this method

A request that provides the `api-version` as a query string parameter might look like:

`GET https://endpoint/api/people/links?api-version=-2017-03-08`

## Authorization

Send a Basic Authentication where the username is `apikey` and the password is your API token.

You can technically provide the token for the username and/or password.

### Tokens are scoped to specific API(s)

An API key may be authorized for a specific API endpoint or scope. Please verify when you
are granted API access that you have access to the endpoint that you intend to.

# User link management

Information about the list of linked users who have a corporate relationship with other accounts is available.

> These APIs require that your API key be authorized for the `links` scope

To improve responsiveness, this API uses cached data. If your service is using
the data for a batch process or job, do consider keeping your own cache of the
data instead of calling this API exhaustively while performing work.

## Get all linked users

> GET /api/people/links

### Response

HTTP

```
Status: 200 OK
Content-Type: application/json; charset=utf-8
```

Body

```
[
  {
    "github": {
      "id": 1,
      "login": "username",
      "organizations": [
        "OrganizationName1",
        "OrganizationName2
      ]
    },
    "aad": {
      "alias": "alias",
      "preferredName": "The Name",
      "userPrincipalName": "upn@domain.com",
      "id": "142-guid-123",
      "emailAddress": "email.address@domain.com"
    }
  },
  {
    "github": {
      "id": 2,
      "login": "username2",
      "organizations": [
        "OrganizationName2"
      ]
    },
    "aad": {
      "alias": "alias2",
      "preferredName": "Name Here",
      "userPrincipalName": "user2@domain.com",
      "id": "guid",
      "emailAddress": "email@domain.com"
    }
  },
  ...
]
```

## Get a specific linked user

This API will retrieve information about a specific user. The first API version to support this was `2017-03-08`.

### by GitHub username

> GET /api/people/links/github/:login

Where `login` is a GitHub username, case insensitive.

#### Response

If a link is not found for the GitHub user

```
Status: 404 Not Found
```

If a link is found

```
Status: 200 OK
```

Response body:

```
{
  "github": {
    "id": 2,
    "login": "username2",
    "organizations": [
      "OrganizationName2"
    ]
  },
  "aad": {
    "alias": "alias2",
    "preferredName": "Name Here",
    "userPrincipalName": "user2@domain.com",
    "id": "guid",
    "emailAddress": "email@domain.com"
  }
}
```

### by Azure Active Directory ID

> This API returns an array if there is at least one matching account or accounts. To support scenarios with other account types or even multiple accounts such as service accounts, it is up to your application to determine how to handle more than one account. Order is not guaranteed.

> GET /api/people/links/aad/:id

Where `id` is an AAD ID.

#### Response

If a link is not registered for this user

````
Status: 404 Not Found
```

If a link is found

```
Status: 200 OK
```

Response body:

```
[
  {
    "github": {
      "id": 2,
      "login": "username2",
      "organizations": [
        "OrganizationName2"
      ]
    },
    "aad": {
      "alias": "alias2",
      "preferredName": "Name Here",
      "userPrincipalName": "user2@domain.com",
      "id": "guid",
      "emailAddress": "email@domain.com"
    }
  }
]
```

It is most common that the array will be of length === 1.

If there are no results, instead of an HTTP 200, you will receive 404 (no empty array).

# Repository management

## Create a repo

> This API requires that your API key be authorized for the `createRepo` scope

This example uses a pure POST request plus headers for authorization:

```
POST https://endpoint/api/orgName/repos?api-version=2016-12-01

HEADERS

authorization: basic :key
content-type: application/json

BODY

{
  "name": "my-test-repo",
  "private": true,
  "ms.license": "MIT",
  "ms.approval": "ReleaseReview",
  "ms.justification": "link to release approval",
  "ms.cla-entity": "Legal Entity Name",
  "ms.cla-mail": "yourteam@domain.com",
  "ms.notify": "yourteam@domain.com",
  "ms.onBehalfOf": "alias",
  "ms.teams": {
    "pull": [
      12346,
      12348
    ],
    "push": [
      12350
    ]
    "admin": [
      12345
    ]
  }
}

OUTPUT after the call is similar to (but redacted some)

OUTPUT BODY

{
  "github": {
    "id": 2,
    "name": "test-repo-ospo-2",
    "full_name": "OrgName/test-repo-ospo-2",
    "owner": {
      "login": "OrgName",
      "id": 1,
      "avatar_url": "https://avatars.githubusercontent.com/u/1?v=3",
      "gravatar_id": "",
      "url": "https://api.github.com/users/OrgName",
      "html_url": "https://github.com/OrgName",
      "repos_url": "https://api.github.com/users/OrgName/repos",
      "events_url": "https://api.github.com/users/OrgName/events{/privacy}",
      "received_events_url": "https://api.github.com/users/OrgName/received_events",
      "type": "Organization",
      "site_admin": false
    },
    "private": true,
    "html_url": "https://github.com/OrgName/test-repo-ospo-2",
    "description": null,
    "fork": false,
    "url": "https://api.github.com/repos/OrgName/test-repo-ospo-2",
    "forks_url": "https://api.github.com/repos/OrgName/test-repo-ospo-2/forks",
    "milestones_url": "https://api.github.com/repos/OrgName/test-repo-ospo-2/milestones{/number}",
    "notifications_url": "https://api.github.com/repos/OrgName/test-repo-ospo-2/notifications{?since,all,participating}",
    "labels_url": "https://api.github.com/repos/OrgName/test-repo-ospo-2/labels{/name}",
    "releases_url": "https://api.github.com/repos/OrgName/test-repo-ospo-2/releases{/id}",
    "deployments_url": "https://api.github.com/repos/OrgName/test-repo-ospo-2/deployments",
    "created_at": "2016-12-14T22:01:04Z",
    "updated_at": "2016-12-14T22:01:04Z",
    "pushed_at": "2016-12-14T22:01:05Z",
    "git_url": "git://github.com/OrgName/test-repo-ospo-2.git",
    "ssh_url": "git@github.com:OrgName/test-repo-ospo-2.git",
    "clone_url": "https://github.com/OrgName/test-repo-ospo-2.git",
    "svn_url": "https://github.com/OrgName/test-repo-ospo-2",
    "homepage": null,
    "size": 0,
    "stargazers_count": 0,
    "watchers_count": 0,
    "language": null,
    "has_issues": true,
    "has_downloads": true,
    "has_wiki": true,
    "has_pages": false,
    "forks_count": 0,
    "mirror_url": null,
    "open_issues_count": 0,
    "forks": 0,
    "open_issues": 0,
    "watchers": 0,
    "default_branch": "master",
    "permissions": {
      "admin": true,
      "push": true,
      "pull": true
    },
    "organization": {
      "login": "OrgName",
      "id": 1,
      "avatar_url": "https://avatars.githubusercontent.com/u/1?v=3",
      "gravatar_id": "",
      "url": "https://api.github.com/users/OrgName",
      "html_url": "https://github.com/OrgName",
      "events_url": "https://api.github.com/users/OrgName/events{/privacy}",
      "received_events_url": "https://api.github.com/users/OrgName/received_events",
      "type": "Organization",
      "site_admin": false
    },
    "network_count": 0,
    "subscribers_count": 3,
    "meta": {
      "x-ratelimit-limit": "62500",
      "x-ratelimit-remaining": "59992",
      "x-ratelimit-reset": "1481754433",
      "x-oauth-scopes": "repo, delete_repo, admin:org, admin:org_hook",
      "x-github-request-id": "ABC",
      "location": "https://api.github.com/repos/OrgName/test-repo-ospo-2",
      "etag": "\"3f68722071b86e49c8e25f1b76e61a32\"",
      "status": "201 Created",
      "statusActual": 201
    }
  },
  "name": "test-repo-ospo-2",
  "tasks": [
    {
      "message": "Successfully added the \"test-repo-ospo-2\" repo to GitHub team ID \"2\" with permission level PUSH."
    }
  ],
  "notified": [
    "alias@domain.com"
  ]
}

```

This example uses headers on top of a standard GitHub client:

```
POST https://endpoint/api/Microsoft/repos

HEADERS

content-type: application/json
api-version: 2016-12-01
authorization: Basic :key
ms-license: MIT
ms-approval: SmallLibrariesToolsSamples
ms-cla-entity: Legal Entity Name
ms-cla-mail: yourteam@domain.com
ms-notify: yourteam@domain.com
ms-onbehalfof: alias

BODY

{
  "name": "my-test-repo",
  "private": true
}

```

Bare minimum GitHub body component, with the type JSON, is the `name` field. You can see the GitHub API documentation here: https://developer.github.com/v3/repos/#create

- name (name of the repo)
- private (true/false)

> Note: GitHub has an input field called `team_id`. This gives _read_ access to a team ID. Our API is more sophisticated and useful since it can also assign teams to the repo with various permissions. We do not recommend providing `team_id` as a result.

API Version:
  - api-version should be in the header or query string; at this time only 2016_12_01 is supported

Casing:
  - At this time, casing is important for values

Team permissions must be set at create time as well. The API will support up to 12 team permissions plus an everyone read team permission if wanted. This design allows for specifying teams as headers. If you are setting a header, you may set it to a JSON stringified object representing the needed value. If you are setting in the body, please just provide the rich object value. You need to provide team IDs, not team names, at this time.

  - ms.teams (or ms-teams and JSON stringified object for header)

Team permission (ms.teams) value:

```
{
  "pull": [1],
  "push": [],
  "admin": [2, 3]
}
```

Always try and provide a minimum number of administrator teams, same goes for write teams (push), and encourage the standard Git workflow.