//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { PostgresLinkProvider } from './postgresLinkProvider';

export default function createPostgresProvider(providers, config) {
  if (!providers.postgresPool) {
    throw new Error(
      'A Postgres pool must be initialized and available at providers.postgresPool to use the PostgresLinkProvider'
    );
  }

  if (!config.github.links.postgres) {
    throw new Error(
      'Missing configuration in the "config.github.links.postgres" block'
    );
  }

  const options = config.github.links.postgres || {};

  if (!options.tableName) {
    throw new Error(
      'Missing Postgres table name for links (REPOS_POSTGRES_LINKS_TABLE_NAME)'
    );
  }
  if (!options.githubThirdPartyName) {
    throw new Error(
      'Missing third-party designator value for GitHub (REPOS_LINKS_POSTGRES_GITHUB_THIRD_PARTY_NAME)'
    );
  }

  return new PostgresLinkProvider(providers, options);
}
