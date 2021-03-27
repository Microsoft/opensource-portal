//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import querystring from 'querystring';

import { Repository } from './repository';
import {
  IPersonalizedUserAggregateRepositoryPermission,
  GraphManager,
} from './graphManager';
import { IRequestTeamPermissions } from '../middleware/github/teamPermissions';
import {
  GitHubRepositoryPermission,
  RepositoryMetadataEntity,
  RepositoryLockdownState,
} from '../entities/repositoryMetadata/repositoryMetadata';
import { IRepositoryMetadataProvider } from '../entities/repositoryMetadata/repositoryMetadataProvider';
import { TeamRepositoryPermission } from './teamRepositoryPermission';

const defaultPageSize = 20; // GitHub.com seems to use a value around 33

export interface IRepositorySearchOptions {
  pageSize?: number;
  phrase?: string;
  type?: string;
  language?: string;
  userRepos?: IPersonalizedUserAggregateRepositoryPermission[];
  teamsType?: string; // ?
  teamsSubType?: string; // ?
  specificTeamRepos?: TeamRepositoryPermission[];
  specificTeamPermissions?: IRequestTeamPermissions;
  graphManager?: GraphManager;
  repositoryMetadataProvider?: IRepositoryMetadataProvider;
  createdSince?: Date;
  metadataType?: string;
}

export class RepositorySearch {
  repos: Repository[];

  pageSize: number;
  observedLanguages: Set<string>;
  observedLanguagesEncoded: Map<string, string>;
  phrase: string;
  type: string;
  language: string;

  metadataType: string;

  createdSince?: Date;

  teamsType: string;
  teamsSubType: string;
  // repoPermissions: any;
  userRepos: IPersonalizedUserAggregateRepositoryPermission[];

  page: number;
  tags: any;
  sort: any;

  totalPages: number;
  totalRepos: number;
  pageFirstRepo: number;
  pageLastRepo: number;

  repositoryMetadataProvider: IRepositoryMetadataProvider;

  private specificTeamRepos: TeamRepositoryPermission[];

  constructor(repos: Repository[], options: IRepositorySearchOptions) {
    options = options || {};
    this.repos = repos; // is repoStore in opensource.microsoft.com, this is different by design
    this.pageSize = options.pageSize || defaultPageSize;

    this.observedLanguages = new Set();
    this.observedLanguagesEncoded = new Map();

    this.phrase = options.phrase;
    this.type = options.type;
    this.language = options.language;

    this.repositoryMetadataProvider = options.repositoryMetadataProvider;
    this.metadataType = options.metadataType;

    this.createdSince = options.createdSince;

    if (options.specificTeamRepos) {
      this.specificTeamRepos = options.specificTeamRepos;
    }

    if (options.teamsType && options.userRepos) {
      // options.repoPermissions) {
      this.teamsType = options.teamsType;
      this.teamsSubType = options.teamsSubType;
      // this.repoPermissions = options.repoPermissions;
      this.userRepos = options.userRepos;
    }
  }

  async search(page: number, sort: string): Promise<RepositorySearch> {
    this.page = page;
    this.sort = sort ? sort.charAt(0).toUpperCase() + sort.slice(1) : 'Pushed';
    let metadataCollection = null;
    if (this.metadataType && this.repositoryMetadataProvider) {
      metadataCollection = await this.repositoryMetadataProvider.queryAllRepositoryMetadatas();
    }
    this.filterByMetadata(metadataCollection)
      .filterByCreatedSince()
      .filterBySpecificTeam(this.specificTeamRepos)
      .filterByLanguageAndRecordAllLanguages(this.language)
      .filterByType(this.type)
      .filterByPhrase(this.phrase)
      .filterByTeams(this.teamsType)
      .determinePages()
      ['sortBy' + this.sort]()
      .getPage(this.page);
    await this.expandEntitiesForkForks();
    return this;
  }

  determinePages(): RepositorySearch {
    this.totalPages = Math.ceil(this.repos.length / this.pageSize);
    this.totalRepos = this.repos.length;
    return this;
  }

  filterByCreatedSince(): RepositorySearch {
    if (!this.createdSince) {
      return this;
    }
    this.repos = this.repos.filter((repo) => {
      const createdAt = new Date(repo.created_at);
      return createdAt >= this.createdSince;
    });
    return this;
  }

  filterByMetadata(metadatas: RepositoryMetadataEntity[]): RepositorySearch {
    if (!metadatas || !metadatas.length) {
      return this;
    }
    const mappedMetadata = new Map<number, RepositoryMetadataEntity>();
    for (const metadata of metadatas) {
      mappedMetadata.set(Number(metadata.repositoryId), metadata);
    }
    this.repos = this.repos.filter((repo) => {
      const id = Number(repo.id);
      switch (this.metadataType) {
        case 'with-metadata': {
          return mappedMetadata.has(id);
        }
        case 'without-metadata': {
          return !mappedMetadata.has(id);
        }
        case 'administrator-locked': {
          const metadata = mappedMetadata.get(id);
          return (
            metadata &&
            metadata.lockdownState ===
              RepositoryLockdownState.AdministratorLocked
          );
        }
        case 'locked': {
          const metadata = mappedMetadata.get(id);
          return (
            metadata &&
            metadata.lockdownState === RepositoryLockdownState.Locked
          );
        }
        case 'unlocked': {
          const metadata = mappedMetadata.get(id);
          return (
            metadata &&
            metadata.lockdownState === RepositoryLockdownState.Unlocked
          );
        }
        default: {
          throw new Error(`Unsupported metadata type ${this.metadataType}`);
        }
      }
    });
    return this;
  }

  getPage(page): RepositorySearch {
    this.repos = this.repos.slice(
      (page - 1) * this.pageSize,
      (page - 1) * this.pageSize + this.pageSize
    );
    this.pageFirstRepo = 1 + (page - 1) * this.pageSize;
    this.pageLastRepo = this.pageFirstRepo + this.repos.length - 1;
    return this;
  }

  sortByStars(): RepositorySearch {
    this.repos.sort((a, b) => {
      return b.stargazers_count - a.stargazers_count;
    });
    return this;
  }

  filterByType(type: string): RepositorySearch {
    let filter = null;
    switch (type) {
      case 'public':
        filter = (r) => {
          return r.private === false;
        };
        break;
      case 'private':
        filter = (r) => {
          return r.private === true;
        };
        break;
      case 'source':
        filter = (r) => {
          return r.fork === false;
        };
        break;
      case 'fork':
        filter = (r) => {
          return r.fork === true;
        };
        break;
    }
    if (filter) {
      this.repos = this.repos.filter(filter);
    }
    return this;
  }

  filterByPhrase(phrase: string): RepositorySearch {
    if (phrase) {
      phrase = phrase.toLowerCase();
      this.repos = this.repos.filter((r) => {
        return this.repoMatchesPhrase(r, phrase);
      });
    }
    return this;
  }

  filterBySpecificTeam(specificTeamRepos): RepositorySearch {
    if (specificTeamRepos) {
      // Also augment individual repos with permissions information
      const reposAndPermissions = new Map();
      specificTeamRepos.forEach((specificTeamAndPermission) => {
        reposAndPermissions.set(
          specificTeamAndPermission.id,
          specificTeamAndPermission.permissions
        );
      });
      this.repos = this.repos.filter((repo) => {
        const permissions = reposAndPermissions.get(repo.id);
        if (permissions) {
          // TODO: a more official flywheel attach vs adding an uninterfaced property
          repo['permissions'] = permissions;
        }
        return !!permissions;
      });
    }
    return this;
  }

  filterByTeams(teamsType: string): RepositorySearch {
    if (teamsType === 'my') {
      const userRepos = this.userRepos;
      if (!userRepos) {
        throw new Error('Missing team and repo permissions to filter by teams');
      }
      const repos = new Set<number>();
      switch (teamsType) {
        case 'my': {
          const subType = this.teamsSubType;
          userRepos.forEach((personalized) => {
            const myPermission = personalized.bestComputedPermission;
            let ok = false;
            if (
              subType === 'admin' &&
              myPermission === GitHubRepositoryPermission.Admin
            ) {
              ok = true;
            } else if (
              subType === 'write' &&
              (myPermission === 'admin' ||
                myPermission === GitHubRepositoryPermission.Push)
            ) {
              ok = true;
            } else if (subType === 'read') {
              ok = true;
            }
            if (ok) {
              repos.add(Number(personalized.repository.id));
            }
          });
          break;
        }
      }
      this.repos = this.repos.filter((repo) => {
        return repos.has(Number(repo.id));
      });
    }
    return this;
  }

  filterByLanguageAndRecordAllLanguages(language: string): RepositorySearch {
    this.repos = this.repos.filter((r) => {
      // Fill the set with all languages before filtering
      if (r.language) {
        this.observedLanguages.add(r.language);
        this.observedLanguagesEncoded.set(
          r.language,
          querystring.escape(r.language)
        );
        this.observedLanguagesEncoded.set(
          r.language.toLowerCase(),
          querystring.escape(r.language)
        );
      }
      if (!language) {
        return true;
      }
      if (r.language) {
        return r.language.toLowerCase() === language.toLowerCase();
      }
    });
    return this;
  }

  sortByForks() {
    this.repos.sort((a, b) => {
      return b.forks_count - a.forks_count;
    });
    return this;
  }

  sortBySize(): RepositorySearch {
    this.repos.sort((a, b) => {
      if (a.size > b.size) {
        return -1;
      } else if (a.size < b.size) {
        return 1;
      }
      return 0;
    });
    return this;
  }

  sortByAlphabet(): RepositorySearch {
    this.repos.sort((a, b) => {
      let nameA = a.name.toLowerCase();
      let nameB = b.name.toLowerCase();
      if (nameA < nameB) {
        return -1;
      }
      if (nameA > nameB) {
        return 1;
      }
      return 0;
    });
    return this;
  }

  async expandEntitiesForkForks(): Promise<RepositorySearch> {
    const forks = this.repos.filter((repo) => repo.fork === true);
    if (forks.length) {
      for (const fork of forks) {
        try {
          await fork.getDetails();
        } catch (ignoredError) {
          /* ignored */
        }
      }
    }
    return this;
  }

  private repoMatchesPhrase(repo: Repository, phrase: string): boolean {
    // Poor man's search, starting with just a raw includes search
    // Assumes that phrase is already lowercase to work
    let string = (
      (repo.name || '') +
      (repo.description || '') +
      (repo.id || '')
    ).toLowerCase();
    return string.includes(phrase);
  }

  private sortDates(fieldName: string, a: Repository, b: Repository): number {
    // Inverted sort (newest first)
    const aa = a[fieldName]
      ? typeof a[fieldName] === 'string'
        ? new Date(a[fieldName])
        : a[fieldName]
      : new Date(0);
    const bb = b[fieldName]
      ? typeof b[fieldName] === 'string'
        ? new Date(b[fieldName])
        : b[fieldName]
      : new Date(0);
    return aa == bb ? 0 : aa < bb ? 1 : -1;
  }

  sortByUpdated(): RepositorySearch {
    this.repos.sort(this.sortDates.bind(this, 'updated_at'));
    return this;
  }

  sortByCreated(): RepositorySearch {
    this.repos.sort(this.sortDates.bind(this, 'created_at'));
    return this;
  }

  sortByPushed(): RepositorySearch {
    this.repos.sort(this.sortDates.bind(this, 'pushed_at'));
    return this;
  }

  filterPublic(publicOnly: boolean): RepositorySearch {
    if (publicOnly) {
      this.repos = this.repos.filter((r) => {
        return !r.private;
      });
    }
    return this;
  }
}
