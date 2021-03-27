//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IProviders } from '../transitional';
import { Repository } from '../business/repository';

// TODO: refresh occassionally.

const RepoSocialImagesCacheKey = 'repos:socialmediaimages';

export default class PublicReposFastFilter {
  #providers: IProviders;
  #initialized: boolean;
  #socialMediaImages: Map<number, string>;

  repositories: Repository[];

  get isInitialized() {
    return this.#initialized;
  }

  constructor(providers: IProviders) {
    this.#providers = providers;
  }

  tryGetSocialMediaImage(repository: Repository) {
    const num = Number(repository.id);
    return this.#socialMediaImages ? this.#socialMediaImages.get(num) : null;
  }

  async initialize() {
    if (!this.#providers.queryCache) {
      throw new Error('Query cache provider must be available');
    }
    if (!this.#providers.queryCache.supportsRepositories) {
      throw new Error('Query cache of repositories must be available');
    }
    const { queryCache, cacheProvider } = this.#providers;

    try {
      const socialMediaImagesValue = await cacheProvider.getCompressed(
        RepoSocialImagesCacheKey
      );
      if (socialMediaImagesValue) {
        const parsed = JSON.parse(socialMediaImagesValue);
        this.#socialMediaImages = new Map(parsed);
      }
    } catch (ignoreError) {
      console.error(ignoreError);
    }

    const repositories = (await queryCache.allRepositories()).filter(
      (repo) => !repo.repository.private
    );
    this.repositories = repositories.map((entry) => entry.repository);

    this.#initialized = true;
  }
}
