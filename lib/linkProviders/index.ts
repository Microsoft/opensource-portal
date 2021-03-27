//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IProviders } from '../../transitional';
import {
  ICorporateLink,
  ICorporateLinkExtended,
} from '../../business/corporateLink';

import createMemoryProvider from './memory';
import createPostgresProvider from './postgres';
import createTableProvider from './table';

const linkProviders = {
  memory: createMemoryProvider,
  postgres: createPostgresProvider,
  table: createTableProvider,
};

const defaultProviderName = 'memory';

// TODO: implement a new queryByThirdPartyIds to retrieve a set of links for a set of known IDs more effectively

// should move out of the postgres-specific page...
export interface ILinkProvider {
  initialize(): Promise<ILinkProvider>;

  thirdPartyType: string;

  getByThirdPartyUsername(username: string): Promise<ICorporateLink>;
  getByThirdPartyId(id: string): Promise<ICorporateLink>;
  queryByCorporateId(id: string): Promise<ICorporateLink[]>;
  queryByCorporateUsername(username: string): Promise<ICorporateLink[]>;
  getAll(): Promise<ICorporateLink[]>;
  getAllCorporateIds(): Promise<string[]>;

  createLink(link: ICorporateLink): Promise<string>;
  updateLink(linkInstance: ICorporateLink): Promise<void>;
  deleteLink(linkInstance: ICorporateLink): Promise<void>;

  dehydrateLink(linkInstance: ICorporateLinkExtended): any;
  rehydrateLink(jsonObject: any): ICorporateLink;
  dehydrateLinks(linkInstances: ICorporateLink[]): any[];
  rehydrateLinks(jsonArray: any): ICorporateLink[];
  serializationIdentifierVersion: string;
}

export interface ILinkProviderCreateOptions {
  providers: IProviders;
  config: any;
  overrideProviderType?: string;
}

export async function createAndInitializeLinkProviderInstance(
  providers,
  config,
  overrideProviderType?: string
): Promise<ILinkProvider> {
  const linkProviderOptions: ILinkProviderCreateOptions = {
    providers,
    config,
  };
  if (overrideProviderType) {
    linkProviderOptions.overrideProviderType = overrideProviderType;
  }
  const provider = createLinkProviderInstance(linkProviderOptions);
  await provider.initialize();
  return provider;
}

export function createLinkProviderInstance(
  linkProviderCreateOptions: ILinkProviderCreateOptions
): ILinkProvider {
  const config = linkProviderCreateOptions.config;
  const providers = linkProviderCreateOptions.providers;
  const provider =
    linkProviderCreateOptions.overrideProviderType ||
    config.github.links.provider.name ||
    defaultProviderName;
  // FUTURE: should also include a parameter for "what kind of third-party", i.e. 'github' to create
  let found = false;
  const providerNames = Object.getOwnPropertyNames(linkProviders);
  for (const supportedProvider of providerNames) {
    if (supportedProvider === provider) {
      found = true;
      let providerInstance = null;
      try {
        const createInstance = linkProviders[provider];
        providerInstance = createInstance(providers, config);
      } catch (createError) {
        throw createError;
      }
      return providerInstance;
    }
  }
  if (found === false) {
    throw new Error(
      `The link provider "${provider}" is not implemented or configured at this time.`
    );
  }
}
