//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import _ from 'lodash';

import { v4 as uuidV4 } from 'uuid';
import { IReposError } from '../../../transitional';
import {
  ICorporateLinkProperties,
  ICorporateLink,
  ICorporateLinkExtended,
  CorporatePropertyNames,
} from '../../../business/corporateLink';

import { CorporateMemoryLink } from './memoryLink';
import { ILinkProvider } from '..';

const defaultThirdPartyType = 'github';
const defaultPageSize = 500;

const linkProviderInstantiationTypeProperty = '_i';
const dehydratedIdentityKey = '_lpi';
const dehydratedMemoryProviderName = 'mem';
const dehydratedMemoryProviderVersion = '0';
const dehydratedMemoryProviderIdentitySeperator = '_';
const dehydratedMemoryProviderIdentity = `${dehydratedMemoryProviderName}${dehydratedMemoryProviderIdentitySeperator}${dehydratedMemoryProviderVersion}`;

enum LinkInstantiatedType {
  MemoryEntity,
  Rehydrated,
}

interface IAlreadyLinkedError extends IReposError {
  alreadyLinked?: boolean;
}

export interface IMemoryLinkProperties extends ICorporateLinkProperties {
  // these have more explicit names to help track down bugs when using extended links
  memoryLinkId: string;
}

const linkInterfacePropertyMapping: IMemoryLinkProperties = {
  memoryLinkId: 'memoryLinkId',

  isServiceAccount: 'serviceAccount',
  serviceAccountMail: 'serviceAccountMail',

  corporateId: 'aadoid',
  corporateUsername: 'aadupn',
  corporateDisplayName: 'aadname',
  corporateMailAddress: 'corporateMailAddres',
  corporateAlias: 'corporateAlias',

  thirdPartyId: 'ghid',
  thirdPartyUsername: 'ghu',
  thirdPartyAvatar: 'ghavatar',
};

export class MemoryLinkProvider implements ILinkProvider {
  // maps link ID to a simple object entity - NOT an actual CorporateMemoryLink. This is to help fight bugs in the wrapped and provider-aware link design
  private _entities: Map<string, any>;

  private _options: any;
  private _thirdPartyType: string;

  public readonly propertyMapping: IMemoryLinkProperties = linkInterfacePropertyMapping;

  public readonly serializationIdentifierVersion: string = dehydratedMemoryProviderIdentity;

  constructor(providers, options) {
    if (!providers) {
      throw new Error(
        'The MemoryLinkProvider requires that available providers are passed into the constructor'
      );
    }

    options = options || {};

    const thirdPartyType = options.thirdPartyType || defaultThirdPartyType;
    if (thirdPartyType !== 'github') {
      throw new Error(
        'At this time only "github" is a supported third-party type.'
      );
    }

    this._options = options;
  }

  async initialize(): Promise<ILinkProvider> {
    this._entities = new Map();
    return (this as any) as ILinkProvider;
  }

  get thirdPartyType() {
    return this._thirdPartyType;
  }

  async getByThirdPartyUsername(
    username: string
  ): Promise<CorporateMemoryLink> {
    username = username.toLowerCase();
    return this.getSingleLinkByProperty(
      this.propertyMapping.thirdPartyUsername,
      username
    );
  }

  async getByThirdPartyId(id: string): Promise<CorporateMemoryLink> {
    if (typeof id !== 'string') {
      id = (id as any).toString();
    }
    return this.getSingleLinkByProperty(this.propertyMapping.thirdPartyId, id);
  }

  async queryByCorporateId(id: string): Promise<CorporateMemoryLink[]> {
    return this.getLinksByProperty(this.propertyMapping.corporateId, id);
  }

  async getAll(): Promise<CorporateMemoryLink[]> {
    const all = Array.from(this._entities.values());
    const sorted = _.sortBy(all, [
      this.propertyMapping.corporateUsername,
      this.propertyMapping.thirdPartyUsername,
    ]);
    const links = this.createLinkInstancesFromMemoryEntityArray(sorted);
    return links;
  }

  async getAllCorporateIds(): Promise<string[]> {
    const all = await this.getAll();
    return all.map((link) => link.corporateId);
  }

  async queryByCorporateUsername(username): Promise<CorporateMemoryLink[]> {
    username = username.toLowerCase();
    return this.getLinksByProperty(
      this.propertyMapping.corporateUsername,
      username
    );
  }

  async createLink(link: ICorporateLink): Promise<string> {
    const generatedLinkId = uuidV4();
    const initialEntity = {};
    initialEntity[linkInterfacePropertyMapping.memoryLinkId] = generatedLinkId;
    for (let linkPropertyName of CorporatePropertyNames) {
      const tableColumnName = linkInterfacePropertyMapping[linkPropertyName];
      if (!tableColumnName) {
        throw new Error(
          `Missing mapping from property ${linkPropertyName} to equivalent key`
        );
      }
      initialEntity[tableColumnName] = link[linkPropertyName];
    }
    if (this._entities.has(generatedLinkId)) {
      const error: IAlreadyLinkedError = new Error('This link already exists');
      throw error;
    }
    this._entities.set(generatedLinkId, initialEntity);
    return generatedLinkId;
  }

  async updateLink(linkInstance: ICorporateLink): Promise<void> {
    const tl = linkInstance as CorporateMemoryLink;
    const replacementEntity = tl.internal().getDirectEntity();
    const linkId = tl.id;
    if (!this._entities.has(linkId)) {
      throw new Error(`No existing entity with link ID ${linkId}`);
    }
    this._entities.set(linkId, replacementEntity);
  }

  async deleteLink(linkInstance: ICorporateLink): Promise<void> {
    // This is inefficient at this time; with the newer design centering
    // around a link ID, this has to query first.
    const tl = linkInstance as CorporateMemoryLink;
    const link = this.getSingleLinkByProperty(
      this.propertyMapping.memoryLinkId,
      tl.id
    );
    if (!link) {
      throw new Error(`No link found with ID ${tl.id}`);
    }
    this._entities.delete(tl.id);
  }

  dehydrateLink(linkInstance: ICorporateLinkExtended): any {
    // CONSIDER: check whether the current link type feels appropriate to us (PGSQL)
    const tlink = linkInstance as CorporateMemoryLink;
    const entity = tlink.internal().getDirectEntity();
    const shriveled = Object.assign({}, entity);
    shriveled[dehydratedIdentityKey] = dehydratedMemoryProviderIdentity;
    return shriveled;
  }

  rehydrateLink(jsonObject: any): ICorporateLink {
    if (!jsonObject) {
      throw new Error('No object provided to rehydrate');
    }
    const identity = jsonObject[dehydratedIdentityKey] as string;
    if (!identity) {
      throw new Error('No stored link provider identity to validate');
    }
    if (identity !== dehydratedMemoryProviderIdentity) {
      const sameProviderType = identity.startsWith(
        `${dehydratedMemoryProviderName}${dehydratedMemoryProviderIdentitySeperator}`
      );
      if (sameProviderType) {
        // Cross-version rehydration not supported
        throw new Error(
          `The hydrated link was created by the same ${dehydratedMemoryProviderName} provider, but a different version: ${identity}`
        );
      } else {
        throw new Error(
          `The hydrated link is incompatible with this runtime environment: ${identity}`
        );
      }
    }
    const clonedObject = Object.assign({}, jsonObject);
    delete clonedObject[dehydratedIdentityKey];
    const pglink = this.createLinkInstanceFromHydratedEntity(clonedObject);
    return pglink;
  }

  dehydrateLinks(linkInstances: ICorporateLink[]): any[] {
    if (!Array.isArray(linkInstances)) {
      throw new Error('linkInstances must be an array');
    }
    if (linkInstances.length > 0) {
      const first = linkInstances[0];
      if (first[linkProviderInstantiationTypeProperty] === undefined) {
        throw new Error(
          'linkInstances[0] does not appear to be a link instantiated by a provider'
        );
      }
    }
    //
    const arr: any[] = linkInstances.map(this.dehydrateLink.bind(this));
    return arr;
  }

  rehydrateLinks(jsonArray: any): ICorporateLink[] {
    if (!Array.isArray(jsonArray)) {
      throw new Error('jsonArray must be an array');
    }
    //
    const arr = jsonArray.map(this.rehydrateLink.bind(this));
    return (arr as any[]) as ICorporateLink[];
  }

  private createLinkInstancesFromMemoryEntityArray(
    rows: any[]
  ): CorporateMemoryLink[] {
    return rows.map(this.createLinkInstanceFromMemoryEntity.bind(this));
  }

  private createLinkInstanceFromMemoryEntity(row: any): CorporateMemoryLink {
    const linkInternalOptions = {
      provider: this,
    };
    const newLink = new CorporateMemoryLink(linkInternalOptions, row);
    newLink[linkProviderInstantiationTypeProperty] =
      LinkInstantiatedType.MemoryEntity; // in case this helps while debugging
    return newLink;
  }

  private createLinkInstanceFromHydratedEntity(jsonObject) {
    const linkInternalOptions = {
      provider: this,
    };
    const newLink = new CorporateMemoryLink(linkInternalOptions, jsonObject);
    newLink[linkProviderInstantiationTypeProperty] =
      LinkInstantiatedType.Rehydrated; // in case this helps while debugging
    return newLink;
  }

  private getUserEntitiesByProperty(
    entities: Map<string, CorporateMemoryLink>,
    propertyName: string,
    value: string
  ): CorporateMemoryLink[] {
    const rows: CorporateMemoryLink[] = [];
    for (const entry of entities.values()) {
      if (entry && entry[propertyName] === value) {
        rows.push(entry);
      }
    }
    return rows;
  }

  private getLinksByProperty(
    propertyName: string,
    value
  ): CorporateMemoryLink[] {
    const rows = this.getUserEntitiesByProperty(
      this._entities,
      propertyName,
      value
    );
    const links = this.createLinkInstancesFromMemoryEntityArray(rows);
    return links;
  }

  private getSingleLinkByProperty(
    propertyName: string,
    value
  ): CorporateMemoryLink {
    const rows = this.getUserEntitiesByProperty(
      this._entities,
      propertyName,
      value
    );
    if (rows.length <= 0) {
      return (false as any) as CorporateMemoryLink;
    }
    if (rows.length > 1) {
      const error: Error = new Error(
        `More than a single result were returned by the query (${rows.length})`
      );
      error['multipleResults'] = rows.length;
      throw error;
    }
    const entityRow = rows[0];
    const link = this.createLinkInstanceFromMemoryEntity(entityRow);
    return link;
  }
}
