//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import request = require('request');
import { ICacheHelper } from './caching';

const DefaultCacheMinutesPerContact = 120;
const BulkCacheMinutes = 60 * 24 * 14;

const BulkCacheKey = 'cc:bulk';

export interface ICorporateContactInformation {
  openSourceContact?: string;
  primaryLegalContact?: string;
  secondaryLegalContact?: string;
  highRiskBusinessReviewer?: string;
  lowRiskBusinessReviewer?: string;
  managerUsername?: string;
  managerDisplayName?: string;
  alias?: string;
  emailAddress?: string;
}

export interface ICorporateContactProvider {
  lookupContacts(corporateUsername: string): Promise<ICorporateContactInformation>;
  getBulkCachedContacts(): Promise<Map<string, ICorporateContactInformation>>;
  setBulkCachedContacts(map: Map<string, ICorporateContactInformation>): Promise<void>;
}

export default function createCorporateContactProviderInstance(config, cacheHelper: ICacheHelper): ICorporateContactProvider {
  return new MicrosoftIdentityService(config, cacheHelper);
}

export interface IMicrosoftIdentityServiceBasics {
  aadId?: string;
  alias?: string;
  costCenterCode?: string;
  emailAddress?: string;
  functionHierarchyExecCode?: string;
  manager?: string;
  preferredName?: string;
  userPrincipalName?: string;
}

interface IMicrosoftIdentityServiceResponse extends IMicrosoftIdentityServiceBasics {
  attorney?: string,
  group?: string;
  highRiskBusiness?: string,
  immediate?: boolean;
  legal?: string,
  legalOssContact?: string;
  legalPrimaryContact?: string;
  legalSecondaryContact?: string;
  lowRiskBusiness?: string;
  maintainer?: string;
  structure?: IMicrosoftIdentityServiceBasics[];
  system?: string;
}

class MicrosoftIdentityService implements ICorporateContactProvider {
  #identityConfig: any;
  #cacheHelper: ICacheHelper;

  constructor(config: any, cacheHelper: ICacheHelper) {
    this.#identityConfig = config.identity;
    this.#cacheHelper = cacheHelper;
  }

  async lookupContacts(corporateUsername: string): Promise<ICorporateContactInformation> {
    let response: IMicrosoftIdentityServiceResponse;
    const cacheKey = `cc:${corporateUsername}`;
    if (this.#cacheHelper) {
      response = await this.#cacheHelper.getObject(cacheKey);
    }
    if (!response) {
      response = await this.callIdentityService(corporateUsername);
      if (this.#cacheHelper && response) {
        // kicks off an async operation
        this.#cacheHelper.setObjectWithExpire(cacheKey, response, DefaultCacheMinutesPerContact);
      }
    }
    if (!response) {
      return null;
    }
    let managerUsername = null, managerDisplayName = null;
    const manager = response.structure && response.structure.length ? response.structure[0] : null;
    if (manager) {
      managerDisplayName = manager.preferredName;
      managerUsername = manager.userPrincipalName;
    }
    return {
      openSourceContact: response.legalOssContact,
      primaryLegalContact: response.legalPrimaryContact,
      secondaryLegalContact: response.legalSecondaryContact,
      highRiskBusinessReviewer: response.highRiskBusiness,
      lowRiskBusinessReviewer: response.lowRiskBusiness,
      alias: response.alias,
      emailAddress: response.emailAddress,
      managerUsername,
      managerDisplayName,
    };
  }

  async getBulkCachedContacts(): Promise<Map<string, ICorporateContactInformation>> {
    let map = new Map<string, IMicrosoftIdentityServiceResponse>();
    if (!this.#cacheHelper) {
      return map;
    }
    const bulk = await this.#cacheHelper.getObject(BulkCacheKey);
    if (bulk && bulk.entities) {
      if (Array.isArray(bulk.entities)) {
        map = new Map<string, IMicrosoftIdentityServiceResponse>(bulk.entities);
      } else {
        console.warn(`Cached bulk entry ${BulkCacheKey} does not contain an array of entities`);
      }
    }
    return map;
  }

  async setBulkCachedContacts(map: Map<string, ICorporateContactInformation>): Promise<void> {
    if (!this.#cacheHelper) {
      return;
    }
    const obj = { entities: Array.from(map.entries()) };
    await this.#cacheHelper.setObjectCompressedWithExpire(BulkCacheKey, obj, BulkCacheMinutes);
    return;
  }

  private getIdentityServiceRequestOptions(endpoint: string) {
    const url = this.#identityConfig.url + endpoint;
    const authToken = 'Basic ' + Buffer.from(this.#identityConfig.pat + ':', 'utf8').toString('base64');
    const headers = {
      Authorization: authToken
    };
    return { url, headers, json: true };
  }

  callIdentityService(corporateUsername: string): Promise<IMicrosoftIdentityServiceResponse> {
    return new Promise((resolve, reject) => {
      const options = this.getIdentityServiceRequestOptions(`/user/${corporateUsername}`);
      request.get(options, (error, response, entry: IMicrosoftIdentityServiceResponse) => {
        if (response.statusCode === 404) {
          return resolve(null);
        }
        if (entry && !error && entry['error'] && entry['error']['message']) {
          error = new Error(entry['error']['message']);
        }
        return error ? reject(error) : resolve(entry);
      });
    });
  }
}
