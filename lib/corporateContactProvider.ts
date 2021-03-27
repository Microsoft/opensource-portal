//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import axios, { AxiosError } from 'axios';

import { CreateError } from '../transitional';
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
  lookupContacts(
    corporateUsername: string
  ): Promise<ICorporateContactInformation>;
  getBulkCachedContacts(): Promise<
    Map<string, ICorporateContactInformation | boolean>
  >;
  setBulkCachedContacts(
    map: Map<string, ICorporateContactInformation | boolean>
  ): Promise<void>;
}

export default function createCorporateContactProviderInstance(
  config,
  cacheHelper: ICacheHelper
): ICorporateContactProvider {
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

interface IMicrosoftIdentityServiceResponse
  extends IMicrosoftIdentityServiceBasics {
  attorney?: string;
  group?: string;
  highRiskBusiness?: string;
  immediate?: boolean;
  legal?: string;
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

  async lookupContacts(
    corporateUsername: string
  ): Promise<ICorporateContactInformation> {
    let response: IMicrosoftIdentityServiceResponse;
    const cacheKey = `cc:${corporateUsername}`;
    if (this.#cacheHelper) {
      try {
        response = await this.#cacheHelper.getObject(cacheKey);
      } catch (ignoreError) {
        /* ignored */
      }
    }
    if (!response) {
      response = await this.callIdentityService(corporateUsername);
      if (this.#cacheHelper && response) {
        // kicks off an async operation
        this.#cacheHelper.setObjectWithExpire(
          cacheKey,
          response,
          DefaultCacheMinutesPerContact
        );
      }
    }
    if (!response) {
      return null;
    }
    let managerUsername = null,
      managerDisplayName = null;
    const manager =
      response.structure && response.structure.length
        ? response.structure[0]
        : null;
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

  async getBulkCachedContacts(): Promise<
    Map<string, ICorporateContactInformation | boolean>
  > {
    let map = new Map<string, IMicrosoftIdentityServiceResponse | boolean>();
    if (!this.#cacheHelper) {
      return map;
    }
    const bulk = await this.#cacheHelper.getObject(BulkCacheKey);
    if (bulk && bulk.entities) {
      if (Array.isArray(bulk.entities)) {
        map = new Map<string, IMicrosoftIdentityServiceResponse>(bulk.entities);
        if (bulk.empties) {
          for (let i = 0; i < bulk.empties.length; i++) {
            map.set(bulk.empties[i], false);
          }
        }
      } else {
        console.warn(
          `Cached bulk entry ${BulkCacheKey} does not contain an array of entities`
        );
      }
    }
    return map;
  }

  async setBulkCachedContacts(
    map: Map<string, ICorporateContactInformation | boolean>
  ): Promise<void> {
    if (!this.#cacheHelper) {
      return;
    }
    const all = Array.from(map.entries());
    const entities = all.filter((e) => typeof e[1] !== 'boolean');
    const empties = all
      .filter((e) => typeof e[1] === 'boolean')
      .map((e) => e[0])
      .filter((e) => e);
    const obj = { entities, empties };
    await this.#cacheHelper.setObjectCompressedWithExpire(
      BulkCacheKey,
      obj,
      BulkCacheMinutes
    );
  }

  private getIdentityServiceRequestOptions(endpoint: string) {
    const url = this.#identityConfig.url + endpoint;
    const authToken =
      'Basic ' +
      Buffer.from(this.#identityConfig.pat + ':', 'utf8').toString('base64');
    const headers = {
      Authorization: authToken,
    };
    return { url, headers };
  }

  async callIdentityService(
    corporateUsername: string
  ): Promise<IMicrosoftIdentityServiceResponse> {
    try {
      const response = await axios(
        this.getIdentityServiceRequestOptions(`/user/${corporateUsername}`)
      );
      if (response.data.error?.message) {
        throw CreateError.InvalidParameters(response.data.error.message);
      }
      const entity = response.data as IMicrosoftIdentityServiceResponse;
      return entity;
    } catch (error) {
      const axiosError = error as AxiosError;
      if (axiosError?.response?.status === 404) {
        return null;
      } else if (axiosError?.response?.status >= 300) {
        throw CreateError.CreateStatusCodeError(
          axiosError.response.status,
          `Response code: ${axiosError.response.status}`
        );
      }
      throw error;
    }
  }
}
