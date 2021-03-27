//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import CosmosHelper from './cosmosHelper';

export interface ICampaignUserState {
  campaignGroupId: string;
  campaignId: string;
  corporateId: string;
  optOut?: Date;
  sent?: Date;
}

export interface ICampaignHelper {
  getState(
    corporateId: string,
    campaignGroupId: string,
    campaignId?: string
  ): Promise<ICampaignUserState>;
  optOut(corporateId: string, campaignGroupId: string): Promise<void>;
  clearOptOut(corporateId: string, campaignGroupId: string): Promise<void>;
  setSent(
    corporateId: string,
    campaignGroupId: string,
    campaignId: string
  ): Promise<void>;
  clearSent(
    corporateId: string,
    campaignGroupId: string,
    campaignId: string
  ): Promise<void>;
  //
  deleteOops(corporateId: string, campaignGroupId: string): Promise<void>;
}

export class StatefulCampaignProvider implements ICampaignHelper {
  #cosmosHelper: CosmosHelper;

  constructor(cosmosHelper: CosmosHelper) {
    this.#cosmosHelper = cosmosHelper;
  }

  async getState(
    corporateId: string,
    campaignGroupId: string,
    campaignId?: string
  ): Promise<ICampaignUserState> {
    const state: ICampaignUserState = {
      campaignGroupId,
      campaignId,
      corporateId,
    };
    const groupKey = this.key(corporateId, campaignGroupId);
    const key = this.key(corporateId, campaignGroupId, campaignId);
    try {
      const groupData = await this.#cosmosHelper.getObject(
        corporateId,
        groupKey
      );
      if (groupData && groupData.optOut) {
        state.optOut = new Date(groupData.optOut);
      }
      // XXX TEMP
      if (groupData && groupData.sent) {
        state.sent = new Date(groupData.sent);
      }
    } catch (noDataError) {
      if (noDataError && noDataError.code === 404) {
        // ok
      } else {
        console.dir(noDataError);
        throw noDataError;
      }
    }
    try {
      if (campaignId) {
        const data = await this.#cosmosHelper.getObject(corporateId, key);
        if (data && data.sent) {
          state.sent = new Date(data.sent);
        }
      }
    } catch (noDataError) {
      if (noDataError && noDataError.code === 404) {
        // ok
      } else {
        console.dir(noDataError);
        throw noDataError;
      }
    }
    return state;
  }

  async optOut(corporateId: string, campaignGroupId: string): Promise<void> {
    const value = Object.assign(this.baseObject(corporateId, campaignGroupId), {
      optOut: new Date().toISOString(),
    });
    await this.#cosmosHelper.setObject(value);
  }

  async clearOptOut(
    corporateId: string,
    campaignGroupId: string
  ): Promise<void> {
    const value = Object.assign(this.baseObject(corporateId, campaignGroupId), {
      optOut: false,
    });
    await this.#cosmosHelper.setObject(value);
  }

  async setSent(
    corporateId: string,
    campaignGroupId: string,
    campaignId: string
  ): Promise<void> {
    const value = Object.assign(
      this.baseObject(corporateId, campaignGroupId, campaignId),
      {
        sent: new Date().toISOString(),
      }
    );
    await this.#cosmosHelper.setObject(value);
  }

  async deleteOops(
    corporateId: string,
    campaignGroupId: string
  ): Promise<void> {
    const id = this.key(corporateId, campaignGroupId);
    try {
      await this.#cosmosHelper.delete(corporateId, id);
    } catch (oopsError) {
      if (oopsError && oopsError.code === 404) {
        return;
      }
      throw oopsError;
    }
  }

  async clearSent(
    corporateId: string,
    campaignGroupId: string,
    campaignId: string
  ): Promise<void> {
    const id = this.key(corporateId, campaignGroupId, campaignId);
    try {
      await this.#cosmosHelper.delete(corporateId, id);
    } catch (oopsError) {
      if (oopsError && oopsError.code === 404) {
        return;
      }
      throw oopsError;
    }
  }

  private key(
    corporateId: string,
    campaignGroupId: string,
    campaignId?: string
  ) {
    return campaignId
      ? `${campaignGroupId}-${campaignId}-${corporateId}`
      : `${campaignGroupId}-${corporateId}`;
  }

  private baseObject(
    corporateId: string,
    campaignGroupId: string,
    campaignId?: string
  ) {
    const id = this.key(corporateId, campaignGroupId, campaignId);
    return {
      id,
      corporateId,
      campaignGroupId,
      campaignId,
    };
  }
}
