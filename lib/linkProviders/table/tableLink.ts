//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  ICorporateLinkExtended,
  ICorporateLinkExtendedDirectMethods,
} from '../../../business/corporateLink';
import { TableLinkProvider } from './tableLinkProvider';

export interface IInternalTableLinkOptions {
  provider: TableLinkProvider;
}

export interface ITableLinkInstanceDataHelpers {
  delete: () => Promise<boolean>;
  update: () => Promise<boolean>;
  save: () => Promise<boolean>;
}

export interface ITableLinkInstanceInternalHelpers
  extends ICorporateLinkExtendedDirectMethods {
  acknowledgeSuccessfulUpdate: () => void;
  getDirtyColumns: () => any;
  getDirectEntity: () => any;
}

export class CorporateTableLink implements ICorporateLinkExtended {
  private _provider: TableLinkProvider;

  private _id: string;
  private _entity: any;
  private _columnUpdates: any;
  private _originalEntity: any;

  constructor(linkInternalOptions: IInternalTableLinkOptions, row: any) {
    this._provider = linkInternalOptions.provider;

    this._id = row[this._provider.propertyMapping.linkId]; // immutable
    this._entity = row;

    this._originalEntity = null;
    this._columnUpdates = {};
  }

  get id(): string {
    return this._id;
  }

  get created(): Date {
    return this._entity[this._provider.propertyMapping.created];
  }

  get corporateId(): string {
    return this._entity[this._provider.propertyMapping.corporateId];
  }

  get corporateMailAddress(): string {
    return this._entity[this._provider.propertyMapping.corporateMailAddress];
  }

  set corporateMailAddress(value: string) {
    _updateColumn(
      this,
      this._provider.propertyMapping.corporateMailAddress,
      value
    );
  }

  get corporateAlias() {
    return this._entity[this._provider.propertyMapping.corporateAlias];
  }

  set corporateAlias(value: string) {
    _updateColumn(this, this._provider.propertyMapping.corporateAlias, value);
  }

  set corporateId(value: string) {
    _updateColumn(this, this._provider.propertyMapping.corporateId, value);
  }

  get corporateUsername(): string {
    return this._entity[this._provider.propertyMapping.corporateUsername];
  }

  set corporateUsername(value: string) {
    _updateColumn(
      this,
      this._provider.propertyMapping.corporateUsername,
      value
    );
  }

  get corporateDisplayName() {
    return this._entity[this._provider.propertyMapping.corporateDisplayName];
  }

  set corporateDisplayName(value: string) {
    _updateColumn(
      this,
      this._provider.propertyMapping.corporateDisplayName,
      value
    );
  }

  get thirdPartyUsername(): string {
    return this._entity[this._provider.propertyMapping.thirdPartyUsername];
  }

  set thirdPartyUsername(value: string) {
    _updateColumn(
      this,
      this._provider.propertyMapping.thirdPartyUsername,
      value
    );
  }

  get thirdPartyId(): string {
    return this._entity[this._provider.propertyMapping.thirdPartyId];
  }

  set thirdPartyId(value: string) {
    _updateColumn(this, this._provider.propertyMapping.thirdPartyId, value);
  }

  get thirdPartyAvatar() {
    return this._entity[this._provider.propertyMapping.thirdPartyAvatar];
  }

  set thirdPartyAvatar(value: string) {
    _updateColumn(this, this._provider.propertyMapping.thirdPartyAvatar, value);
  }

  get serviceAccountMail() {
    return this._entity[this._provider.propertyMapping.serviceAccountMail];
  }

  set serviceAccountMail(value: string) {
    _updateColumn(
      this,
      this._provider.propertyMapping.serviceAccountMail,
      value
    );
  }

  get isServiceAccount() {
    return this._entity[this._provider.propertyMapping.isServiceAccount];
  }

  set isServiceAccount(value: boolean) {
    _updateColumn(this, this._provider.propertyMapping.isServiceAccount, value);
  }

  dataStore(): ITableLinkInstanceDataHelpers {
    return createDataHelpers(this, this._provider);
  }

  internal(): ITableLinkInstanceInternalHelpers {
    return createInternalHelpers(this);
  }
}

function _updateColumn(self, columnName, newValue): void {
  if (!self._originalEntity) {
    self._originalEntity = Object.assign({}, self._entity);
  }

  if (newValue !== self._originalEntity[columnName]) {
    self._entity[columnName] = newValue;
    self._columnUpdates[columnName] = newValue;
    // CONSIDER: debug output instead of console
    console.log(`${columnName} toggled for link ${self._id} to: ${newValue}`);
  } else {
    // Noop or being set back to the original
    if (self._entity[columnName] !== self._originalEntity[columnName]) {
      self._entity[columnName] = self._originalEntity[columnName];
      delete self._columnUpdates[columnName];
      console.log(
        `${columnName} toggled back to original value for ${self._id} to: ${self._entity[columnName]}`
      );
    }
  }
}

// Unfortunately I miss my internal and Symbol hacks in TS so this is
// for my entertainment I guess.

function createInternalHelpers(self): ITableLinkInstanceInternalHelpers {
  return {
    acknowledgeSuccessfulUpdate: acknowledgeSuccessfulUpdate.bind(null, self),
    getDirtyColumns: getDirtyColumns.bind(null, self),
    getDirectEntity: getDirectEntity.bind(null, self),
  };
}

function acknowledgeSuccessfulUpdate(self) {
  self._columnUpdates = {};
  self._originalEntity = null;
}

function getDirectEntity(self) {
  return self._entity;
}

function getDirtyColumns(self) {
  return self._columnUpdates;
}

function createDataHelpers(
  link: CorporateTableLink,
  provider: TableLinkProvider
): ITableLinkInstanceDataHelpers {
  return {
    update: async (): Promise<boolean> => {
      return provider.updateLink(link);
    },

    save: async (): Promise<boolean> => {
      try {
        return provider.updateLink(link);
      } catch (error) {
        // SAVE is different in update in that it only saves if changes
        // are needed (not an error if such). UPDATE will throw if no
        // changes are necessary to be stored.
        if (error && error['noUpdatesRequired'] === true) {
          return false;
        } else if (error) {
          throw error;
        }
      }
    },

    delete: async (): Promise<boolean> => {
      return provider.deleteLink(link);
    },
  };
}
