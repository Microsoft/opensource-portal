//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ICorporateLink } from '../../../business/corporateLink';
import { PostgresLinkProvider } from './postgresLinkProvider';

export interface IInternalPostgresLinkOptions {
  provider: PostgresLinkProvider;
}

export interface IPostgresLinkInstanceDataHelpers {
  delete: () => Promise<boolean>;
  update: () => Promise<boolean>;
  save: () => Promise<boolean>;
}

export interface IPostgresLinkInstanceInternalHelpers {
  acknowledgeSuccessfulUpdate: () => void;
  getDirtyColumns: () => any;
  getDirectEntity: () => any;
}

export class CorporateLinkPostgres implements ICorporateLink {
  private _provider: PostgresLinkProvider;

  private _id: string;
  private _entity: any;
  private _columnUpdates: any;
  private _originalEntity: any;

  constructor(linkInternalOptions: IInternalPostgresLinkOptions, row: any) {
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

  set created(value: Date) {
    throw new Error("The 'created' field cannot be set or updated");
  }

  get corporateId(): string {
    return this._entity[this._provider.propertyMapping.corporateId];
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

  get corporateMailAddress() {
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

  get isServiceAccount() {
    return this._entity[this._provider.propertyMapping.isServiceAccount];
  }

  set isServiceAccount(value: boolean) {
    _updateColumn(this, this._provider.propertyMapping.isServiceAccount, value);
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

  dataStore(): IPostgresLinkInstanceDataHelpers {
    return createDataHelpers(this, this._provider);
  }

  internal(): IPostgresLinkInstanceInternalHelpers {
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

function createInternalHelpers(self): IPostgresLinkInstanceInternalHelpers {
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
  link: CorporateLinkPostgres,
  provider: PostgresLinkProvider
): IPostgresLinkInstanceDataHelpers {
  return {
    update: async (): Promise<boolean> => {
      return provider.updateLink(link);
    },

    save: async (): Promise<boolean> => {
      let actualUpdate = null;
      try {
        actualUpdate = provider.updateLink(link);
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
      return actualUpdate;
    },

    delete: async (): Promise<boolean> => {
      await provider.deleteLink(link);
      return true;
    },
  };
}
