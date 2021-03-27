//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// Provides a one-to-many mapping of links, where a single third-party username is
// associated with at most one corporate identity. A corporate identity may be associated
// with multiple third-party identities. The instance of the provider is able to also
// take in an optional name/identifier for the type of third-party, though the default
// assumed provider type is 'github' for the purposes of this implementation.

// An ideal provider provides a thin business-oriented wrapper over the data layer
// free of any caching or other logic. A provider may require additional options or
// capabilities through the 'options' constructor object to provide specific goals
// or capabilities.

// It is assumed that a link have a unique identifier, whether constructed or passed
// through the data layer, that can uniquely refer to a link for key operations.

const onlySupportedThirdPartyType = 'github';

import { v4 as uuidV4 } from 'uuid';

import { InnerError } from '../../../transitional';
import {
  ICorporateLinkProperties,
  ICorporateLink,
  ICorporateLinkExtended,
} from '../../../business/corporateLink';

import { CorporateLinkPostgres } from './postgresLink';
import {
  PostgresPoolQueryAsync,
  PostgresPoolQuerySingleRowAsync,
} from '../../postgresHelpers';
import { ILinkProvider } from '..';

const linkProviderInstantiationTypeProperty = '_i';

const dehydratedIdentityKey = '_lpi';
const dehydratedPostgresProviderName = 'pg';
const dehydratedPostgresProviderVersion = '0';
const dehydratedPostgresProviderIdentitySeperator = '_';
const dehydratedPostgresProviderIdentity = `${dehydratedPostgresProviderName}${dehydratedPostgresProviderIdentitySeperator}${dehydratedPostgresProviderVersion}`;

enum LinkInstantiatedType {
  Row,
  Rehydrated,
}

export interface IPostgresLinkProperties extends ICorporateLinkProperties {
  linkId: string;
  created: string;
}

const linkInterfacePropertyMapping: IPostgresLinkProperties = {
  thirdPartyId: 'thirdpartyid',
  thirdPartyUsername: 'thirdpartyusername',
  thirdPartyAvatar: 'thirdpartyavatar',
  corporateId: 'corporateid',
  corporateUsername: 'corporateusername',
  corporateDisplayName: 'corporatename',
  corporateMailAddress: 'corporatemail',
  corporateAlias: 'corporatealias',

  isServiceAccount: 'serviceaccount',
  serviceAccountMail: 'serviceaccountmail',

  // specific to this type
  linkId: 'linkid',
  created: 'created',
};

const coreColumns = [
  'linkid',
  'thirdpartyid',
  'thirdpartyusername',
  'thirdpartyavatar',
  'corporateid',
  'corporateusername',
  'corporatename',
  'corporatemail',
  'corporatealias',
  'serviceaccount',
  'serviceaccountmail',
  'created',
];
const coreColumnsList = coreColumns.join(', ');

export class PostgresLinkProvider implements ILinkProvider {
  private _thirdPartyType: string;
  private _providers: any;
  private _pool: any;
  private _tableName: any;
  private _internalThirdPartyTypeValue: string;
  private _options: any;

  public readonly propertyMapping: IPostgresLinkProperties = linkInterfacePropertyMapping;

  public readonly serializationIdentifierVersion: string = dehydratedPostgresProviderIdentity;

  constructor(providers, options) {
    if (!providers) {
      throw new Error(
        'The PostgresLinkProvider requires that available providers are passed into the constructor'
      );
    }
    options = options || {};

    if (!options.tableName) {
      throw new Error(
        'Missing the name of the table for links for the Postgres link provider'
      );
    }

    const thirdPartyType =
      options.thirdPartyType || onlySupportedThirdPartyType;
    const internalThirdPartyTypeValue =
      options.thirdPartyTypeValue ||
      (thirdPartyType === 'github'
        ? options.githubThirdPartyName
        : thirdPartyType);

    this._thirdPartyType = thirdPartyType;
    this._internalThirdPartyTypeValue = internalThirdPartyTypeValue;
    this._providers = providers;
    this._tableName = options.tableName;
    this._pool = providers.postgresPool;
    this._options = options;
  }

  async initialize(): Promise<ILinkProvider> {
    const self = this;
    const rows = await PostgresPoolQueryAsync(
      this._pool,
      `
      SELECT
        COUNT(thirdpartyid) as thirdpartycount,
        COUNT(corporateid) as corporatecount
      FROM ${self._tableName}
      WHERE
        thirdpartytype = $1`,
      [self._internalThirdPartyTypeValue]
    );
    if (rows.length === 1) {
      const row = rows[0];
      console.log(
        `FYI: Postgres: there are ${row.thirdpartycount} third-party links to ${row.corporatecount} corporate users`
      );
    }
    return this;
  }

  get thirdPartyType() {
    return this._thirdPartyType;
  }

  getByPostgresLinkId(id: string): Promise<CorporateLinkPostgres> {
    // Specific to this link provider implementation
    return this._getSingleRow({
      columnName: 'linkid',
      columnValue: id,
      columnIsLowercase: false,
    });
  }

  async getByThirdPartyUsername(
    username: string
  ): Promise<CorporateLinkPostgres> {
    return this._getSingleRow({
      columnName: 'thirdpartyusername',
      columnValue: username,
      columnIsLowercase: true,
    });
  }

  getByThirdPartyId(id: string): Promise<CorporateLinkPostgres> {
    return this._getSingleRow({
      columnName: 'thirdpartyid',
      columnValue: id,
      columnIsLowercase: false,
    });
  }

  queryByCorporateId(id: string): Promise<CorporateLinkPostgres[]> {
    return this._getRows({
      columnName: 'corporateid',
      columnValue: id,
      columnIsLowercase: false,
    });
  }

  async getAll(): Promise<CorporateLinkPostgres[]> {
    const internalThirdPartyTypeValue = this._internalThirdPartyTypeValue;
    // CONSIDER: the table provider sorts by aadupn and then ghu!
    // TODO: is an order by of interest here?
    const results = await PostgresPoolQueryAsync(
      this._pool,
      `
      SELECT
        ${coreColumnsList}
      FROM ${this._tableName}
      WHERE
        thirdpartytype = $1
    `,
      [internalThirdPartyTypeValue]
    );
    let r = [];
    for (let i = 0; i < results.rows.length; i++) {
      const row = results.rows[i];
      const link = this.createLinkInstanceFromRow(row);
      const ll = link;
      r.push(ll);
    }
    return r;
  }

  async getAllCorporateIds(): Promise<string[]> {
    const internalThirdPartyTypeValue = this._internalThirdPartyTypeValue;
    const results = await PostgresPoolQueryAsync(
      this._pool,
      `
      SELECT
        corporateid
      FROM ${this._tableName}
      WHERE
        thirdpartytype = $1
    `,
      [internalThirdPartyTypeValue]
    );
    let r = [];
    if (results && results.rows) {
      r = results.rows.map((row) => String(row.corporateid));
    }
    return r;
  }

  queryByCorporateUsername(username: string): Promise<CorporateLinkPostgres[]> {
    return this._getRows({
      columnName: 'corporateusername',
      columnValue: username,
      columnIsLowercase: true,
    });
  }

  async createLink(link: ICorporateLink): Promise<string> {
    const self = this;
    const linkId = uuidV4(); // primary key protected
    if (
      !link.thirdPartyId ||
      !link.thirdPartyUsername ||
      !link.corporateId ||
      !link.corporateUsername
    ) {
      throw new Error('Missing a required value');
    }
    const created = new Date();
    try {
      const insertResult = await PostgresPoolQueryAsync(
        self._pool,
        `
        INSERT INTO ${this._tableName}(
          linkid,
          thirdpartytype,
          thirdpartyid, thirdpartyusername, thirdpartyavatar,
          corporateid, corporateusername, corporatename,
          serviceaccount, serviceaccountmail,
          created)
        VALUES (
          $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11
        )
      `,
        [
          linkId,
          self._internalThirdPartyTypeValue,
          link.thirdPartyId,
          link.thirdPartyUsername,
          link.thirdPartyAvatar,
          link.corporateId,
          link.corporateUsername,
          link.corporateDisplayName,
          link.isServiceAccount,
          link.serviceAccountMail,
          created,
        ]
      );
      return linkId;
    } catch (error) {
      if (error.message && error.message.includes('duplicate key value')) {
        const ie: InnerError = new Error(
          'A link already exists for the identity'
        );
        ie.inner = error;
        error = ie;
      }
      throw error;
    }
  }

  async updateLink(linkInstance: ICorporateLink): Promise<any> {
    const pgl = linkInstance as CorporateLinkPostgres;
    const id = pgl.id;
    const values = [id];
    const internal = pgl.internal();
    const updates = internal.getDirtyColumns();
    const columns = Object.getOwnPropertyNames(updates);
    if (columns.length === 0) {
      const noUpdatesRequired = new Error(
        'No updates were required for the link'
      );
      noUpdatesRequired['noUpdatesRequired'] = true;
      throw noUpdatesRequired;
    }
    const sets = columns
      .map((columnName) => {
        values.push(updates[columnName]);
        const index = values.length;
        return `\n        ${columnName} = \$${index}`;
      })
      .join();
    let sql = `
      UPDATE ${this._tableName}
      SET ${sets}
      WHERE
        linkid = $1
    `;
    const results = await PostgresPoolQueryAsync(this._pool, sql, values);
    // TODO: how to validate updates?
    internal.acknowledgeSuccessfulUpdate();
    return results.rowCount > 0; // whether a row was actually updated
  }

  async deleteLink(link: ICorporateLink): Promise<void> {
    const pgl = link as CorporateLinkPostgres;
    const id = pgl.id;
    await this._deleteSingleRow({ columnName: 'linkid', columnValue: id });
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

  rehydrateLinks(jsonArray: any[]): ICorporateLink[] {
    if (!Array.isArray(jsonArray)) {
      throw new Error('jsonArray must be an array');
    }
    //
    const arr = jsonArray.map(this.rehydrateLink.bind(this));
    return (arr as any[]) as ICorporateLink[];
  }

  dehydrateLink(linkInstance: ICorporateLinkExtended): any {
    // CONSIDER: check whether the current link type feels appropriate to us (PGSQL)
    const pglink = linkInstance as CorporateLinkPostgres;
    const entity = pglink.internal().getDirectEntity();
    const shriveled = Object.assign({}, entity);
    shriveled[dehydratedIdentityKey] = dehydratedPostgresProviderIdentity;
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
    if (identity !== dehydratedPostgresProviderIdentity) {
      const sameProviderType = identity.startsWith(
        `${dehydratedPostgresProviderName}${dehydratedPostgresProviderIdentitySeperator}`
      );
      if (sameProviderType) {
        // Cross-version rehydration not supported
        throw new Error(
          `The hydrated link was created by the same ${dehydratedPostgresProviderName} provider, but a different version: ${identity}`
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

  public createFromRows(rows: any[]): CorporateLinkPostgres[] {
    const list = rows.map((row) => this.createLinkInstanceFromRow(row));
    return list;
  }

  private async _getRows({
    columnName,
    columnValue,
    columnIsLowercase,
  }): Promise<CorporateLinkPostgres[]> {
    let columnWrapperStart = columnIsLowercase ? 'lower(' : '';
    let columnWrapperFinish = columnIsLowercase ? ')' : '';
    const results = await PostgresPoolQueryAsync(
      this._pool,
      `
      SELECT
        ${coreColumnsList}
      FROM ${this._tableName}
      WHERE
        thirdpartytype = $1 AND
        ${columnWrapperStart}${columnName}${columnWrapperFinish} = $2
    `,
      [
        this._internalThirdPartyTypeValue,
        columnIsLowercase ? columnValue.toLowerCase() : columnValue,
      ]
    );
    let r = [];
    for (let i = 0; i < results.rows.length; i++) {
      const row = results.rows[i];
      const link = this.createLinkInstanceFromRow(row);
      const ll = link;
      r.push(ll);
    }
    return r;
  }

  private async _getSingleRow({
    columnName,
    columnValue,
    columnIsLowercase,
  }): Promise<CorporateLinkPostgres> {
    let columnWrapperStart = columnIsLowercase ? 'lower(' : '';
    let columnWrapperFinish = columnIsLowercase ? ')' : '';
    const sql = `
      SELECT
        ${coreColumnsList}
      FROM ${this._tableName}
      WHERE
        thirdpartytype = $1 AND
        ${columnWrapperStart}${columnName}${columnWrapperFinish} = $2
    `;
    const values = [
      this._internalThirdPartyTypeValue,
      columnIsLowercase ? columnValue.toLowerCase() : columnValue,
    ];
    const row = await PostgresPoolQuerySingleRowAsync(this._pool, sql, values);
    return this.createLinkInstanceFromRow(row);
  }

  private async _deleteSingleRow({
    columnName,
    columnValue,
  }): Promise<boolean> {
    const deleteResult = await PostgresPoolQueryAsync(
      this._pool,
      `
      DELETE
      FROM ${this._tableName}
      WHERE
        ${columnName} = $1
    `,
      [columnValue]
    );
    return deleteResult && deleteResult['rowCount'] > 0;
  }

  private createLinkInstanceFromRow(row): CorporateLinkPostgres {
    const linkInternalOptions = {
      provider: this,
    };
    const newLink = new CorporateLinkPostgres(linkInternalOptions, row);
    newLink[linkProviderInstantiationTypeProperty] = LinkInstantiatedType.Row; // in case this helps while debugging
    return newLink;
  }

  private createLinkInstanceFromHydratedEntity(
    jsonObject
  ): CorporateLinkPostgres {
    const linkInternalOptions = {
      provider: this,
    };
    const newLink = new CorporateLinkPostgres(linkInternalOptions, jsonObject);
    newLink[linkProviderInstantiationTypeProperty] =
      LinkInstantiatedType.Rehydrated; // in case this helps while debugging
    return newLink;
  }
}
