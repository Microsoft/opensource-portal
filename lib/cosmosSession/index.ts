//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { CosmosClient, Database, Container } from '@azure/cosmos';
import { Store } from 'express-session';
import { ErrorHelper, IAppSession } from '../../transitional';

export interface ICosmosSessionProviderOptions {
  endpoint: string;
  key: string;
  ttl?: number;
  database?: string;
  collection?: string;
}

export default class CosmosSessionStore extends Store {
  private _options: ICosmosSessionProviderOptions;
  private _client: CosmosClient;
  private _initialized: boolean;
  private _database: Database;
  private _collection: Container;

  // session: Express.Session
  constructor(options: ICosmosSessionProviderOptions) {
    super();
    this._options = options;
  }

  async initialize() {
    if (this._initialized) {
      return;
    }
    const { endpoint, key } = this._options;
    if (!endpoint) {
      throw new Error('options.endpoint required');
    }
    if (!key) {
      throw new Error('options.key required');
    }
    if (!this._options.collection) {
      throw new Error('options.collection required');
    }
    if (!this._options.database) {
      throw new Error('options.database required');
    }
    this._client = new CosmosClient({ endpoint, key });
    this._database = (
      await this._client.databases.createIfNotExists({
        id: this._options.database,
      })
    ).database;
    this._collection = (
      await this._database.containers.createIfNotExists({
        id: this._options.collection,
      })
    ).container;
    this._initialized = true;
  }

  get = (sid: string, callback) => {
    this.throwIfNotInitialized();
    this._collection
      .item(sid, sid)
      .read()
      .then((response) => {
        if (response.resource) {
          const clone = Object.assign({}, response.resource);
          delete clone._attachments;
          delete clone._etag;
          delete clone._rid;
          delete clone._self;
          delete clone._ts;
          delete clone.ttl;
          if (callback) {
            return callback(null, clone);
          }
        } else {
          return callback(null, null);
        }
      })
      .catch((error) => {
        console.dir(error);
        if (callback) {
          return callback(error);
        }
      });
  };

  // This required method is used to upsert a session into the store given a session
  // ID (sid) and session (session) object. The callback should be called as
  // callback(error) once the session has been set in the store.
  destroy = (sid: string, callback) => {
    this.throwIfNotInitialized();
    this._collection
      .item(sid, sid)
      .delete()
      .then((ok) => {
        if (callback) {
          return callback();
        }
      })
      .catch((error) => {
        if (ErrorHelper.IsNotFound(error)) {
          // this is OK
        } else {
          console.dir(error);
        }
        if (callback) {
          return callback();
          // We do not bubble any errors here.
        }
      });
  };

  // The session argument should be a session if found, otherwise null or undefined if the
  // session was not found (and there was no error). A special case is made when
  // error.code === 'ENOENT' to act like callback(null, null).
  set = (sid: string, session: IAppSession, callback) => {
    this.throwIfNotInitialized();
    if (sid !== session.id) {
      throw new Error(
        "The 'sid' parameter value must match the value of 'session.id'."
      );
    }
    const item = Object.assign({}, session, {
      id: sid,
      ttl: this._options.ttl,
      seen: new Date(),
    });
    this._collection.items
      .upsert(item)
      .then((ok) => {
        if (callback) {
          return callback();
        }
      })
      .catch((error) => {
        console.dir(error);
        if (callback) {
          return callback(
            null,
            new Error(`Error upserting data to the database: ${error}`)
          );
        }
      });
  };

  // This recommended method is used to "touch" a given session given a session ID
  // (sid) and session (session) object. The callback should be called as
  // callback(error) once the session has been touched.
  //
  // This is primarily used when the store will automatically delete idle sessions
  // and this method is used to signal to the store the given session is active,
  // potentially resetting the idle timer.
  touch = (sid: string, session: IAppSession, callback) => {
    this.set(sid, session, callback);
  };

  private throwIfNotInitialized() {
    if (!this._initialized) {
      throw new Error(
        'This provider must be initialized before it can be used'
      );
    }
  }

  // optional: all: (callback: (err: any, obj?: { [sid: string]: IAppSession; } | null) => void) => void;
  // optional: length: (callback: (err: any, length?: number | null) => void) => void;
  // optional: clear: (callback?: (err?: any) => void) => void;
}
