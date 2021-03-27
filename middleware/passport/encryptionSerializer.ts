//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { decryptEntity, encryptEntity } from '../../lib/encryption';
import { wrapError } from '../../utils';

import { LegacySerializer } from './serializer';

const userEncryptedEntities = {
  azure: new Set(),
  github: new Set(['accessToken']),
  githubIncreasedScope: new Set(['accessToken']),
};

const userEntityId = {
  github: 'id',
  githubIncreasedScope: 'id',
  azure: 'oid',
};

function validateNoRichProperties(properties) {
  for (const key in properties) {
    if (properties[key] === undefined || properties[key] === null) {
      continue;
    }
    if (typeof properties[key] === 'object') {
      console.warn(
        `The property ${key} is an object. To help with diagnosing the underlying area with the problem, here is the current value of the object:`
      );
      console.warn(properties[key]);
      return new Error(`Session property ${key} is an object.`);
    }
  }
}

function serializeEntity(options, entityName, entity, callback) {
  const config = options.config;
  const partitionKey = entityName;
  const idPropertyName = userEntityId[entityName];
  const rowKey = entity[idPropertyName];
  const richObjectError = validateNoRichProperties(entity);
  if (richObjectError !== undefined) {
    return callback(richObjectError);
  }
  if (rowKey === undefined) {
    return callback(
      new Error('The unique identifier for the user entity was not available.')
    );
  }
  const keyResolver = options.keyResolver;
  if (keyResolver === undefined) {
    return callback(
      new Error('A key resolver must be supplied to use encryption.')
    );
  }
  const encryptionOptions = {
    keyEncryptionKeyId: config.session.encryptionKeyId,
    keyResolver: keyResolver,
    encryptedPropertyNames: userEncryptedEntities[entityName],
    binaryProperties: 'base64',
  };
  encryptEntity(
    partitionKey,
    rowKey,
    entity,
    encryptionOptions,
    (encryptError, encryptedEntity) => {
      if (encryptError) {
        return callback(
          wrapError(
            encryptError,
            'There was a problem with the security subsystem starting your session.'
          )
        );
      }
      callback(null, encryptedEntity);
    }
  );
}

function deserializeEntity(options, entityName, entity, callback) {
  const partitionKey = entityName;
  const idPropertyName = userEntityId[entityName];
  if (idPropertyName === undefined) {
    return callback(new Error('The entity type is not configured properly.'));
  }
  const rowKey = entity[idPropertyName];
  if (rowKey === undefined) {
    return callback(
      new Error('The unique identifier for the user entity was not available.')
    );
  }
  const keyResolver = options.keyResolver;
  if (keyResolver === undefined) {
    return callback(
      new Error('A key resolver must be supplied to encrypt/decrypt.')
    );
  }
  const encryptionOptions = {
    keyResolver: keyResolver,
    binaryProperties: 'base64',
  };
  decryptEntity(
    partitionKey,
    rowKey,
    entity,
    encryptionOptions,
    (encryptError, decryptedEntity) => {
      if (encryptError) {
        const userError = wrapError(
          encryptError,
          'There was a problem with the security subsystem retrieving your session.'
        );
        userError['forceSignOut'] = true;
        return callback(userError);
      }
      callback(null, decryptedEntity);
    }
  );
}

export default class EncryptionSerializer implements LegacySerializer {
  private options: any;

  serialize(config, user, done) {
    return Promise.all(
      Object.getOwnPropertyNames(userEncryptedEntities).map((entityName) => {
        return new Promise((resolve, reject) => {
          const entityPresent = user[entityName];
          if (entityPresent !== undefined) {
            const entityOriginalValue = entityPresent;
            delete user[entityName];
            return serializeEntity(
              this.options,
              entityName,
              entityOriginalValue,
              (error, value) => {
                user[entityName] = value;
                return error ? reject(error) : resolve(undefined);
              }
            );
          } else {
            return resolve(undefined);
          }
        });
      })
    )
      .then((ok) => {
        return done(null, user);
      })
      .catch((error) => {
        return done(error);
      });
  }

  deserialize(config, user, done) {
    const u = {};
    return Promise.all(
      Object.getOwnPropertyNames(user).map((entityName) => {
        return new Promise((resolve, reject) => {
          if (userEncryptedEntities[entityName] !== undefined) {
            let entityValue = user[entityName];
            return deserializeEntity(
              this.options,
              entityName,
              entityValue,
              (error, result) => {
                u[entityName] = result;
                return error ? reject(error) : resolve(undefined);
              }
            );
          } else {
            return resolve(undefined);
          }
        });
      })
    )
      .then((ok) => {
        for (const unencryptedEntity in user) {
          if (userEncryptedEntities[unencryptedEntity] === undefined) {
            u[unencryptedEntity] = user[unencryptedEntity];
          }
        }
        return done(null, u);
      })
      .catch((error) => {
        return done(error);
      });
  }

  initialize(options: any, app: any) {
    this.options = options;
    app._sessionSerializer = this;
  }
}
