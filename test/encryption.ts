//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import chai = require('chai');
const assert = chai.assert;
import crypto = require('crypto');
import { v4 as uuidV4 } from 'uuid';

const encryption = require('../lib/encryption');

function generate32bitKey(callback) {
  crypto.randomBytes(32, callback);
}

describe('encryption', () => {
  describe('encryptEntity', () => {
    it('unencrypted entities can be processed', () => {
      let dynamicKeyId = uuidV4();
      generate32bitKey((error, key) => {
        let keyEncryptionKeys = {
          dynamicKeyId: key,
        };
        let sampleEncryptionOptions = {
          keyEncryptionKeyId: dynamicKeyId,
          encryptedPropertyNames: ['secret'],
          keyEncryptionKeys: keyEncryptionKeys,
        };
        let entity = {
          hello: 'world',
          notSecret: 'this is not a secret',
          secret: 'this is a secret',
          superSecret: 'the password is password',
        };
        let partitionKey = 'partition' + uuidV4();
        let rowKey = 'row' + uuidV4();
        encryption.decryptEntity(
          partitionKey,
          rowKey,
          entity,
          sampleEncryptionOptions,
          (anotherError, roundtripEntity) => {
            assert.deepEqual(
              entity,
              roundtripEntity,
              'roundtripEntity is equal to entity'
            );
          }
        );
      });
    }),
      it('should have encryption metadata', () => {
        let dynamicKeyId = uuidV4();
        generate32bitKey((error, key) => {
          let keyEncryptionKeys = {
            dynamicKeyId: key,
          };
          let sampleEncryptionOptions = {
            keyEncryptionKeyId: dynamicKeyId,
            encryptedPropertyNames: ['secret', 'superSecret'],
            keyEncryptionKeys: keyEncryptionKeys,
          };
          let secretEntity = {
            hello: 'world',
            notSecret: 'this is not a secret',
            secret: 'this is a secret',
            superSecret: 'the password is password',
          };
          let partitionKey = 'partition' + uuidV4();
          let rowKey = 'row' + uuidV4();
          encryption.encryptEntity(
            partitionKey,
            rowKey,
            secretEntity,
            sampleEncryptionOptions,
            (error, encryptedEntity) => {
              assert.isDefined(
                encryptedEntity['_ClientEncryptionMetadata1'],
                'defined encryption metadata'
              );
              assert.isDefined(
                encryptedEntity['_ClientEncryptionMetadata2'],
                'defined encrypted keys list entry'
              );
            }
          );
        });
      }),
      it('should be able to decrypt itself', () => {
        let dynamicKeyId = uuidV4();
        generate32bitKey((error, key) => {
          let keyEncryptionKeys = {
            dynamicKeyId: key,
          };
          let sampleEncryptionOptions = {
            keyEncryptionKeyId: dynamicKeyId,
            encryptedPropertyNames: ['secret', 'superSecret'],
            keyEncryptionKeys: keyEncryptionKeys,
          };
          let secretEntity = {
            hello: 'world',
            notSecret: 'this is not a secret',
            secret: 'this is a secret',
            superSecret: 'the password is password',
          };
          let partitionKey = 'partition' + uuidV4();
          let rowKey = 'row' + uuidV4();
          encryption.encryptEntity(
            partitionKey,
            rowKey,
            secretEntity,
            sampleEncryptionOptions,
            (error, encryptedEntity) => {
              assert.notDeepEqual(
                encryptedEntity,
                secretEntity,
                'encryptedEntity is not equal to secretEntity'
              );
              encryption.decryptEntity(
                partitionKey,
                rowKey,
                encryptedEntity,
                sampleEncryptionOptions,
                (anotherError, roundtripEntity) => {
                  assert.deepEqual(
                    roundtripEntity,
                    secretEntity,
                    'roundtripEntity is equal to secretEntity'
                  );
                }
              );
            }
          );
        });
      });
  });
});
