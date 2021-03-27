//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

const assert = require('chai').assert;
const keyVaultHelper = require('keyvault-configuration-resolver');
const fakeKeyVaultClient = require('./fakeKeyVaultClient');

function createFakeWithKeys() {
  const faker = fakeKeyVaultClient();
  const secretId = faker.storeSecret('test', 'big secret', {
    tag1: 'p1',
    tag2: 'and tag 2',
  });
  return [faker, secretId];
}

describe('configuration', () => {
  // config as code: tests have moved to the refactored npm, painless-config-as-code

  describe('keyVaultHelper', () => {
    it('non-URL values passthrough', () => {
      const fake = createFakeWithKeys(); // es6 destructuring would be nice
      const keyVaultClient = keyVaultHelper(fake[0]);
      const config = {
        a: 'animal',
        b: 'bat',
        c: 'cherry',
        d: true,
        e: 5,
      };
      keyVaultClient.getObjectSecrets(config, (error) => {
        assert.isNull(error, 'no exception');
        assert.equal(config.a, 'animal', 'string works');
        assert.equal(config.b, 'bat', 'string works');
        assert.equal(config.c, 'cherry', 'string works');
        assert.isTrue(config.d, 'bool works');
        assert.equal(config.e, 5, 'number is unaffected');
      });
    });
    it('keyvault:// protocol works', () => {
      const fake = createFakeWithKeys(); // es6 destructuring would be nice
      const keyVaultClient = keyVaultHelper(fake[0]);
      const secretId = fake[1];
      const keyVaultSchemeSecretId = secretId.replace(
        'https://',
        'keyvault://'
      );
      const config = {
        bigPasscode: keyVaultSchemeSecretId,
      };
      keyVaultClient.getObjectSecrets(config, () => {
        assert.equal(config.bigPasscode, 'big secret', 'secret read OK');
      });
    });
    it('deeply nested KeyVault URLs work', () => {
      const fake = createFakeWithKeys(); // es6 destructuring would be nice
      const keyVaultClient = keyVaultHelper(fake[0]);
      const secretId = fake[1];
      const keyVaultSchemeSecretId = secretId.replace(
        'https://',
        'keyvault://'
      );
      const config = {
        deep: {
          object: {
            nesting: {
              test: {
                value: {
                  is: keyVaultSchemeSecretId,
                },
              },
            },
          },
        },
      };
      keyVaultClient.getObjectSecrets(config, () => {
        assert.equal(
          config.deep.object.nesting.test.value.is,
          'big secret',
          'secret read OK'
        );
      });
    });
    it('keyvault:// tag properties work', () => {
      const fake = createFakeWithKeys(); // es6 destructuring would be nice
      const keyVaultClient = keyVaultHelper(fake[0]);
      const secretId = fake[1];
      const keyVaultSchemeSecretId = secretId.replace(
        'https://',
        'keyvault://'
      );
      const keyVaultSchemeSecretIdWithTag = secretId.replace(
        'https://',
        'keyvault://tag1@'
      );
      const config = {
        taggedProperty: keyVaultSchemeSecretIdWithTag,
        kvProperty: keyVaultSchemeSecretId,
      };
      keyVaultClient.getObjectSecrets(config, () => {
        assert.equal(config.kvProperty, 'big secret', 'secret read OK');
        assert.equal(config.taggedProperty, 'p1', 'tag read OK');
      });
    });
    it('keyvault:// tag properties return undefined if missing', () => {
      const fake = createFakeWithKeys(); // es6 destructuring would be nice
      const keyVaultClient = keyVaultHelper(fake[0]);
      const secretId = fake[1];
      const keyVaultSchemeSecretIdWithTag = secretId.replace(
        'https://',
        'keyvault://undefinedtagthing@'
      );
      const config = {
        taggedProperty: keyVaultSchemeSecretIdWithTag,
      };
      keyVaultClient.getObjectSecrets(config, () => {
        assert.isUndefined(config.taggedProperty, '=== undefined');
      });
    });
    it('URL values passthrough', () => {
      const fake = createFakeWithKeys(); // es6 destructuring would be nice
      const keyVaultClient = keyVaultHelper(fake[0]);
      const secretId = fake[1];
      const config = {
        a: secretId,
      };
      keyVaultClient.getObjectSecrets(config, (error) => {
        assert.isNull(error, 'no exception');
        assert.equal(config.a, secretId, 'KeyVault URL is passed through');
      });
    });
    it('keyvault:// on an invalid secret stops processing', () => {
      const fake = createFakeWithKeys(); // es6 destructuring would be nice
      const keyVaultClient = keyVaultHelper(fake[0]);
      const config = {
        a: 'keyvault://invalid/secrets/hello/1',
      };
      keyVaultClient.getObjectSecrets(config, (error) => {
        assert.isNotNull(
          error,
          'exception thrown due to KeyVault client error'
        );
      });
    });
  });
});
