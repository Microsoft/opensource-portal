//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { URL } from 'url';

const cachedKeys = new Map();
const cacheKeysInMemory = true;

const secretsPath = '/secrets/';

function keyVaultSecretResolver(keyVaultClient, id, callback) {
  const cachedKey = cachedKeys.get(id);
  if (cachedKey !== undefined) {
    return callback(null, cachedKey);
  }
  const secretUrl = new URL(id);
  const vaultBaseUrl = secretUrl.origin;
  const i = secretUrl.pathname.indexOf(secretsPath);
  if (i < 0) {
    return callback(
      new Error('The requested resource must be a KeyVault secret')
    );
  }
  let secretName = secretUrl.pathname.substr(i + secretsPath.length);
  let version = '';
  const versionIndex = secretName.indexOf('/');
  if (versionIndex >= 0) {
    version = secretName.substr(versionIndex + 1);
    secretName = secretName.substr(0, versionIndex);
  }
  keyVaultClient.getSecret(
    vaultBaseUrl,
    secretName,
    version,
    (getSecretError, secretResponse) => {
      if (getSecretError) {
        return callback(getSecretError);
      }
      const secretValue = secretResponse.value;
      if (cacheKeysInMemory === true) {
        cachedKeys.set(id, secretValue);
      }
      return callback(null, secretValue);
    }
  );
}

export default function createKeyVaultResolver(keyVaultClient) {
  return keyVaultSecretResolver.bind(undefined, keyVaultClient);
}
