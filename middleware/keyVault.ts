//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import adalNode, { TokenResponse } from 'adal-node'; // NOTE: this is deprecated-ish
import { KeyVaultClient, KeyVaultCredentials } from 'azure-keyvault';

export default function createClient(kvConfig) {
  if (!kvConfig.clientId) {
    throw new Error(
      'KeyVault client ID required at this time for the middleware to initialize.'
    );
  }
  if (!kvConfig.clientSecret) {
    throw new Error(
      'KeyVault client credential/secret required at this time for the middleware to initialize.'
    );
  }
  const authenticator = (challenge, authCallback) => {
    const context = new adalNode.AuthenticationContext(challenge.authorization);
    return context.acquireTokenWithClientCredentials(
      challenge.resource,
      kvConfig.clientId,
      kvConfig.clientSecret,
      (tokenAcquisitionError, tokenResponse) => {
        if (tokenAcquisitionError) {
          return authCallback(tokenAcquisitionError);
        }
        const tk = tokenResponse as TokenResponse;
        const authorizationValue = `${tk.tokenType} ${tk.accessToken}`;
        return authCallback(null, authorizationValue);
      }
    );
  };
  const credentials = new KeyVaultCredentials(authenticator, null);
  const keyVaultClient = new KeyVaultClient(credentials);
  return keyVaultClient;
}
