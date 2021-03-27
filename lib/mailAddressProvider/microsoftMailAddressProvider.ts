//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IProviders } from '../../transitional';
import { IMailAddressProvider } from '.';

export default function createMailAddressProvider(options) {
  const config = options.config;
  if (!config.identity || !config.identity.url || !config.identity.pat) {
    throw new Error('Not configured for the Identity service');
  }
  const providers = options.providers as IProviders;
  if (!providers) {
    throw new Error(
      'The microsoftMailAddressProvider requires that all provider instances are passed in as options'
    );
  }
  return {
    getAddressFromUpn: (upn: string) => {
      return providers.graphProvider.getMailAddressByUsername(upn);
    },
  } as IMailAddressProvider;
}
