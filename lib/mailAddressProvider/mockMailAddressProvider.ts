//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IMailAddressProvider } from '.';

export default function createMailAddressProvider() {
  const upnToEmails = new Map();
  return ({
    getAddressFromUpn: async (upn, callback) => {
      if (upnToEmails.has(upn)) {
        return upnToEmails.get(upn);
      }
      throw new Error(`No e-mail address known for "${upn}".`);
    },
    // testability:
    getUpnToEmails: function () {
      return upnToEmails;
    },
  } as unknown) as IMailAddressProvider;
}
