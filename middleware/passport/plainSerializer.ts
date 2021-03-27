//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { LegacySerializer } from './serializer';

// function i() {
//   console.warn('Plain text session tokens are in use. Not recommended for production.');
// }

export default class PlainSerializer implements LegacySerializer {
  serialize(config, user, done) {
    return done(null, user);
  }

  deserialize(config, user, done) {
    return done(null, user);
  }

  initialize(options: any, app: any) {}
}
