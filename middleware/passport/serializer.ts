//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import EncryptionSerializer from './encryptionSerializer';
import PlainSerializer from './plainSerializer';

export abstract class LegacySerializer {
  abstract serialize(config: any, user: any, done): void;
  abstract deserialize(config: any, user: any, done): void;
  abstract initialize(options: any, app: any): void;
}

interface ISerializerInstances {
  encrypted: LegacySerializer;
  plain: LegacySerializer;
}

function createSerializer() {
  const serializer: ISerializerInstances = {
    encrypted: null,
    plain: null,
  };
  function ensureSerializer(config) {
    const serializerKey = config.session.encryption;
    const isEncrypted = serializerKey === true;
    const key = isEncrypted ? 'encrypted' : 'plain';
    if (!serializer[key]) {
      serializer[key] = isEncrypted
        ? new EncryptionSerializer()
        : new PlainSerializer();
    }
    if (serializer[key] === undefined) {
      throw new Error(`Could not prepare serializer of type ${serializerKey}`);
    }
    return serializer[key];
  }

  function createSerialize(options) {
    const instance = ensureSerializer(options.config);
    return instance.serialize.bind(instance, options);
  }

  function createDeserialize(options) {
    const instance = ensureSerializer(options.config);
    return instance.deserialize.bind(instance, options);
  }

  function initialize(options, app) {
    const serializerInstance = ensureSerializer(options.config);
    const initializer = serializerInstance.initialize;
    if (initializer) {
      // Allow an opportunity to provide a warning or connect a route
      initializer.call(serializerInstance, options, app);
    }
  }

  return {
    initialize: initialize,
    serialize: createSerialize,
    deserialize: createDeserialize,
  };
}

export default createSerializer();
