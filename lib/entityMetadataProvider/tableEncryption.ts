//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// This is a Node.js implementation of client-side table entity encryption,
// compatible with the official Azure storage .NET library.

import azure from 'azure-storage';
import crypto from 'crypto';
import jose from 'node-jose';

const entityGenerator = azure.TableUtilities.entityGenerator;

// Azure Storage .NET client library - entity encryption keys:
//
// Key:      _ClientEncryptionMetadata1
// Type:     JSON stringified object
// Purpose:  Contains information about the client encryption agent used to
//           encrypt the entity. Contains a uniquely generated content
//           encryption key for the specific row of data.
// Constant: tableEncryptionKeyDetails
//
// Key:      _ClientEncryptionMetadata2
// Type:     Binary buffer
// Purpose:  Encrypted JSON stringified object containing the list of encrypted
//           fields in the entity.
// Constant: tableEncryptionPropertyDetails
const tableEncryptionPropertyDetails = '_ClientEncryptionMetadata2';
const tableEncryptionKeyDetails = '_ClientEncryptionMetadata1';

// Azure Storage encryption agent values: as implemented today, the encryption
// agent for .NET is of version 1.0; initialization vectors are 16-bytes,
// CEKs are 32-bytes, etc. The agent includes the AES algorithm used for
// content keys, but the algorithm is the .NET framework-recognized value and
// not the OpenSSL defined constant. We maintain a map therefore to map
// between the two, but only those which are currently supported by the .NET
// Azure Storage library.
const azureStorageEncryptionAgentProtocol = '1.0';
const azureStorageKeyWrappingAlgorithm = 'A256KW';
const azureStorageContentEncryptionIVBytes = 16;
const azureStorageContentEncryptionKeyBytes = 32;
const azureStorageEncryptionAgentEncryptionAlgorithm =
  'AES_CBC_256'; /* .NET value */
const mapDotNetFrameworkToOpenSslAlgorithm = new Map([
  [azureStorageEncryptionAgentEncryptionAlgorithm, 'aes-256-cbc'],
]);

function openSslFromNetFrameworkAlgorithm(algorithm) {
  const openSslAlgorithm = mapDotNetFrameworkToOpenSslAlgorithm.get(algorithm);
  if (openSslAlgorithm === undefined) {
    throw new Error(
      `The OpenSSL algorithm constant for the .NET Framework value "${algorithm}" is not defined or tested.`
    );
  }
  return openSslAlgorithm;
}

// Hash, encrypt, decrypt, wrap, unwrap and key generation routines

function getSha256Hash(buffer) {
  return crypto.createHash('sha256').update(buffer).digest();
}

function encryptValue(contentEncryptionKey, iv, value): Buffer {
  const cipher = crypto.createCipheriv(
    openSslFromNetFrameworkAlgorithm(
      azureStorageEncryptionAgentEncryptionAlgorithm
    ),
    contentEncryptionKey,
    iv
  );
  return Buffer.concat([cipher.update(value), cipher.final()]);
}

function decryptValue(
  algorithm,
  contentEncryptionKey,
  iv,
  encryptedValue
): Buffer {
  const decipher = crypto.createDecipheriv(algorithm, contentEncryptionKey, iv);
  return Buffer.concat([decipher.update(encryptedValue), decipher.final()]);
}

function generate32bitKey(callback) {
  crypto.randomBytes(azureStorageContentEncryptionKeyBytes, callback);
}

function generateContentEncryptionKey(callback) {
  crypto.randomBytes(
    azureStorageContentEncryptionIVBytes,
    (cryptoError, contentEncryptionIV) => {
      if (cryptoError) {
        return callback(cryptoError);
      }
      generate32bitKey((createKeyError, contentEncryptionKey) => {
        if (createKeyError) {
          return callback(createKeyError);
        }
        callback(null, contentEncryptionIV, contentEncryptionKey);
      });
    }
  );
}

function wrapContentKey(
  keyWrappingAlgorithm,
  keyEncryptionKey,
  contentEncryptionKey,
  callback
) {
  jose.JWA.encrypt(
    keyWrappingAlgorithm,
    keyEncryptionKey,
    contentEncryptionKey
  ).then((result) => {
    return callback(null, result.data);
  }, callback);
}

function unwrapContentKey(
  keyWrappingAlgorithm,
  keyEncryptionKey,
  wrappedContentKeyEncryptedKey,
  callback
) {
  jose.JWA.decrypt(
    keyWrappingAlgorithm,
    keyEncryptionKey,
    wrappedContentKeyEncryptedKey
  ).then((contentEncryptionKey) => {
    return callback(null, contentEncryptionKey);
  }, callback);
}

// Azure encryption metadata object

function createEncryptionData(
  keyId,
  wrappedContentEncryptionKey,
  contentEncryptionIV,
  keyWrappingAlgorithm
) {
  const encryptionData = {
    // PascalCase object per the .NET library
    WrappedContentKey: {
      KeyId: keyId,
      EncryptedKey: base64StringFromBuffer(wrappedContentEncryptionKey),
      Algorithm: keyWrappingAlgorithm,
    },
    EncryptionAgent: {
      Protocol: azureStorageEncryptionAgentProtocol,
      EncryptionAlgorithm: azureStorageEncryptionAgentEncryptionAlgorithm,
    },
    ContentEncryptionIV: base64StringFromBuffer(contentEncryptionIV),
    KeyWrappingMetadata: {},
  };
  return encryptionData;
}

function validateEncryptionData(encryptionData) {
  if (!encryptionData || !encryptionData.EncryptionAgent) {
    throw new Error('No encryption data or encryption data agent.');
  }
  const agent = encryptionData.EncryptionAgent;
  if (!agent.Protocol) {
    throw new Error(
      'Encryption agent protocol version must be present in the encryption data properties.'
    );
  }
  if (agent.Protocol !== azureStorageEncryptionAgentProtocol) {
    throw new Error(
      `Encryption agent value "${agent.EncryptionAgent}" is not recognized or tested with this library.`
    );
  }
  if (!agent.EncryptionAlgorithm) {
    throw new Error(
      'Encryption algorithm type must be present in the encryption data properties.'
    );
  }
  if (!mapDotNetFrameworkToOpenSslAlgorithm.get(agent.EncryptionAlgorithm)) {
    throw new Error(
      `Encryption agent value "${agent.EncryptionAgent}" is not recognized or tested with this library.`
    );
  }
}

function resolveKeyEncryptionKeyFromOptions(
  encryptionOptions,
  keyId,
  callback
) {
  if (!encryptionOptions) {
    return callback(new Error('Encryption options must be specified.'));
  }
  if (!keyId) {
    throw new Error('No key encryption key ID provided.');
  }
  if (
    (!encryptionOptions.keyEncryptionKeys ||
      typeof encryptionOptions.keyEncryptionKeys !== 'object') &&
    (!encryptionOptions.keyResolver ||
      typeof encryptionOptions.keyResolver !== 'function')
  ) {
    return callback(
      new Error(
        'Encryption options must provide either a "keyResolver" function or "keyEncryptionKeys" object.'
      )
    );
  }
  const resolver =
    encryptionOptions.keyResolver ||
    function (keyId, callback) {
      const key = encryptionOptions.keyEncryptionKeys[keyId];
      callback(null, key);
    };
  resolver(keyId, (resolveError, key) => {
    if (resolveError) {
      return callback(resolveError);
    }
    if (!key) {
      return callback(
        new Error(`Could not retrieve a key with identifier "${keyId}".`)
      );
    }
    return callback(null, bufferFromBase64String(key));
  });
}

// Compute Truncated Column Hash:
// Each encrypted entity (row) has its own content encryption key, init vector,
// and then each column is encrypted using an IV that comes from key table
// properties, the row identity and the column name.
function computeTruncatedColumnHash(
  contentEncryptionIV,
  partitionKey,
  rowKey,
  columnName
) {
  // IMPORTANT:
  // The .NET storage library (the reference implementation for Azure client-side
  // storage) has a likely bug in the ordering and concatenation of parameters
  // to generate the truncated column hash; it uses string.Join(partitionKey, rowKey, column)
  // instead of String.Concat. The likely original intention of the author seems
  // to be a concat in the order of partition key, row key, and then column, but
  // instead the resulting string is actually row key, partition key, column,
  // because string.Join treats the first parameter (the partition key in this
  // case) as the separator for joining an array of values. This code uses array
  // join to identically reproduce the .NET behavior here so that the two
  // implementations remain compatible.
  const columnIdentity = Buffer.from(
    [rowKey, columnName].join(partitionKey),
    'utf8'
  );
  const combined = Buffer.concat([contentEncryptionIV, columnIdentity]);
  const hash = getSha256Hash(combined);
  return hash.slice(0, azureStorageContentEncryptionIVBytes);
}

// Buffer/string functions

function base64StringFromBuffer(val) {
  return Buffer.isBuffer(val) ? val.toString('base64') : val;
}

function bufferFromBase64String(val) {
  return Buffer.isBuffer(val) ? val : Buffer.from(val, 'base64');
}

function translateBuffersToBase64(properties) {
  for (const key in properties) {
    if (Buffer.isBuffer(properties[key]['_'])) {
      properties[key] = entityGenerator.String(
        base64StringFromBuffer(properties[key]['_'])
      );
    }
  }
  return properties;
}

// The default encryption resolver implementation: given a list of properties
// to encrypt, return true when that property is being processed.

function createDefaultEncryptionResolver(propertiesToEncrypt) {
  const encryptedKeySet = new Set(propertiesToEncrypt);
  // Default resolver does not use partition/row, but user could
  return (partition, row, name) => {
    return encryptedKeySet.has(name);
  };
}

function encryptProperty(
  contentEncryptionKey,
  contentEncryptionIV,
  partitionKey,
  rowKey,
  property,
  value
) {
  let columnIV = computeTruncatedColumnHash(
    contentEncryptionIV,
    partitionKey,
    rowKey,
    property
  );
  // Store the encrypted properties as binary values on the service instead of
  // base 64 encoded strings because strings are stored as a sequence of WCHARs
  // thereby further reducing the allowed size by half. During retrieve, it is
  // handled by the response parsers correctly even when the service does not
  // return the type for JSON no-metadata.
  return encryptValue(contentEncryptionKey, columnIV, value);
}

function decryptProperty(
  aesAlgorithm,
  contentEncryptionKey,
  contentEncryptionIV,
  partitionKey,
  rowKey,
  propertyName,
  encryptedValue
) {
  const columnIV = computeTruncatedColumnHash(
    contentEncryptionIV,
    partitionKey,
    rowKey,
    propertyName
  );
  return decryptValue(
    aesAlgorithm,
    contentEncryptionKey,
    columnIV,
    bufferFromBase64String(encryptedValue)
  );
}

function encryptProperties(
  encryptionResolver,
  contentEncryptionKey,
  contentEncryptionIV,
  partitionKey,
  rowKey,
  unencryptedProperties,
  callback
) {
  const encryptedProperties = {};
  const encryptedPropertiesList = [];
  if (!unencryptedProperties) {
    return callback(new Error('The entity properties are not set.'));
  }
  const propertyNames = Object.getOwnPropertyNames(unencryptedProperties);
  try {
    for (const property of propertyNames) {
      const wrappedValue = unencryptedProperties[property];
      if (
        property === tableEncryptionKeyDetails ||
        property === tableEncryptionPropertyDetails
      ) {
        return callback(
          new Error(
            'A table encryption property is present in the entity properties to consider for encryption. The property must be removed.'
          )
        );
      }
      if (!wrappedValue['_']) {
        return callback(
          new Error(
            `Property named ${property} is not from a table entity object. Should support ${property}._.`
          )
        );
      }
      const value = wrappedValue['_'];
      if (property === 'PartitionKey' || property === 'RowKey') {
        encryptedProperties[property] = wrappedValue; // {'_': value};
        continue;
      }
      if (property === 'Timestamp') {
        continue;
      }
      if (encryptionResolver(partitionKey, rowKey, property) !== true) {
        encryptedProperties[property] = wrappedValue;
        continue;
      }
      if (value === undefined || value === null) {
        return callback(
          new Error(
            `Null or undefined properties cannot be encrypted. Property in question: ${property}`
          )
        );
      }
      let type = typeof value;
      if (type !== 'string') {
        return callback(
          new Error(
            `${type} properties cannot be encrypted; property in question: ${property}`
          )
        );
      }
      const encryptedValue = encryptProperty(
        contentEncryptionKey,
        contentEncryptionIV,
        partitionKey,
        rowKey,
        property,
        value
      );
      encryptedPropertiesList.push(property);
      encryptedProperties[property] = entityGenerator.Binary(encryptedValue);
    }
  } catch (outerError) {
    return callback(outerError);
  }
  return callback(null, encryptedProperties, encryptedPropertiesList);
}

function decryptProperties(
  allEntityProperties,
  encryptedPropertyNames,
  partitionKey,
  rowKey,
  contentEncryptionKey,
  encryptionData,
  contentEncryptionIV
) {
  validateEncryptionData(encryptionData);
  const aesAlgorithm = openSslFromNetFrameworkAlgorithm(
    encryptionData.EncryptionAgent.EncryptionAlgorithm
  );
  const decryptedProperties = {};
  for (const key in allEntityProperties) {
    if (
      key === tableEncryptionKeyDetails ||
      key === tableEncryptionPropertyDetails
    ) {
      continue;
    }
    if (!encryptedPropertyNames.has(key)) {
      decryptedProperties[key] = allEntityProperties[key];
      continue;
    }
    const innerValue = allEntityProperties[key]['_'];
    const value = decryptProperty(
      aesAlgorithm,
      contentEncryptionKey,
      contentEncryptionIV,
      partitionKey,
      rowKey,
      key,
      innerValue
    );
    decryptedProperties[key] = { _: value.toString('utf8') };
  }
  return decryptedProperties;
}

export function encryptTableEntity(
  partitionKey: string,
  rowKey: string,
  tableEntity: any,
  encryptionOptions: ITableEncryptionOperationOptions
): Promise<any> {
  return new Promise((resolve, reject) => {
    encryptTableEntityCallback(
      partitionKey,
      rowKey,
      tableEntity,
      encryptionOptions,
      (encryptError, encryptedEntity) => {
        return encryptError ? reject(encryptError) : resolve(encryptedEntity);
      }
    );
  });
}

function encryptTableEntityCallback(
  partitionKey: string,
  rowKey: string,
  properties,
  encryptionOptions,
  callback
) {
  if (!partitionKey || !rowKey || !properties) {
    return callback(
      new Error(
        'Must provide a partition key, row key and properties for the entity.'
      )
    );
  }
  const returnBinaryProperties = encryptionOptions.binaryProperties || 'buffer';
  if (
    returnBinaryProperties !== 'base64' &&
    returnBinaryProperties !== 'buffer'
  ) {
    return callback(
      new Error(
        'The binary properties value is not valid. Please provide "buffer" or "base64".'
      )
    );
  }
  const keyEncryptionKeyId = encryptionOptions.keyEncryptionKeyId;
  resolveKeyEncryptionKeyFromOptions(
    encryptionOptions,
    keyEncryptionKeyId,
    (keyLocateError, keyEncryptionKey) => {
      if (keyLocateError) {
        return callback(keyLocateError);
      }
      let encryptionResolver = encryptionOptions.encryptionResolver;
      if (!encryptionResolver) {
        const propertiesToEncrypt = encryptionOptions.encryptedPropertyNames;
        if (!propertiesToEncrypt) {
          return callback(
            new Error(
              'Encryption options must contain either a list of properties to encrypt or an encryption resolver.'
            )
          );
        }
        encryptionResolver = createDefaultEncryptionResolver(
          propertiesToEncrypt
        );
      }
      generateContentEncryptionKey(
        (generateKeyError, contentEncryptionIV, contentEncryptionKey) => {
          if (generateKeyError) {
            return callback(generateKeyError);
          }
          const keyWrappingAlgorithm = azureStorageKeyWrappingAlgorithm;
          wrapContentKey(
            keyWrappingAlgorithm,
            keyEncryptionKey,
            contentEncryptionKey,
            (wrapError, wrappedContentEncryptionKey) => {
              if (wrapError) {
                return callback(wrapError);
              }
              encryptProperties(
                encryptionResolver,
                contentEncryptionKey,
                contentEncryptionIV,
                partitionKey,
                rowKey,
                properties,
                (
                  encryptError,
                  encryptedProperties,
                  encryptionPropertyDetailsSet
                ) => {
                  if (encryptError) {
                    return callback(encryptError);
                  }
                  if (encryptionPropertyDetailsSet.length === 0) {
                    return callback(null, encryptedProperties);
                  }
                  const metadataSerialized = JSON.stringify(
                    encryptionPropertyDetailsSet
                  );
                  encryptedProperties[
                    tableEncryptionPropertyDetails
                  ] = entityGenerator.Binary(
                    encryptProperty(
                      contentEncryptionKey,
                      contentEncryptionIV,
                      partitionKey,
                      rowKey,
                      tableEncryptionPropertyDetails,
                      metadataSerialized
                    )
                  );
                  encryptedProperties[
                    tableEncryptionKeyDetails
                  ] = entityGenerator.String(
                    JSON.stringify(
                      createEncryptionData(
                        keyEncryptionKeyId,
                        jose.util.asBuffer(wrappedContentEncryptionKey),
                        contentEncryptionIV,
                        keyWrappingAlgorithm
                      )
                    )
                  );
                  if (returnBinaryProperties === 'base64') {
                    translateBuffersToBase64(encryptedProperties);
                  }
                  return callback(null, encryptedProperties);
                }
              );
            }
          );
        }
      );
    }
  );
}

export interface ITableEncryptionOperationOptions {
  keyEncryptionKeyId: string;
  keyResolver: any;
  encryptedPropertyNames: Set<string>;
  binaryProperties: string;
}

export function decryptTableEntity(
  partitionKey: string,
  rowKey: string,
  tableEntity: any,
  encryptionOptions: ITableEncryptionOperationOptions
): Promise<any> {
  return new Promise((resolve, reject) => {
    decryptTableEntityCallback(
      partitionKey,
      rowKey,
      tableEntity,
      encryptionOptions,
      (decryptError, decryptedEntity) => {
        return decryptError ? reject(decryptError) : resolve(decryptedEntity);
      }
    );
  });
}

function decryptTableEntityCallback(
  partitionKey: string,
  rowKey: string,
  properties: any,
  encryptionOptions: ITableEncryptionOperationOptions,
  callback
) {
  if (!partitionKey || !rowKey || !properties) {
    return callback(
      new Error('A partition key, row key and properties must be provided.')
    );
  }
  const returnBinaryProperties = encryptionOptions.binaryProperties || 'buffer';
  if (
    returnBinaryProperties !== 'base64' &&
    returnBinaryProperties !== 'buffer'
  ) {
    return callback(
      new Error(
        'The binary properties value is not valid. Please provide "buffer" or "base64".'
      )
    );
  }
  let detailsValue = properties[tableEncryptionKeyDetails]
    ? properties[tableEncryptionKeyDetails]['_']
    : undefined;
  if (detailsValue === undefined) {
    return callback(null, properties);
  }
  let tableEncryptionKey = null;
  try {
    tableEncryptionKey = JSON.parse(detailsValue);
  } catch (parseError) {
    return callback(parseError);
  }
  const iv = bufferFromBase64String(tableEncryptionKey.ContentEncryptionIV);
  const wrappedContentKey = tableEncryptionKey.WrappedContentKey;
  if (wrappedContentKey.Algorithm !== azureStorageKeyWrappingAlgorithm) {
    return callback(
      new Error(
        `The key wrapping algorithm "${wrappedContentKey.Algorithm}" is not tested or supported in this library.`
      )
    );
  }
  const keyWrappingAlgorithm = wrappedContentKey.Algorithm;
  const wrappedContentKeyIdentifier = wrappedContentKey.KeyId;
  const wrappedContentKeyEncryptedKey = bufferFromBase64String(
    wrappedContentKey.EncryptedKey
  );
  const aesAlgorithm = openSslFromNetFrameworkAlgorithm(
    tableEncryptionKey.EncryptionAgent.EncryptionAlgorithm
  );
  resolveKeyEncryptionKeyFromOptions(
    encryptionOptions,
    wrappedContentKeyIdentifier,
    (kvkLocateError, kvk) => {
      if (kvkLocateError) {
        return callback(kvkLocateError);
      }
      const keyEncryptionKeyValue = bufferFromBase64String(kvk);
      unwrapContentKey(
        keyWrappingAlgorithm,
        keyEncryptionKeyValue,
        wrappedContentKeyEncryptedKey,
        (unwrapError, contentEncryptionKey) => {
          if (unwrapError) {
            return callback(unwrapError);
          }
          const metadataIV = computeTruncatedColumnHash(
            iv,
            partitionKey,
            rowKey,
            tableEncryptionPropertyDetails
          );
          const tableEncryptionDetails = bufferFromBase64String(
            properties[tableEncryptionPropertyDetails]['_']
          );
          try {
            const decryptedPropertiesSet = decryptValue(
              aesAlgorithm,
              contentEncryptionKey,
              metadataIV,
              tableEncryptionDetails
            );
            const json = decryptedPropertiesSet.toString('utf8');
            const listOfEncryptedProperties = JSON.parse(json);
            const decrypted = decryptProperties(
              properties,
              new Set(listOfEncryptedProperties),
              partitionKey,
              rowKey,
              contentEncryptionKey,
              tableEncryptionKey,
              iv
            );
            if (returnBinaryProperties === 'base64') {
              translateBuffersToBase64(decrypted);
            }
            return callback(null, decrypted);
          } catch (error) {
            return callback(error);
          }
        }
      );
    }
  );
}
