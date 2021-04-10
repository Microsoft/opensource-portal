//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { TableUtilities } from 'azure-storage';

const entityGenerator = TableUtilities.entityGenerator;

function reduceEntity(instance) {
  if (instance === undefined || instance === null) {
    return instance;
  }
  for (let column in instance) {
    if (instance[column] && instance[column]._ !== undefined) {
      instance[column] = instance[column]._;
    }
  }
  return instance;
}

function createEntity(partitionKey, rowKey, obj?, callback?) {
  if (typeof (obj) === 'function') {
    callback = obj;
    obj = undefined;
  }
  var entity = {
    PartitionKey: entityGenerator.String(partitionKey),
    RowKey: entityGenerator.String(rowKey)
  };
  if (obj) {
    mergeIntoEntity(entity, obj, null);
  }
  if (callback) {
    return callback(null, entity);
  } else {
    return entity;
  }
}

function mergeIntoEntity(entity, obj, callback?) {
  // Pretty legacy code...
  if (obj) {
    for (let key in obj) {
      // Currently stripping metadata
      if (key === '.metadata') {
        continue;
      }
      if (obj[key] === undefined || obj[key] === null) {
        // Skip undefined/null objects, including the key
        continue;
      }
      if (typeof obj[key] === 'string') {
        entity[key] = entityGenerator.String(obj[key]);
      } else if (obj[key] === true) {
        entity[key] = entityGenerator.Boolean(true);
      } else if (obj[key] === false) {
        entity[key] = entityGenerator.Boolean(false);
      } else if (Buffer.isBuffer(obj[key])) {
        entity[key] = entityGenerator.Binary(obj[key]);
      } else if (obj[key] instanceof Date) {
        entity[key] = entityGenerator.DateTime(obj[key]);
      } else if (typeof obj[key] === 'number') {
        // Opinionated entity processing: store all numbers as strings
        entity[key] = entityGenerator.String(obj[key].toString());
      } else {
        console.warn('Consider whether a new entity merge clause is required for key ' + key + ' of type:' + typeof obj[key]);
        if (obj[key].toString) {
          entity[key] = entityGenerator.String(obj[key].toString());
        } else {
          entity[key] = entityGenerator.String(obj[key]);
        }
      }
    }
  }
  if (callback) {
    callback(null, entity);
  } else {
    return entity;
  }
}

export default {
  reduce: reduceEntity,
  create: createEntity,
  merge: mergeIntoEntity,
};
