//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import redis = require('redis');
import { ICacheHelper } from '.';

const debug = require('debug')('redis');
const debugCrossOrganization = require('debug')('redis-cross-org');

const zlib = require('zlib');

const compressionOptions = {
  type: 'gzip',
  params: {
    level: zlib.Z_BEST_SPEED,
  },
};

export interface ISetCompressedOptions {
  minutesToExpire?: number;
}

export interface IRedisHelperOptions {
  redisClient: redis.RedisClient;
  prefix?: string;
}

export default class RedisHelper implements ICacheHelper {
  private _redis: redis.RedisClient;
  private _prefix: string;

  constructor(options: IRedisHelperOptions) {
    if (!options.redisClient) {
      throw new Error('options.redisClient required by RedisHelper');
    }
    this._prefix = options.prefix ? options.prefix + '.' : '';
    this._redis = options.redisClient;
  }

  private key(key: string) {
    return this._prefix + key;
  }

  get(key: string): Promise<string> {
    key = this.key(key);
    debug('GET ' + key);
    if (key.includes('.x#')) {
      debugCrossOrganization('    GET ' + key);
    }
    return new Promise((resolve, reject) => {
      this._redis.get(key, (error, value) => {
        return error ? reject(error) : resolve(value);
      });
    });
  }

  getCompressed(key: string): Promise<string> {
    key = this.key(key);
    debug('GET ' + key);
    if (key.includes('.x#')) {
      debugCrossOrganization('    GET ' + key);
    }
    const bufferKey = Buffer.from(key);
    return new Promise((resolve, reject) => {
      this._redis.get(
        (bufferKey as any) as string /* Buffer */,
        (error, buffer) => {
          if (error) {
            return process.nextTick(reject, error);
          }
          if (buffer === undefined || buffer === null) {
            return process.nextTick(resolve, buffer);
          }
          zlib.gunzip(buffer, (unzipError, unzipped) => {
            // Fallback if there is a data error (i.e. it's not compressed)
            if (unzipError && unzipError.errno === zlib.Z_DATA_ERROR) {
              const originalValue = buffer.toString();
              return process.nextTick(resolve, originalValue);
            } else if (unzipError) {
              return process.nextTick(reject, unzipError);
            }
            try {
              const unzippedValue = unzipped.toString();
              return process.nextTick(resolve, unzippedValue);
            } catch (otherError) {
              return process.nextTick(reject, otherError);
            }
          });
        }
      );
    });
  }

  setCompressed(
    key: string,
    value: string,
    options?: ISetCompressedOptions
  ): Promise<void> {
    key = this.key(key);
    const minutesToExpire = options ? options.minutesToExpire : null;
    if (minutesToExpire) {
      debug(`SET ${key} EX ${minutesToExpire}m`);
    } else {
      debug(`SET ${key}`);
    }
    const val = Buffer.from(value);
    return new Promise((resolve, reject) => {
      zlib.gzip(val, compressionOptions, (gzipError, compressed) => {
        if (gzipError) {
          return reject(gzipError);
        }
        const bufferKey = Buffer.from(key);
        const finalize = (error, ok) => {
          return error ? reject(error) : resolve(ok);
        };
        if (minutesToExpire) {
          this._redis.set(
            (bufferKey as any) as string /* Buffer key type to make TypeScript happy */,
            compressed,
            'EX',
            minutesToExpire * 60,
            finalize
          );
        } else {
          this._redis.set(
            (bufferKey as any) as string /* Buffer key type to make TypeScript happy */,
            compressed,
            finalize
          );
        }
      });
    });
  }

  async getObject(key: string): Promise<any> {
    const value = await this.get(key);
    return JSON.parse(value);
  }

  async getObjectCompressed(key: string): Promise<any> {
    const value = await this.getCompressed(key);
    return JSON.parse(value);
  }

  setObject(key: string, object: any): Promise<void> {
    const json = JSON.stringify(object);
    return this.set(key, json);
  }

  set(key: string, value: string): Promise<void> {
    key = this.key(key);
    debug('SET ' + key);
    return new Promise((resolve, reject) => {
      this._redis.set(key, value, (error, ok) => {
        return error ? reject(error) : resolve();
      });
    });
  }

  setObjectWithExpire(
    key: string,
    object: any,
    minutesToExpire: number
  ): Promise<void> {
    const json = JSON.stringify(object);
    return this.setWithExpire(key, json, minutesToExpire);
  }

  setObjectCompressedWithExpire(
    key: string,
    object: any,
    minutesToExpire: number
  ): Promise<void> {
    const json = JSON.stringify(object);
    return this.setCompressedWithExpire(key, json, minutesToExpire);
  }

  setCompressedWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void> {
    if (!minutesToExpire) {
      throw new Error('No minutes to expiration value');
    }
    const options = { minutesToExpire };
    return this.setCompressed(key, value, options);
  }

  setWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void> {
    if (!minutesToExpire) {
      throw new Error('No minutes to expiration value');
    }
    key = this.key(key);
    debug(`SET ${key} EX ${minutesToExpire}m`);
    return new Promise((resolve, reject) => {
      this._redis.set(key, value, 'EX', minutesToExpire * 60, (error, ok) => {
        // CONSIDER: do they want the return value from Redis here?
        return error ? reject(error) : resolve();
      });
    });
  }

  expire(key: string, minutesToExpire: number): Promise<void> {
    if (!minutesToExpire) {
      throw new Error('No minutes to expiration value');
    }
    key = this.key(key);
    debug(`EXP ${key}`);
    return new Promise((resolve, reject) => {
      this._redis.expire(key, minutesToExpire * 60, (error, ok) => {
        return error ? reject(error) : resolve();
      });
    });
  }

  delete(key: string): Promise<void> {
    key = this.key(key);
    debug('DEL ' + key);
    return new Promise((resolve, reject) => {
      this._redis.del(key, (error, ok) => {
        return error ? reject(error) : resolve();
      });
    });
  }
}
