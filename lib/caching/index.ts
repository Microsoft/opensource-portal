//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

export interface ICacheHelper {
  get(key: string): Promise<string>;
  getCompressed(key: string): Promise<string>;
  getObject(key: string): Promise<any>;
  getObjectCompressed(key: string): Promise<any>;
  set(key: string, value: string): Promise<void>;
  setObject(key: string, value: any): Promise<void>;
  setObjectWithExpire(
    key: string,
    value: any,
    minutesToExpire: number
  ): Promise<void>;
  setObjectCompressedWithExpire(
    key: string,
    value: any,
    minutesToExpire: number
  ): Promise<void>;
  setCompressed(key: string, value: string): Promise<void>;
  setCompressedWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void>;
  setWithExpire(
    key: string,
    value: string,
    minutesToExpire: number
  ): Promise<void>;
  expire(key: string, minutesToExpire: number): Promise<void>;
  delete(key: string): Promise<void>;
}
