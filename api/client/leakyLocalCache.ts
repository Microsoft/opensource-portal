//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ICorporateLink, Operations } from '../../business';

const defaultMinutes = 5;

export default class LeakyLocalCache<K, T> {
  // TODO: use one of many NPMs to do this better and cleanup behind the scenes
  private _map: Map<K, T>;
  private _expires: Map<K, Date>;
  private _expireMs: number;

  constructor(localExpirationMinutes?: number) {
    this._map = new Map();
    this._expires = new Map();
    const minutes = localExpirationMinutes || defaultMinutes;
    this._expireMs = 1000 * 60 * minutes;
  }

  get(key: K): T {
    const expires = this._expires.get(key);
    const now = new Date();
    if (!expires) {
      this._map.delete(key);
      return;
    }
    const value = this._map.get(key);
    if (value === undefined) {
      return;
    }
    if (now < expires) {
      return value;
    }
    this._map.delete(key);
    this._expires.delete(key);
    return;
  }

  set(key: K, value: T) {
    if (value !== undefined) {
      const now = new Date();
      const expires = new Date(now.getTime() + this._expireMs);
      this._map.set(key, value);
      this._expires.set(key, expires);
    }
  }
}

export const leakyLocalCacheLinks = new LeakyLocalCache<
  boolean,
  ICorporateLink[]
>();

export async function getLinksLightCache(operations: Operations) {
  const cached = leakyLocalCacheLinks.get(true);
  if (cached) {
    return cached;
  }
  const links = await operations.getLinks();
  leakyLocalCacheLinks.set(true, links);
  return links;
}
