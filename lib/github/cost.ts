//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

function createCostContainer() {
  return {
    local: {
      cacheHits: 0,
      cacheMisses: 0,
    },
    redis: {
      cacheHit: 0,
      cacheMisses: 0,
      getCalls: 0,
      setCalls: 0,
      expireCalls: 0,
    },
    github: {
      restApiCalls: 0,
      usedApiTokens: 0,
      cacheHits: 0,
      remainingApiTokens: 0,
    },
  };
}

function addCosts(a, b) {
  if (a && b) {
    a.local.cacheHits += b.local.cacheHits;
    a.local.cacheMisses += b.local.cacheMisses;
    a.redis.cacheHit += b.redis.cacheHit;
    a.redis.cacheMisses += b.redis.cacheMisses;
    a.redis.getCalls += b.redis.getCalls;
    a.redis.setCalls += b.redis.setCalls;
    a.redis.expireCalls += b.redis.expireCalls;
    a.github.restApiCalls += b.github.restApiCalls;
    a.github.usedApiTokens += b.github.usedApiTokens;
    a.github.cacheHits += b.github.cacheHits;

    // Min; though if the refresh happens in the middle this will be off
    if (b.github.remainingApiTokens > 0) {
      let floor =
        a.github.remainingApiTokens <= 0
          ? b.github.remainingApiTokens
          : a.github.remainingApiTokens;
      a.github.remainingApiTokens = Math.min(
        floor,
        b.github.remainingApiTokens
      );
    }
  }
  return a;
}

export default {
  create: createCostContainer,
  add: addCosts,
};
