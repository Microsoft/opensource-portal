//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

const newsPackageName = 'HOMEPAGE_NEWS_PACKAGE';
const newsCountName = 'HOMEPAGE_NEWS_COUNT';

const defaultCount = 5;

module.exports = function (graphApi) {
  const environmentProvider = graphApi.environment;
  const newsPackageValue = environmentProvider.get(newsPackageName);
  const count = parseInt(
    environmentProvider.get(newsCountName) || defaultCount,
    10
  );

  let articles = [];

  try {
    if (newsPackageValue) {
      const readArticles = require(newsPackageValue);
      if (readArticles && Array.isArray(readArticles)) {
        articles = readArticles;
      }
    }
  } catch (newsLoadError) {
    console.dir(newsLoadError); // silent
  }

  return {
    all: articles,
    homepage: articles.slice(0, count),
  };
};
