//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

// Config hack to be able to understand whether the runtime environment
// is hosted within a 'dist' folder or not.

const path = require('path');

const entryPoint = process.argv[1];
const hasDist =
  entryPoint.includes('/dist/') || entryPoint.includes('\\dist\\');

const aboveConfig = path.join(__dirname, '..');

// console.log(`TypeScript entrypoint: ${entryPoint}, app dir: ${aboveConfig}, dist: ${hasDist}`);

module.exports = {
  dist: hasDist,
  appDirectory: aboveConfig,
};
