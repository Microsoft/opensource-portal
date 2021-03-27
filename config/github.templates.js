//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

const fs = require('fs');
const path = require('path');

const typescriptConfig = require('./typescript');
const arrayFromString = require('./utils/arrayFromString');

// GITHUB_ORGANIZATIONS_TEMPLATES_TYPE: 'npm' or 'fs', 'fs' default
// GITHUB_ORGANIZATIONS_TEMPLATES_RELATIVE_DIRECTORY: relative to app dir, defaults to 'data/templates'
// GITHUB_ORGANIZATIONS_TEMPLATES_PACKAGE_NAME: npm package name if type mode is 'npm', no default

// GITHUB_ORGANIZATIONS_DEFAULT_TEMPLATES

module.exports = (graphApi) => {
  const environmentProvider = graphApi.environment;
  const configurationEnvironmentName = environmentProvider.get(
    'CONFIGURATION_ENVIRONMENT'
  );
  const defaultDirectory = path.join('data', 'templates');

  // 'npm' or 'fs'
  let templateSourceType =
    environmentProvider.get('GITHUB_ORGANIZATIONS_TEMPLATES_TYPE') || 'fs';
  if (templateSourceType !== 'npm' && templateSourceType !== 'fs') {
    console.warn(
      `GITHUB_ORGANIZATIONS_TEMPLATES_TYPE must be either 'fs' or 'npm', defaulting to 'fs' as '${templateSourceType} is unrecognized`
    );
    templateSourceType = 'fs';
  }

  let templates = {
    directory:
      templateSourceType === 'fs'
        ? environmentProvider.get(
            'GITHUB_ORGANIZATIONS_TEMPLATES_RELATIVE_DIRECTORY'
          ) || defaultDirectory
        : null,
    definitions: null,
    defaultTemplates: arrayFromString(
      environmentProvider.get('GITHUB_ORGANIZATIONS_DEFAULT_TEMPLATES') || ''
    ),
  };

  if (templateSourceType === 'fs') {
    templates.directory = path.join(
      typescriptConfig.appDirectory,
      templates.directory
    );
    try {
      const filename = path.join(templates.directory, 'definitions.json');
      const str = fs.readFileSync(filename, 'utf8');
      templates.definitions = JSON.parse(str);
    } catch (notFound) {
      console.warn(
        `Template definitions could not be loaded from ${filename}: ${notFound.toString()}`
      );
      throw notFound;
    }
  } else if (templateSourceType === 'npm') {
    const npmName = environmentProvider.get(
      'GITHUB_ORGANIZATIONS_TEMPLATES_PACKAGE_NAME'
    );
    if (!npmName) {
      throw new Error(
        "When GITHUB_ORGANIZATIONS_TEMPLATES_TYPE is set to 'npm', GITHUB_ORGANIZATIONS_TEMPLATES_PACKAGE_NAME must be set"
      );
    }
    try {
      const templatePackageData = require(npmName);
      if (!templatePackageData || typeof templatePackageData !== 'object') {
        throw new Error(`${npmName} did not return data or an object`);
      }
      if (!templatePackageData.directory) {
        throw new Error(
          `${npmName} did not return an object with a 'directory' property`
        );
      }
      if (!templatePackageData.definitions) {
        throw new Error(
          `${npmName} did not return an object with a 'definitions' property`
        );
      }
      templates = templatePackageData;
    } catch (templatesNpmLoadError) {
      const combinedError = new Error(
        `Trouble loading npm package ${npmName} as configured in GITHUB_ORGANIZATIONS_TEMPLATES_PACKAGE_NAME: ${templatesNpmLoadError.toString()}`
      );
      combinedError.innerError = templatesNpmLoadError;
      throw combinedError;
    }
  }

  if (templates.definitions) {
    const clonedDefinitions = Object.assign({}, templates.definitions);
    const names = Object.getOwnPropertyNames(clonedDefinitions);
    for (let i = 0; i < names.length; i++) {
      const definition = clonedDefinitions[names[i]];
      if (
        !definition.environments ||
        definition.environments.includes(configurationEnvironmentName)
      ) {
        // keep this template around
      } else {
        delete clonedDefinitions[names[i]];
      }
    }
    templates.definitions = clonedDefinitions;
  }

  return templates;
};
