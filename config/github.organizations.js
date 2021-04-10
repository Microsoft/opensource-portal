//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

const typescriptConfig = require('./typescript');
const arrayFromString = require('./utils/arrayFromString');

const fs = require('fs');
const path = require('path');

// Resolves the organization configuration data; GITHUB_ORGANIZATIONS_FILE or GITHUB_ORGANIZATIONS_ENVIRONMENT_NAME

const organizationsFileVariableName = 'GITHUB_ORGANIZATIONS_FILE';
const organizationsEnvironmentVariableName = 'GITHUB_ORGANIZATIONS_ENVIRONMENT_NAME';
const organizationsEnvironmentTypeVariableName = 'GITHUB_ORGANIZATIONS_ENVIRONMENT_TYPE_NAME';

const defaultEnvironmentTypeName = 'github.organizations';

function getModuleConfiguration(environmentInstances, environmentType, environmentName) {
  if (!environmentInstances) {
    throw new Error(`${organizationsEnvironmentVariableName} configured but no environment instances were loaded by the painless-config-as-code system`);
  }
  if (!environmentType) {
    throw new Error('No GitHub organizations environment type configured');
  }
  if (!Array.isArray(environmentInstances)) {
    return;
  }
  for (let i = 0; i < environmentInstances.length; i++) {
    let instance = environmentInstances[i];
    try {
      let organizations = instance(environmentName, environmentType);
      if (organizations) {
        return organizations;
      }
    } catch (noProvider) {
      /* The environment does not have the type */
    }
  }
}

module.exports = (graphApi) => {
  const environmentProvider = graphApi.environment;
  const environmentInstances = environmentProvider ? environmentProvider.environmentInstances : null;
  const orgs = [];
  orgs.onboarding = [];
  orgs.ignore = [];
  const defaultLegalEntities = arrayFromString(environmentProvider.get('GITHUB_ORGANIZATIONS_DEFAULT_LEGAL_ENTITIES'));
  const defaultTemplates = arrayFromString(environmentProvider.get('GITHUB_ORGANIZATIONS_DEFAULT_TEMPLATES'));
  const organizationsFile = environmentProvider.get(organizationsFileVariableName);
  const organizationsEnvironmentName = environmentProvider.get(organizationsEnvironmentVariableName);
  const organizationsEnvironmentType = environmentProvider.get(organizationsEnvironmentTypeVariableName) || defaultEnvironmentTypeName;
  if (organizationsFile && organizationsEnvironmentName) {
    console.warn(`GitHub organization loader: Configuration contains both ${organizationsFileVariableName} and ${organizationsEnvironmentVariableName} values. Only the file will be loaded.`);
  }
  let contents = null;
  if (organizationsFile) {
    // This will resolve locally; in this we may want to be able to
    // discover through other mechanisms, too.
    const filename = path.join(typescriptConfig.appDirectory, 'data', organizationsFile);
    try {
      const str = fs.readFileSync(filename, 'utf8');
      contents = JSON.parse(str);
    } catch (notFound) {
      console.warn(`Template definitions could not be loaded from ${filename}: ${notFound.toString()}`);
      throw notFound;
    }
  } else if (organizationsEnvironmentName && !environmentInstances) {
    throw new Error(`${organizationsEnvironmentVariableName} configured but no environment instances were loaded by the painless-config-as-code system`);
  } else if (organizationsEnvironmentName) {
    contents = getModuleConfiguration(environmentInstances, organizationsEnvironmentType, organizationsEnvironmentName);
  }

  if (contents && Array.isArray(contents)) {
    contents.forEach((org) => {
      const isOnboarding = org.onboarding === true;
      const ignore = org.ignore === true;
      if (!org.legalEntities) {
        org.legalEntities = defaultLegalEntities;
      }
      if (!org.templates) {
        org.templates = defaultTemplates;
      }
      let group = orgs;
      if (isOnboarding) {
        group = orgs.onboarding;
      } else if (ignore) {
        group = orgs.ignore;
      }
      group.push(org);
    });
  }

  if (!orgs.onboarding.length) {
    delete orgs.onboarding;
  }
  if (!orgs.ignore.length) {
    delete orgs.ignore;
  }

  return orgs;
};
