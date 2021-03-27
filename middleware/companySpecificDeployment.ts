//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import path from 'path';
import {
  ICompanySpecificStartup,
  ICompanySpecificStartupProperties,
} from '../interfaces';

import AppPackage from '../package.json';

let instance = null;

function getCompanySpecificDeploymentName() {
  const asAny = AppPackage as any;
  if (asAny?.companySpecificDeploymentDirectory) {
    return asAny.companySpecificDeploymentDirectory as string;
  }
}

function getCompanySpecificDeployment(): ICompanySpecificStartup {
  const name = getCompanySpecificDeploymentName();
  if (!name) {
    return null;
  }
  if (instance) {
    return instance;
  }
  try {
    const pn = path.join(__dirname, '..', name);
    const dynamicInclude = require(pn);
    const entrypoint =
      dynamicInclude && dynamicInclude.default
        ? dynamicInclude.default
        : dynamicInclude;
    if (!(entrypoint as ICompanySpecificStartupProperties).isCompanySpecific) {
      throw new Error(
        `The ${name} company-specific call did not include the 'isCompanySpecific' moniker. Check for circular dependencies: ${pn}`
      );
    }
    instance = entrypoint;
    return instance;
  } catch (includeError) {
    throw includeError;
  }
}

export default getCompanySpecificDeployment;
