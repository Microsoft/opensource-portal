//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IProviders } from '../transitional';
import {
  IAttachCompanySpecificRoutes,
  IAttachCompanySpecificMiddleware,
  ICorporationAdministrationSection,
  IAttachCompanySpecificStrings,
  ICompanySpecificFeatures,
  IAttachCompanySpecificViews,
} from './companySpecific';

export * from './companySpecific';

// We're great at long variable names!

export interface ICompanySpecificStartupProperties {
  isCompanySpecific: true;
  routes?: IAttachCompanySpecificRoutes;
  middleware?: IAttachCompanySpecificMiddleware;
  administrationSection?: ICorporationAdministrationSection;
  strings?: IAttachCompanySpecificStrings;
  features?: ICompanySpecificFeatures;
  views?: IAttachCompanySpecificViews;
}

export type ICompanySpecificStartupFunction = (
  config: any,
  p: IProviders,
  rootdir: string
) => Promise<void>;

export type ICompanySpecificStartup = ICompanySpecificStartupFunction &
  ICompanySpecificStartupProperties;
