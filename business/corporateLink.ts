//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["warn"] }] */

export const CorporatePropertyNames = [
  'isServiceAccount',
  'serviceAccountMail',

  'corporateId',
  'corporateUsername',
  'corporateDisplayName',
  'corporateMailAddress',

  'thirdPartyId',
  'thirdPartyUsername',
  'thirdPartyAvatar',
];

export interface ICorporateLinkProperties {
  isServiceAccount: any;
  serviceAccountMail: any;

  corporateId: any;
  corporateUsername: any;
  corporateDisplayName: any;
  corporateMailAddress: any;
  corporateAlias: any;

  thirdPartyId: any;
  thirdPartyUsername: any;
  thirdPartyAvatar: any;
}

export interface ICorporateLink extends ICorporateLinkProperties {
  isServiceAccount: boolean;
  serviceAccountMail: string;

  corporateId: string;
  corporateUsername: string;
  corporateDisplayName: string;
  corporateMailAddress: string;
  corporateAlias: string;

  thirdPartyId: string;
  thirdPartyUsername: string;
  thirdPartyAvatar: string;
}

export interface ICorporateLinkExtended {
  internal: () => ICorporateLinkExtendedDirectMethods;
}

export interface ICorporateLinkExtendedDirectMethods {
  getDirectEntity: () => any;
}

export function corporateLinkToJson(link: ICorporateLink): ICorporateLink {
  return (
    link && {
      corporateAlias: link.corporateAlias,
      corporateDisplayName: link.corporateDisplayName,
      corporateId: link.corporateId,
      corporateMailAddress: link.corporateMailAddress,
      corporateUsername: link.corporateUsername,
      serviceAccountMail: link.serviceAccountMail,
      isServiceAccount: link.isServiceAccount,
      thirdPartyAvatar: link.thirdPartyAvatar,
      thirdPartyId: link.thirdPartyId,
      thirdPartyUsername: link.thirdPartyUsername,
    }
  );
}
