//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { ICorporateLink } from '../interfaces';

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

export function corporateLinkToJson(link: ICorporateLink): ICorporateLink {
  return link && {
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
  };
}
