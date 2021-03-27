//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  Account,
  ICachedEmployeeInformation,
  IUnlinkMailStatus,
  Operations,
  UnlinkPurpose,
} from '..';

export async function sendTerminatedAccountMail(
  operations: Operations,
  account: Account,
  purpose: UnlinkPurpose,
  details: string[],
  errorsCount: number
): Promise<IUnlinkMailStatus> {
  if (
    !operations.providers.mailProvider ||
    !account.link ||
    !account.link.corporateId
  ) {
    return null;
  }

  purpose = purpose || UnlinkPurpose.Unknown;

  let errorMode = errorsCount > 0;

  let operationsMail = operations.getLinksNotificationMailAddress();
  if (!operationsMail && errorMode) {
    return;
  }
  let operationsArray = operationsMail.split(',');

  let cachedEmployeeManagementInfo: ICachedEmployeeInformation = null;
  let displayName =
    account.link.corporateDisplayName ||
    account.link.corporateUsername ||
    account.link.corporateId;
  let upn = account.link.corporateUsername || account.link.corporateId;
  try {
    cachedEmployeeManagementInfo = await operations.getCachedEmployeeManagementInformation(
      account.link.corporateId
    );
    if (
      !cachedEmployeeManagementInfo ||
      !cachedEmployeeManagementInfo.managerMail
    ) {
      cachedEmployeeManagementInfo = {
        id: account.link.corporateId,
        displayName,
        userPrincipalName: upn,
        managerDisplayName: null,
        managerId: null,
        managerMail: null,
      };
      throw new Error(
        `No manager e-mail address or information retrieved from a previous cache for corporate user ID ${account.link.corporateId}`
      );
    }
    if (cachedEmployeeManagementInfo.displayName) {
      displayName = cachedEmployeeManagementInfo.displayName;
    }
    if (cachedEmployeeManagementInfo.userPrincipalName) {
      upn = cachedEmployeeManagementInfo.userPrincipalName;
    }
  } catch (getEmployeeInfoError) {
    errorMode = true;
    details.push(getEmployeeInfoError.toString());
  }

  let to: string[] = [];
  if (errorMode) {
    to.push(...operationsArray);
  } else {
    to.push(cachedEmployeeManagementInfo.managerMail);
  }
  to = to.filter((val) => val);
  const bcc = [];
  if (!errorMode) {
    bcc.push(...operationsArray);
  }
  const toAsString = to.join(', ');

  let subjectPrefix = '';
  let subjectSuffix = '';
  let headline = `${displayName} has been unlinked from GitHub`;
  switch (purpose) {
    case UnlinkPurpose.Self:
      headline = `${displayName} unlinked themselves from GitHub`;
      subjectPrefix = 'FYI: ';
      subjectSuffix = ' [self-service remove]';
      break;
    case UnlinkPurpose.Deleted:
      subjectPrefix = 'FYI: ';
      subjectSuffix = '[account deleted]';
      headline = `${displayName} deleted their GitHub account`;
      break;
    case UnlinkPurpose.Operations:
      subjectPrefix = 'FYI: ';
      subjectSuffix = ' [corporate GitHub operations]';
      break;
    case UnlinkPurpose.Termination:
      subjectPrefix = '[UNLINKED] ';
      headline = `${displayName} may not be an active employee`;
      break;
    case UnlinkPurpose.Unknown:
    default:
      subjectSuffix = ' [unknown]';
      break;
  }
  const config = operations.providers.config;
  const mail = {
    to,
    bcc,
    subject: `${subjectPrefix}${
      upn || displayName
    } unlinked from GitHub ${subjectSuffix}`.trim(),
    category: ['link'],
    content: undefined,
  };
  mail.content = await operations.emailRender('managerunlink', {
    reason: `As a manager you receive one-time security-related messages regarding your direct reports who have linked their GitHub account to the company.
              This mail was sent to: ${toAsString}`,
    headline,
    notification: 'information',
    app: `${config.brand.companyName} GitHub`,
    link: account.link,
    companyName: config.brand.companyName,
    managementInformation: cachedEmployeeManagementInfo,
    purpose,
    details,
  });
  return {
    to,
    bcc,
    receipt: await operations.sendMail(Object.assign(mail)),
  };
}
