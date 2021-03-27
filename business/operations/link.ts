//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  ICreatedLinkOutcome,
  ICreateLinkOptions,
  LinkOperationSource,
  SupportedLinkType,
} from '.';
import { ICorporateLink, Operations } from '..';
import getCompanySpecificDeployment from '../../middleware/companySpecificDeployment';
import {
  CreateError,
  ErrorHelper,
  setImmediateAsync,
} from '../../transitional';

export async function linkAccounts(
  operations: Operations,
  options: ICreateLinkOptions
): Promise<ICreatedLinkOutcome> {
  const {
    config,
    linkProvider,
    graphProvider,
    insights,
  } = operations.providers;
  if (!linkProvider) {
    throw CreateError.ServerError('linkProvider required');
  }
  if (config.graph?.require === true && !graphProvider) {
    throw CreateError.ServerError('Graph provider required');
  }
  if (!options.link) {
    throw CreateError.InvalidParameters('options.link required');
  }
  const link = options.link;
  if (
    !options.operationSource ||
    (options.operationSource !== LinkOperationSource.Api &&
      options.operationSource !== LinkOperationSource.Portal)
  ) {
    throw CreateError.InvalidParameters(
      'options.operationSource missing or invalid'
    );
  }
  if (!link.corporateId) {
    throw CreateError.InvalidParameters('options.link.corporateId required');
  }
  if (!link.thirdPartyId) {
    throw CreateError.InvalidParameters('options.link.thirdPartyId required');
  }
  const correlationId = options.correlationId || 'no-correlation-id';
  const insightsOperationsPrefix =
    options.operationSource === LinkOperationSource.Portal ? 'Portal' : 'Api';
  const insightsLinkType = link.isServiceAccount ? 'ServiceAccount' : 'User';
  const insightsPrefix = `${insightsOperationsPrefix}${insightsLinkType}Link`;
  const insightsLinkedMetricName = `${insightsPrefix}s`;
  const insightsAllUpMetricsName = `${insightsLinkType}Links`;

  insights.trackEvent({
    name: `${insightsPrefix}Start`,
    properties: ({ ...link, correlationId } as any) as {
      [key: string]: string;
    },
  });

  if (!options.skipGitHubValidation) {
    const githubAccount = operations.getAccount(link.thirdPartyId);
    try {
      await githubAccount.getDetails();
      link.thirdPartyUsername = githubAccount.login;
      link.thirdPartyAvatar = githubAccount.avatar_url;
    } catch (validateAccountError) {
      throw ErrorHelper.EnsureHasStatus(validateAccountError, 400);
    }
  }

  let mailAddress: string = null;
  if (config.graph?.require === true && !options.skipCorporateValidation) {
    try {
      const corporateInfo = await operations.validateCorporateAccountCanLink(
        link.corporateId
      );
      const corporateAccount = corporateInfo.graphEntry;
      if (!corporateAccount) {
        throw CreateError.NotFound(
          `Corporate ID ${link.corporateId} not found`
        );
      }
      mailAddress = corporateAccount.mail || link.serviceAccountMail;
      link.corporateDisplayName = corporateAccount.displayName;
      link.corporateUsername = corporateAccount.userPrincipalName;
      link.corporateMailAddress = corporateAccount.mail;
      // NOTE: strongly typed to the AAD graph info response right now instead of more generic
      if (corporateAccount.mailNickname) {
        link.corporateAlias = corporateAccount.mailNickname.toLowerCase();
      }
      // Validate that the corporate account can be linked
      if (corporateInfo.type === SupportedLinkType.ServiceAccount) {
        if (!link.serviceAccountMail) {
          throw CreateError.InvalidParameters(
            `Corporate account ${link.corporateUsername} must provide a Service Account e-mail address`
          );
        }
        link.isServiceAccount = true;
      }
    } catch (validateCorporateError) {
      throw ErrorHelper.EnsureHasStatus(validateCorporateError, 400);
    }
  }

  let newLinkId: string = null;
  try {
    newLinkId = await linkProvider.createLink(link);
    const eventData = { ...link, linkId: newLinkId, correlationId };
    insights.trackEvent({
      name: `${insightsPrefix}Created`,
      properties: (eventData as any) as { [key: string]: string },
    });
    insights.trackMetric({ name: insightsLinkedMetricName, value: 1 });
    insights.trackMetric({ name: insightsAllUpMetricsName, value: 1 });
    setImmediateAsync(operations.fireLinkEvent.bind(operations, eventData));
  } catch (createLinkError) {
    if (ErrorHelper.IsConflict(createLinkError)) {
      insights.trackEvent({
        name: `${insightsPrefix}AlreadyLinked`,
        properties: ({ ...link, correlationId } as any) as {
          [key: string]: string;
        },
      });
      throw ErrorHelper.EnsureHasStatus(createLinkError, 409);
    }
    insights.trackException({
      exception: createLinkError,
      properties: ({
        ...link,
        event: `${insightsPrefix}InsertError`,
        correlationId,
      } as any) as { [key: string]: string },
    });
    throw createLinkError;
  }

  if (!options.skipSendingMail) {
    setImmediateAsync(
      sendLinkedAccountMail.bind(
        null,
        operations,
        link,
        mailAddress,
        correlationId,
        false /* do not throw on errors */
      )
    );
  }

  const getApi = `${operations.baseUrl}api/people/links/${newLinkId}`;
  insights.trackEvent({
    name: `${insightsPrefix}End`,
    properties: { newLinkId, getApi },
  });
  return { linkId: newLinkId, resourceLink: getApi };
}

async function sendLinkedAccountMail(
  operations: Operations,
  link: ICorporateLink,
  mailAddress: string | null,
  correlationId: string | null,
  throwIfError: boolean
): Promise<void> {
  const {
    insights,
    mailProvider,
    mailAddressProvider,
    config,
  } = operations.providers;
  if (!mailProvider) {
    return;
  }
  if (!mailAddress && !mailAddressProvider) {
    return;
  }
  if (!mailAddress) {
    try {
      mailAddress = await mailAddressProvider.getAddressFromUpn(
        link.corporateUsername
      );
    } catch (getAddressError) {
      if (throwIfError) {
        throw getAddressError;
      }
      return;
    }
  }
  const to = [mailAddress];
  const toAsString = to.join(', ');
  const mail = {
    to,
    bcc: operations.getLinksNotificationMailAddress(),
    subject: `${link.corporateUsername} linked to ${link.thirdPartyUsername}`,
    correlationId,
    content: undefined,
  };
  const companySpecific = getCompanySpecificDeployment();
  const companySpecificStrings = companySpecific?.strings || {};
  const contentOptions = {
    reason: `One-time record of your action to link your account, sent to: ${toAsString}`,
    headline: `Welcome to GitHub, ${link.thirdPartyUsername}`,
    notification: 'information',
    app: `${config.brand.companyName} GitHub`,
    correlationId,
    docs: config && config.urls ? config.urls.docs : null,
    companyName: config.brand.companyName,
    customStrings: companySpecificStrings,
    link,
  };
  try {
    mail.content = await operations.emailRender('link', contentOptions);
  } catch (renderError) {
    insights.trackException({
      exception: renderError,
      properties: ({
        content: contentOptions,
        eventName: 'LinkMailRenderFailure',
      } as any) as { [key: string]: string },
    });
    if (throwIfError) {
      throw renderError;
    }
    return;
  }
  const customData = {
    content: contentOptions,
    receipt: null,
    eventName: undefined,
  };
  try {
    const receipt = await operations.sendMail(mail);
    insights.trackEvent({
      name: 'LinkMailSuccess',
      properties: (customData as any) as { [key: string]: string },
    });
    customData.receipt = receipt;
  } catch (sendMailError) {
    customData.eventName = 'LinkMailFailure';
    insights.trackException({
      exception: sendMailError,
      properties: (customData as any) as { [key: string]: string },
    });
    if (throwIfError) {
      throw sendMailError;
    }
    return;
  }
}
