//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import MockMailService from './mockMailService';
import SmtpMailService from './smtpMailService';
import getCompanySpecificDeployment from '../../middleware/companySpecificDeployment';

export interface IMail {
  from?: string;
  to: string | string[];
  cc?: string | string[];
  bcc?: string | string[];
  content?: any;
  subject?: string;
  category?: string[];
  correlationId?: string;
  senderProfile?: string;
}

export interface IMailProvider {
  info: string;
  sendMail(mail: IMail): Promise<any>;
  html: boolean;
  getSentMessages(): any[];
  initialize(): Promise<string | void>;
}

function patchOverride(provider, newToAddress, htmlOrNot) {
  const sendMail = provider.sendMail.bind(provider);
  provider.sendMail = (mailOptions: IMail): Promise<any> => {
    let originalTo = mailOptions.to;
    if (typeof originalTo !== 'string' && originalTo.join) {
      originalTo = originalTo.join(', ');
    }
    if (!mailOptions.content) {
      mailOptions.content = '';
    }
    mailOptions.to = newToAddress;
    if (mailOptions.cc) {
      if (typeof mailOptions.cc === 'string') {
        mailOptions.cc = [mailOptions.cc];
      }
      if (Array.isArray(mailOptions.cc) && mailOptions.cc.length) {
        originalTo += ` (CC: ${mailOptions.cc.join(', ')})`;
        mailOptions.cc = null;
      }
    }
    if (mailOptions.bcc) {
      if (typeof mailOptions.bcc === 'string') {
        mailOptions.bcc = [mailOptions.bcc];
      }
      if (Array.isArray(mailOptions.bcc) && mailOptions.bcc.length) {
        originalTo += ` (BCC: ${mailOptions.bcc.join(', ')})`;
        mailOptions.bcc = null;
      }
    }
    const initialContent = mailOptions.content;
    const redirectMessage = `This mail was intended for ${originalTo} but was instead sent to ${newToAddress} per a configuration override.\n`;
    mailOptions.content = htmlOrNot
      ? `<p><em>${redirectMessage}</em></p>\n${initialContent}`
      : `${redirectMessage}\n${initialContent}`;
    return sendMail(mailOptions);
  };
  return provider;
}

export function createMailProviderInstance(config): IMailProvider {
  const deployment = getCompanySpecificDeployment();
  let mailProvider: IMailProvider = null;
  if (deployment?.features?.mailProvider?.tryCreateInstance) {
    mailProvider = deployment.features.mailProvider.tryCreateInstance(config);
    if (mailProvider) {
      return mailProvider;
    }
  }

  const mailConfig = config.mail;
  if (mailConfig === undefined) {
    return;
  }
  const provider = mailConfig.provider;
  if (!provider) {
    return;
  }
  switch (provider) {
    case 'smtpMailService': {
      mailProvider = new SmtpMailService(config);
      break;
    }
    case 'mockMailService': {
      mailProvider = new MockMailService(config);
      break;
    }
    default: {
      throw new Error(
        `The mail provider "${mailConfig.provider}" is not implemented or configured at this time.`
      );
    }
  }
  if (mailConfig.overrideRecipient) {
    patchOverride(
      mailProvider,
      mailConfig.overrideRecipient,
      mailProvider.html
    );
  }
  return mailProvider;
}
