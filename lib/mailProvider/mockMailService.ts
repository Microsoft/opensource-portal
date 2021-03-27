//
// Copyright (c) Microsoft.
//

import { v4 as uuidV4 } from 'uuid';
import { IMail, IMailProvider } from '.';

export default class MockMailService implements IMailProvider {
  html: true;
  private sentMessages = [];
  customServiceConfig: any;
  appVersion: any;

  constructor(config) {
    this.customServiceConfig = config.mail.customService;
    if (this.customServiceConfig.version !== 'prototype') {
      throw new Error(
        `The custom mail service version "${this.customServiceConfig.version}" is not supported`
      );
    }
    this.appVersion = config.logging.version;
  }

  async initialize() {}

  get info(): string {
    return `mockMailService-${this.customServiceConfig.version} v${this.appVersion}`;
  }

  async sendMail(mail: IMail): Promise<any> {
    const receipt = Object.assign(
      {
        id: uuidV4(),
      },
      mail
    );
    this.sentMessages.push(receipt);
    return receipt.id;
  }

  getSentMessages() {
    return this.sentMessages;
  }
}
