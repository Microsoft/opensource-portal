//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  QueueServiceClient,
  QueueClient,
  DequeuedMessageItem,
} from '@azure/storage-queue';

import { IQueueMessage, IQueueProcessor } from '.';
import { Json, IDictionary } from '../../transitional';

export interface IAzureQueuesProcessorOptions {
  account: string;
  queue: string;
  sas: string;
}

export class AzureQueuesMessage implements IQueueMessage {
  constructor(message: DequeuedMessageItem) {
    this.popReceipt = message.popReceipt;
    this.identifier = message.messageId;
    this.unparsedBody = new Buffer(message.messageText, 'base64').toString(
      'utf8'
    );
    const parsed = JSON.parse(this.unparsedBody);
    if (parsed && parsed.body && typeof (parsed.body === 'string')) {
      // our own envelope format designed to work well with Azure Logic Apps
      this.unparsedBody = parsed.body;
      delete parsed.body;
      this.customProperties = parsed;
      this.body = JSON.parse(this.unparsedBody);
      this.wasEnveloped = true;
    } else {
      this.body = parsed;
      this.customProperties = {};
      this.wasEnveloped = false;
    }
  }

  body: Json;
  unparsedBody: string;
  identifier: string;
  customProperties: IDictionary<string>;
  popReceipt: string;
  wasEnveloped: boolean;
}

export default class AzureQueuesProcessor implements IQueueProcessor {
  #queueClient: QueueClient = null;
  #options: IAzureQueuesProcessorOptions = null;
  #initialized: boolean;

  constructor(options: IAzureQueuesProcessorOptions) {
    if (!options.account) {
      throw new Error('options.account required');
    }
    if (!options.sas) {
      throw new Error('options.sas required');
    }
    if (!options.queue) {
      throw new Error('options.queue required');
    }
    this.#options = options;
  }

  async initialize(): Promise<void> {
    const { account, sas, queue } = this.#options;
    const client = new QueueServiceClient(
      `https://${account}.queue.core.windows.net${sas}`
    );
    this.#queueClient = client.getQueueClient(queue);
    this.#initialized = true;
  }

  async receiveMessages(): Promise<AzureQueuesMessage[]> {
    this.requireInitialized();
    const response = await this.#queueClient.receiveMessages();
    if (response.receivedMessageItems.length > 0) {
      return response.receivedMessageItems.map(
        (message) => new AzureQueuesMessage(message)
      );
    }
    return [];
  }

  async deleteMessage(message: IQueueMessage): Promise<void> {
    this.requireInitialized();
    const assumedType = message as AzureQueuesMessage;
    if (!assumedType.popReceipt) {
      throw new Error(
        'Message must be of type AzureQueuesMessage and have a property popReceipt.'
      );
    }
    const deleteMessageResponse = await this.#queueClient.deleteMessage(
      assumedType.identifier,
      assumedType.popReceipt
    );
    // console.log(deleteMessageResponse);
  }

  private requireInitialized() {
    if (!this.#initialized) {
      throw new Error('Client must be initialized before use.');
    }
  }
}
