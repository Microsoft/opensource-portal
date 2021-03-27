//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import serviceBus from 'azure-sb';
import moment from 'moment';

import { IQueueMessage, IQueueProcessor } from '.';
import { IDictionary, Json } from '../../transitional';

const EnqueuedTimeUtc = 'EnqueuedTimeUtc';
const DeliveryCount = 'DeliveryCount';

const PeekLockOption = {
  isPeekLock: true,
};

export interface IServiceBusQueueProcessorOptions {
  queue: string;
  connectionString: string;
}

export class ServiceBusMessage implements IQueueMessage {
  #lockedMessage: serviceBus.Azure.ServiceBus.Message = null;
  constructor(message: serviceBus.Azure.ServiceBus.Message) {
    this.#lockedMessage = message;
    this.brokerProperties = (Object.assign(
      {},
      message
    ) as unknown) as IDictionary<string>;
    if (message.brokerProperties[EnqueuedTimeUtc]) {
      const originalEventQueued = moment(
        message.brokerProperties[EnqueuedTimeUtc]
      );
      const now = moment();
      this.enqueuedSecondsAgo = now.diff(originalEventQueued, 'seconds', true);
    }
    // let deliveryCount = message.brokerProperties[DeliveryCount] !== undefined ? message.brokerProperties[DeliveryCount] : null;
    this.identifier = message.brokerProperties.MessageId;
    this.customProperties = message.customProperties;
    this.unparsedBody = message.body;
    this.body = JSON.parse(message.body);
  }

  body: Json;
  customProperties: IDictionary<string>;
  brokerProperties: IDictionary<string>;
  unparsedBody: any;

  identifier: string;
  enqueuedSecondsAgo?: number;

  lockedMessage() {
    return this.#lockedMessage;
  }
}

export default class ServiceBusQueueProcessor implements IQueueProcessor {
  #options: IServiceBusQueueProcessorOptions;
  #service: serviceBus.ServiceBusService;
  #initialized: boolean;

  constructor(options: IServiceBusQueueProcessorOptions) {
    if (!options.connectionString) {
      throw new Error('options.connectionString required');
    }
    if (!options.queue) {
      throw new Error('options.queue required');
    }
    this.#options = options;
  }

  async initialize(): Promise<void> {
    this.#service = serviceBus.createServiceBusService(
      this.#options.connectionString
    );
    this.#initialized = true;
  }

  receiveMessages(): Promise<ServiceBusMessage[]> {
    if (!this.#initialized) {
      return Promise.reject(new Error('Provider not initialized'));
    }
    return new Promise((resolve, reject) => {
      return this.#service.receiveQueueMessage(
        this.#options.queue,
        PeekLockOption,
        (peekError, lockedMessage) => {
          if (
            (peekError as unknown) === 'No messages to receive' ||
            (!peekError && !lockedMessage)
          ) {
            return resolve([]);
          } else if (peekError) {
            return reject(peekError);
          }
          try {
            const envelope = new ServiceBusMessage(lockedMessage);
            return resolve([envelope]);
          } catch (error) {
            return reject(error);
          }
        }
      );
    });
  }

  async deleteMessage(message: IQueueMessage): Promise<void> {
    if (!this.#initialized) {
      return Promise.reject(new Error('Provider not initialized'));
    }
    return new Promise((resolve, reject) => {
      try {
        const assumedType = message as ServiceBusMessage;
        const lockedMessage = assumedType.lockedMessage();
        return this.#service.deleteMessage(
          lockedMessage,
          (error, deleteResponse) => {
            if (error) {
              console.warn(deleteResponse || error);
            }
            return error ? reject(error) : resolve();
          }
        );
      } catch (deleteError) {
        return reject(deleteError);
      }
    });
  }
}
