//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import {
  IEntityMetadata,
  EntityMetadataBase,
  IEntityMetadataBaseOptions,
} from '../../lib/entityMetadataProvider/entityMetadata';
import {
  AuditLogRecord,
  AuditLogRecordQueryUndoCandidatesByThirdPartyId,
  AuditLogRecordQueryRecordsByActorThirdPartyId,
  AuditLogRecordQueryRecordsByUserThirdPartyId,
  AuditLogRecordQueryRecordsByRepositoryId,
  AuditLogRecordQueryRecordsByTeamId,
} from './auditLogRecord';
import { EntityImplementation } from './auditLogRecord';

const thisProviderType = EntityImplementation.Type;

export interface IAuditLogRecordProviderCreateOptions
  extends IEntityMetadataBaseOptions {}

export interface IAuditLogRecordProvider {
  initialize(): Promise<void>;
  getRecord(recordId: string): Promise<AuditLogRecord>;
  insertRecord(record: AuditLogRecord): Promise<string>;

  queryAuditLogForActorThirdPartyId(id: string): Promise<AuditLogRecord[]>;
  queryAuditLogForUserThirdPartyId(id: string): Promise<AuditLogRecord[]>;
  queryAuditLogForThirdPartyIdUndoOperations(
    id: string
  ): Promise<AuditLogRecord[]>;
  queryAuditLogForRepositoryOperations(
    repositoryId: string
  ): Promise<AuditLogRecord[]>;
  queryAuditLogForTeamOperations(teamId: string): Promise<AuditLogRecord[]>;
}

export class AuditLogRecordProvider
  extends EntityMetadataBase
  implements IAuditLogRecordProvider {
  constructor(options: IAuditLogRecordProviderCreateOptions) {
    super(options);
    EntityImplementation.EnsureDefinitions();
  }

  async getRecord(recordId: string): Promise<AuditLogRecord> {
    this.ensureHelpers(thisProviderType);
    let metadata: IEntityMetadata = null;
    metadata = await this._entities.getMetadata(thisProviderType, recordId);
    return this.deserialize<AuditLogRecord>(thisProviderType, metadata);
  }

  async insertRecord(record: AuditLogRecord): Promise<string> {
    const entity = this.serialize(thisProviderType, record);
    await this._entities.setMetadata(entity);
    return entity.entityId;
  }

  async queryAuditLogForRepositoryOperations(
    repositoryId: string
  ): Promise<AuditLogRecord[]> {
    const query = new AuditLogRecordQueryRecordsByRepositoryId(repositoryId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<AuditLogRecord>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryAuditLogForTeamOperations(
    teamId: string
  ): Promise<AuditLogRecord[]> {
    const query = new AuditLogRecordQueryRecordsByTeamId(teamId);
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<AuditLogRecord>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryAuditLogForActorThirdPartyId(
    thirdPartyId: string
  ): Promise<AuditLogRecord[]> {
    const query = new AuditLogRecordQueryRecordsByActorThirdPartyId(
      thirdPartyId
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<AuditLogRecord>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryAuditLogForUserThirdPartyId(
    thirdPartyId: string
  ): Promise<AuditLogRecord[]> {
    const query = new AuditLogRecordQueryRecordsByUserThirdPartyId(
      thirdPartyId
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<AuditLogRecord>(
      thisProviderType,
      metadatas
    );
    return results;
  }

  async queryAuditLogForThirdPartyIdUndoOperations(
    thirdPartyId: string
  ): Promise<AuditLogRecord[]> {
    const query = new AuditLogRecordQueryUndoCandidatesByThirdPartyId(
      thirdPartyId
    );
    const metadatas = await this._entities.fixedQueryMetadata(
      thisProviderType,
      query
    );
    const results = this.deserializeArray<AuditLogRecord>(
      thisProviderType,
      metadatas
    );
    return results;
  }
}
