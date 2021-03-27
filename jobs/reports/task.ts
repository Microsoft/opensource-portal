//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/* eslint no-console: ["error", { allow: ["warn", "dir", "log"] }] */

// The reporting system in use here is best described as a "hairball" or snowball design
// as an initial approach. A large, informative set of contextual data is built by the
// report providers. It is OK for a provider in the pipeline to depend on the data
// collected before its execution.

import azure from 'azure-storage';
import os from 'os';
import fileSize from 'file-size';
import moment from 'moment-timezone';
import path from 'path';

import app, { IReposJob, IReposJobResult } from '../../app';

// import { buildConsolidatedMap as buildRecipientMap } from './consolidated';

import {
  build as organizationsBuild,
  consolidate as organizationsConsoldate,
  process as organizationsProcess,
} from './organizations';
import {
  build as repositoriesBuild,
  consolidate as repositoriesConsolidate,
  process as repositoriesProcess,
} from './repositories';
import {
  build as teamsBuild,
  consolidate as teamsConsolidate,
  process as teamsProcess,
  IReportsTeamContext,
} from './teams';

import mailer from './mailer';

import { Operations } from '../../business/operations';
import { ICacheHelper } from '../../lib/caching';
import { ICorporateLink } from '../../business/corporateLink';
import { writeTextToFile } from '../../utils';
import { Repository } from '../../business/repository';
import { Team } from '../../business/team';
import { writeDeflatedTextFile } from './fileCompression';

// Debug-related values for convienience
const fakeSend = false;
const skipStore = false;
const slice = undefined; // 250;

const reportGeneratedFormat = 'h:mm a dddd, MMMM Do YYYY';

const providerNames = ['organizations', 'repositories', 'teams'];

export interface IReportsContext {
  operations: Operations;
  insights: any;
  entities?: {
    repos?: Repository[];
    teams?: Team[];
  };
  processing: any;
  reportsBy: {
    upn: Map<string, any>;
    email: Map<string, any>;
  };
  providers: any;
  started: string;
  organizationData: any;
  teamData?: IReportsTeamContext[];
  reports: {
    reportRedisClient: ICacheHelper;
    send: boolean;
    store: boolean;
    dataLake: boolean;
  };
  consolidated?: any;
  visitedDefinitions: any;
  config: any;
  app: any;
  linkData?: Map<number, ICorporateLink>;
  repositoryData?: any[];
  reportsByRecipient?: Map<string, any>;
  settings: {
    basedir: string;
    slice?: any;
    parallelRepoProcessing: number;
    repoDelayAfter: number;
    teamDelayAfter: number;
    tooManyOrgOwners: number;
    tooManyRepoAdministrators: number;
    orgPercentAvailablePrivateRepos: number;
    fakeSend?: string;
    storeLocalReportPath?: string;
    witnessEventKey: string;
    witnessEventReportsTimeToLiveMinutes: any; // ?
    consolidatedSchemaVersion: string;
    fromAddress: string;
    dataLakeAccount?: any; // ?
    campaign: {
      source: 'administrator-digest';
      medium: 'email';
      campaign: 'github-digests';
    };
  };
}

async function buildReport(context): Promise<void> {
  try {
    await processReports(context);
  } catch (error) {
    console.dir(error);
  }
  try {
    await buildReports(context);
  } catch (error) {
    console.dir(error);
  }
  try {
    await consolidateReports(context);
  } catch (error) {
    console.dir(error);
  }
  try {
    await storeReports(context);
  } catch (error) {
    console.dir(error);
  }
  try {
    await sendReports(context);
  } catch (error) {
    console.dir(error);
  }
  try {
    await recordMetrics(context);
  } catch (error) {
    console.dir(error);
  }
  try {
    await dataLakeUpload(context);
  } catch (error) {
    console.dir(error);
  }
  try {
    await finalizeEvents(context);
  } catch (error) {
    console.dir(error);
  }
}

export default async function run({
  providers,
  started,
}: IReposJob): Promise<IReposJobResult> {
  const { mailProvider, operations, config } = providers;
  const okToContinue =
    config &&
    config.github &&
    config.github.jobs &&
    config.github.jobs.reports &&
    config.github.jobs.reports.enabled === true;
  if (!okToContinue) {
    console.log('config.github.jobs.reports.enabled is not set');
    return {};
  }

  console.log('OK, so, this job is actually not setup to work right now...');
  return {};

  // -- THIS JOB IS OFFLINE FOR NOW --

  console.log(`Report run started ${started}`);

  const insights = providers.insights;
  if (!insights) {
    throw new Error('No app insights client available');
  }
  insights.trackEvent({
    name: 'JobReportsStarted',
    properties: {
      hostname: os.hostname(),
    },
  });
  if (!mailProvider) {
    throw new Error('No mail provider available');
  }
  const reportConfig =
    config && config.github && config.github.jobs
      ? config.github.jobs.reports
      : {};
  const context: IReportsContext = {
    providers,
    operations,
    insights,
    entities: {},
    processing: {},
    reportsBy: {
      upn: new Map(),
      email: new Map(),
    },
    started: moment().format(),
    organizationData: {},
    settings: {
      basedir: config.typescript.appDirectory,
      slice: slice || undefined,
      parallelRepoProcessing: 2,
      repoDelayAfter: 200, // 200ms to wait between repo actions, to help reduce GitHub load
      teamDelayAfter: 200, // 200ms to wait between team actions, to help reduce GitHub load
      tooManyOrgOwners: 5,
      tooManyRepoAdministrators: 15,
      orgPercentAvailablePrivateRepos: 0.15,
      fakeSend: fakeSend ? path.join(__dirname, 'sent') : undefined,
      storeLocalReportPath: skipStore
        ? path.join(__dirname, 'report.json')
        : undefined,
      witnessEventKey: reportConfig.witnessEventKey,
      witnessEventReportsTimeToLiveMinutes:
        reportConfig.witnessEventReportsTimeToLiveMinutes,
      consolidatedSchemaVersion: '170503',
      fromAddress: reportConfig.mail.from,
      dataLakeAccount: null,
      campaign: {
        source: 'administrator-digest',
        medium: 'email',
        campaign: 'github-digests',
      },
    },
    reports: {
      reportRedisClient: null, // reportRedisClient,
      send:
        true && (fakeSend || (reportConfig.mail && reportConfig.mail.enabled)),
      store: true && !skipStore,
      dataLake:
        true &&
        !skipStore &&
        reportConfig.dataLake &&
        reportConfig.dataLake.enabled,
    },
    visitedDefinitions: {},
    consolidated: {},
    config,
    app,
  };
  if (
    context.reports.dataLake === true &&
    reportConfig.dataLake &&
    reportConfig.dataLake.azureStorage &&
    reportConfig.dataLake.azureStorage.key
  ) {
    context.settings.dataLakeAccount = reportConfig.dataLake.azureStorage;
  }
  await buildReport(context);
  console.log('reporting done');
  return {};
}

// ------------------------------------------------------------------

async function buildReports(context) {
  try {
    await organizationsBuild(context);
  } catch (globalBuildError) {
    console.dir(globalBuildError);
  }
  try {
    await repositoriesBuild(context);
  } catch (globalBuildError) {
    console.dir(globalBuildError);
  }
  try {
    await teamsBuild(context);
  } catch (globalBuildError) {
    console.dir(globalBuildError);
  }
  return context;
}

async function processReports(context) {
  try {
    await organizationsProcess(context);
  } catch (globalProcessError) {
    console.dir(globalProcessError);
  }
  try {
    await repositoriesProcess(context);
  } catch (globalProcessError) {
    console.dir(globalProcessError);
  }
  try {
    await teamsProcess(context);
  } catch (globalProcessError) {
    console.dir(globalProcessError);
  }
  return context;
}

async function consolidateReports(
  context: IReportsContext
): Promise<IReportsContext> {
  try {
    await organizationsConsoldate(context);
  } catch (globalConsolidationError) {
    console.dir(globalConsolidationError);
  }
  try {
    await repositoriesConsolidate(context);
  } catch (globalConsolidationError) {
    console.dir(globalConsolidationError);
  }
  try {
    await teamsConsolidate(context);
  } catch (globalConsolidationError) {
    console.dir(globalConsolidationError);
  }
  return context;
}

// ------------------------------------------------------------------

async function finalizeEvents(context: IReportsContext) {
  context.insights.trackEvent({ name: 'JobReportsFinalizing' });
  return context;
}

async function dataLakeUpload(context: IReportsContext) {
  const insights = context.insights;
  if (!context.reports.dataLake) {
    insights.trackEvent({ name: 'JobReportsReportDataLakeSkipped' });
    return context;
  }
  insights.trackEvent({ name: 'JobReportsReportDataLakeStarted' });

  // specific properties used for each row
  // issueProviderName:
  // issueTimestamp: started
  // issueTypeName: typeName

  let dataLakeOutput = [];

  const started = context.started;
  const consolidated = context.consolidated;
  for (let i = 0; i < providerNames.length; i++) {
    const providerName = providerNames[i];
    const root = consolidated[providerName];
    if (root) {
      const definitions = {};
      for (let x = 0; x < root.definitions.length; x++) {
        const d = root.definitions[x];
        definitions[d.name] = d;
      }
      if (root.entities) {
        for (let j = 0; j < root.entities.length; j++) {
          const entity = root.entities[j];
          if (entity && entity.issues) {
            const issueList = Object.getOwnPropertyNames(entity.issues);
            for (let k = 0; k < issueList.length; k++) {
              const issueTypeName = issueList[k];
              const issues = entity.issues[issueTypeName];
              const definition = definitions[issueTypeName];
              let targetCollectionName = null;
              if (definition && definition.hasTable) {
                targetCollectionName = 'rows';
              } else if (definition && definition.hasList) {
                targetCollectionName = 'listItems';
              }
              if (
                targetCollectionName &&
                issues[targetCollectionName] &&
                Array.isArray(issues[targetCollectionName])
              ) {
                const collection = issues[targetCollectionName];
                for (let l = 0; l < collection.length; l++) {
                  const row = collection[l];
                  const rowValue =
                    typeof row === 'object' ? row : { text: row };
                  if (!row.entityName) {
                    rowValue.entityName = entity.name;
                  }
                  if (!row.entity) {
                    const entityClone = Object.assign({}, entity);
                    delete entityClone.recipients;
                    delete entityClone.issues;
                    rowValue.entity = entityClone;
                  }
                  const dataLakeRow = Object.assign(
                    {
                      issueProviderName: providerName,
                      issueTimestamp: started,
                      issueTypeName: issueTypeName,
                    },
                    rowValue
                  );
                  dataLakeOutput.push(JSON.stringify(dataLakeRow));
                }
              }
            }
          }
        }
      }
    }
  }
  if (dataLakeOutput.length) {
    insights.trackEvent({ name: 'JobReportsReportDataLakeSaving' });
    return saveDataLakeOutput(context, dataLakeOutput);
  } else {
    insights.trackEvent({ name: 'JobReportsReportDataLakeEmptyReport' });
    return context;
  }
}

function saveDataLakeOutput(
  context: IReportsContext,
  dataLakeOutput: string[]
): Promise<IReportsContext> {
  // Each line of the file is its own independent JSON object
  const text = dataLakeOutput.join('\r\n');
  const insights = context.insights;
  const dla = context.settings.dataLakeAccount;
  if (!dla) {
    return Promise.reject(
      new Error('Missing Azure Data Lake / Azure Storage Account information')
    );
  }
  return new Promise((resolve, reject) => {
    const backupBlobService = azure.createBlobService(dla.account, dla.key);
    const containerName = dla.containerName;
    backupBlobService.createContainerIfNotExists(
      containerName,
      (createContainerError) => {
        if (createContainerError) {
          insights.trackException({ exception: createContainerError });
          return reject(createContainerError);
        }
        const blobPrefix = dla.blobPrefix || 'consolidatedReports';
        const backupBlobName = `${blobPrefix}_${moment
          .utc()
          .format('YYYY_MM_DD')}.json.gz`;
        writeDeflatedTextFile(text, (writeError, deflatedTempPath) => {
          if (writeError) {
            insights.trackException({ exception: writeError });
            return reject(writeError);
          }
          backupBlobService.createBlockBlobFromLocalFile(
            containerName,
            backupBlobName,
            deflatedTempPath,
            (cloudError) => {
              if (cloudError) {
                insights.trackException({ exception: cloudError });
                return reject(cloudError);
              }
              // Successful
              insights.trackEvent({
                name: 'JobReportsReportDataLakeBackup',
                properties: {
                  filename: backupBlobName,
                  containerName: containerName,
                  account: dla.account,
                },
              });
              return resolve(context);
            }
          );
        });
      }
    );
  });
}

async function storeReports(
  context: IReportsContext
): Promise<IReportsContext> {
  context.insights.trackEvent({ name: 'JobReportsReportStoringStarted' });
  const report = Object.assign({}, context.consolidated);
  const consolidatedSchemaVersion = context.settings.consolidatedSchemaVersion;
  report.metadata = {
    started: context.started,
    startedText: moment(context.started)
      .tz('America/Los_Angeles')
      .format(reportGeneratedFormat),
    finished: moment().format(),
    version: consolidatedSchemaVersion,
  };
  const json = JSON.stringify(report);
  const storeLocalReportPath = context.settings.storeLocalReportPath;
  if (storeLocalReportPath) {
    await storeLocalReport(report, storeLocalReportPath, context);
  }
  if (!context.reports.store) {
    context.insights.trackEvent({ name: 'JobReportsReportStoringSkipped' });
    return context;
  }
  const stringSizeUncompressed = fileSize(
    Buffer.byteLength(json, 'utf8')
  ).human();
  context.insights.trackEvent({
    name: 'JobReportsReportStoring',
    properties: {
      size: stringSizeUncompressed,
      version: consolidatedSchemaVersion,
    },
  });
  const ttl = context.settings.witnessEventReportsTimeToLiveMinutes;
  if (!ttl) {
    throw new Error(
      'No witnessEventReportsTimeToLiveMinutes configuration value defined for the report TTL. To make efficient use of Redis memory, a TTL must be provided.'
    );
  }
  const reportingRedis = context.reports.reportRedisClient;
  const reportingKey = context.settings.witnessEventKey;
  if (reportingRedis && reportingKey) {
    reportingRedis.setCompressedWithExpire(reportingKey, json, ttl);
  }
  return context;
}

async function storeLocalReport(
  report,
  storeLocalReportPath,
  context: IReportsContext
): Promise<IReportsContext> {
  const prettyFile = JSON.stringify(report, undefined, 2);
  writeTextToFile(storeLocalReportPath, prettyFile);
  return context;
}

async function recordMetrics(
  context: IReportsContext
): Promise<IReportsContext> {
  const insights = context.insights;
  const consolidated = context.consolidated;
  let overallIssues = 0;
  for (let i = 0; i < providerNames.length; i++) {
    const providerName = providerNames[i];
    const root = consolidated[providerName];
    if (root) {
      const metricRoot = `JobRepoReportIssues${providerName}`;
      const countByIssue = new Map();
      const definitions = {};
      for (let x = 0; x < root.definitions.length; x++) {
        const d = root.definitions[x];
        definitions[d.name] = d;
      }
      if (root.entities) {
        for (let j = 0; j < root.entities.length; j++) {
          const entity = root.entities[j];
          if (entity && entity.issues) {
            const issueList = Object.getOwnPropertyNames(entity.issues);
            for (let k = 0; k < issueList.length; k++) {
              const issues = entity.issues[issueList[k]];
              const definition = definitions[issueList[k]];
              let targetCollectionName = null;
              if (definition && definition.hasTable) {
                targetCollectionName = 'rows';
              } else if (definition && definition.hasList) {
                targetCollectionName = 'listItems';
              }
              if (
                targetCollectionName &&
                issues[targetCollectionName] &&
                Array.isArray(issues[targetCollectionName])
              ) {
                const count = issues[targetCollectionName].length;
                let currentValue = countByIssue.get(issueList[k]);
                if (!currentValue) {
                  currentValue = 0;
                }
                countByIssue.set(issueList[k], currentValue + count);
              }
            }
          }
        }
      }
      // Report metric for this provider and all of its issues (total)
      const issueNameList = Array.from(countByIssue.keys());
      for (let j = 0; j < issueNameList.length; j++) {
        const issueName = issueNameList[j];
        const count = countByIssue.get(issueName);
        const metricName = `${metricRoot}${issueName}`;
        insights.trackMetric({ name: metricName, value: count });
        overallIssues += count;
      }
    }
  }
  insights.trackMetric({
    name: 'JobRepoReportsIssuesOverall',
    properties: overallIssues,
  });
  return context;
}

async function sendReports(context: IReportsContext): Promise<IReportsContext> {
  if (!context.reports.send) {
    context.insights.trackEvent({ name: 'JobReportsSendingSkipped' });
    return context;
  }
  context.insights.trackEvent({ name: 'JobReportsReportSendingStarted' });
  await mailer(context);
  return context;
}
