//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import path from 'path';

import CosmosSessionStore from '../lib/cosmosSession';

import { RedisOptions } from '../transitional';
import { createAndInitializeLinkProviderInstance } from '../lib/linkProviders';

import { Operations } from '../business';
import { createAndInitializeEntityMetadataProviderInstance, IEntityMetadataProvidersOptions } from '../lib/entityMetadataProvider';
import { createAndInitializeRepositoryMetadataProviderInstance } from '../entities/repositoryMetadata';

import { createMailAddressProviderInstance, IMailAddressProvider } from '../lib/mailAddressProvider';

import ErrorRoutes from './error-routes';

import redis from 'redis';
import { Pool as PostgresPool } from 'pg';

const redisMock = require('redis-mock');
const debug = require('debug')('startup');
const pgDebug = require('debug')('pgpool');

import appInsights from './appInsights';
import keyVault from './keyVault';

import healthCheck from './healthCheck';

import expressRoutes from '../routes/';
import alternateRoutes from './alternateApps';

import RedisHelper from '../lib/caching/redis';
import { createTokenProvider } from '../entities/token';
import { createAndInitializeApprovalProviderInstance } from '../entities/teamJoinApproval';
import { CreateLocalExtensionKeyProvider } from '../entities/localExtensionKey';
import { CreateGraphProviderInstance, IGraphProvider } from '../lib/graphProvider/';
import initializeCorporateViews from './corporateViews';

import keyVaultResolver from '../lib/keyVaultResolver';

import { createMailProviderInstance } from '../lib/mailProvider/';
import { RestLibrary } from '../lib/github';
import { CreateRepositoryCacheProviderInstance } from '../entities/repositoryCache';
import { CreateRepositoryCollaboratorCacheProviderInstance } from '../entities/repositoryCollaboratorCache';
import { CreateTeamCacheProviderInstance } from '../entities/teamCache';
import { CreateTeamMemberCacheProviderInstance } from '../entities/teamMemberCache';
import { CreateRepositoryTeamCacheProviderInstance } from '../entities/repositoryTeamCache';
import { CreateOrganizationMemberCacheProviderInstance } from '../entities/organizationMemberCache';
import QueryCache from '../business/queryCache';
import { createAndInitializeOrganizationSettingProviderInstance } from '../entities/organizationSettings';
import { IEntityMetadataProvider } from '../lib/entityMetadataProvider/entityMetadataProvider';
import { createAndInitializeAuditLogRecordProviderInstance } from '../entities/auditLogRecord';
import CosmosCache from '../lib/caching/cosmosdb';
import BlobCache from '../lib/caching/blob';
import { StatefulCampaignProvider } from '../lib/campaigns';
import CosmosHelper from '../lib/cosmosHelper';
import createCorporateContactProviderInstance from '../lib/corporateContactProvider';
import { IQueueProcessor } from '../lib/queues';
import ServiceBusQueueProcessor from '../lib/queues/servicebus';
import AzureQueuesProcessor from '../lib/queues/azurequeue';
import { UserSettingsProvider } from '../entities/userSettings';
import getCompanySpecificDeployment from './companySpecificDeployment';

import RouteCorrelationId from './correlationId';
import RouteHsts from './hsts';
import RouteSslify from './sslify';

import MiddlewareIndex from '.';
import { ICacheHelper } from '../lib/caching';
import { IApplicationProfile, IProviders, IReposApplication, InnerError } from '../interfaces';

const DefaultApplicationProfile: IApplicationProfile = {
  applicationName: 'GitHub Management Portal',
  serveStaticAssets: true,
  serveClientAssets: true,
  logDependencies: true,
  webServer: true,
  sessions: true,
};

type CompanyStartupEntrypoint = (config: any, providers: IProviders, rootdir: string) => Promise<void>;

async function initializeAsync(app: IReposApplication, express, rootdir: string, config): Promise<void> {
  const providers = app.get('providers') as IProviders;
  providers.postgresPool = await ConnectPostgresPool(config.data.postgres);
  providers.linkProvider = await createAndInitializeLinkProviderInstance(providers, config);
  if (config.github.cache.provider === 'cosmosdb') {
    const cosmosCache = new CosmosCache(config.github.cache.cosmosdb);
    await cosmosCache.initialize();
    providers.cacheProvider = cosmosCache;
  } else if (config.github.cache.provider === 'blob') {
    const blobCache = new BlobCache(config.github.cache.blob);
    await blobCache.initialize();
    providers.cacheProvider = blobCache;
  } else if (config.github.cache.provider === 'redis') {
    const redisClient = await connectRedis(config, config.redis, 'cache');
    const redisHelper = new RedisHelper({ redisClient, prefix: config.redis.prefix });
    // providers.redisClient = redisClient;
    providers.cacheProvider = redisHelper;
  } else {
    throw new Error('No cache provider available');
  }

  providers.graphProvider = await createGraphProvider(providers, config);
  app.set('graphProvider', providers.graphProvider);

  providers.mailAddressProvider = await createMailAddressProvider(config, providers);
  app.set('mailAddressProvider', providers.mailAddressProvider);

  const mailProvider = createMailProviderInstance(config);
  const mailInitializedMessage = await mailProvider.initialize();
  debug(`mail provider type=${config.mail.provider} ${mailInitializedMessage}`);
  providers.mailProvider = mailProvider;

  providers.github = configureGitHubLibrary(providers.cacheProvider, config);
  app.set('github', providers.github);

  // always check if config exists to prevent crashing because of trying to access an undefined object
  const emOptions: IEntityMetadataProvidersOptions = {
    tableOptions: {
      account: config.github?.links?.table?.account,
      key: config.github?.links?.table?.key,
      prefix: config.github?.links?.table?.prefix,
      encryption: {
        keyEncryptionKeyId: config.github?.links?.table?.encryptionKeyId,
        keyResolver: providers.keyEncryptionKeyResolver,
      },
    },
    postgresOptions: {
      pool: providers.postgresPool,
    },
  };
  let tableProviderEnabled = emOptions.tableOptions && emOptions.tableOptions.account && emOptions.tableOptions.key;
  let postgresProviderEnabled = emOptions.postgresOptions && emOptions.postgresOptions.pool;
  const tableEntityMetadataProvider = tableProviderEnabled ? await createAndInitializeEntityMetadataProviderInstance(
    emOptions,
    'table') : null;
  const pgEntityMetadataProvider = postgresProviderEnabled ? await createAndInitializeEntityMetadataProviderInstance(
    emOptions,
    'postgres') : null;
  const memoryEntityMetadataProvider = await createAndInitializeEntityMetadataProviderInstance(
    emOptions,
    'memory');
  const defaultProvider = pgEntityMetadataProvider || tableEntityMetadataProvider || memoryEntityMetadataProvider;
  function providerNameToInstance(value: string): IEntityMetadataProvider {
    switch (value) {
      case 'firstconfigured':
        return defaultProvider;
      case 'postgres':
        return pgEntityMetadataProvider;
      case 'table':
        return tableEntityMetadataProvider;
      case 'memory':
        return memoryEntityMetadataProvider;
      default:
        return null;
    }
  }
  providers['_temp:nameToInstance'] = providerNameToInstance;
  providers.defaultEntityMetadataProvider = defaultProvider;
  providers.approvalProvider = await createAndInitializeApprovalProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.teamjoin) });
  providers.tokenProvider = await createTokenProvider({ entityMetadataProvider: providerNameToInstance(config.entityProviders.tokens) });
  providers.localExtensionKeyProvider = await CreateLocalExtensionKeyProvider({ entityMetadataProvider: providerNameToInstance(config.entityProviders.localextensionkey) });
  providers.organizationMemberCacheProvider = await CreateOrganizationMemberCacheProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.organizationmembercache) });
  providers.organizationSettingsProvider = await createAndInitializeOrganizationSettingProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.organizationsettings) });
  providers.repositoryCacheProvider = await CreateRepositoryCacheProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.repositorycache) });
  providers.repositoryCollaboratorCacheProvider = await CreateRepositoryCollaboratorCacheProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.repositorycollaboratorcache) });
  providers.repositoryTeamCacheProvider = await CreateRepositoryTeamCacheProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.repositoryteamcache) });
  providers.teamCacheProvider = await CreateTeamCacheProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.teamcache) });
  providers.teamMemberCacheProvider = await CreateTeamMemberCacheProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.teammembercache) });
  providers.auditLogRecordProvider = await createAndInitializeAuditLogRecordProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.auditlogrecord) });
  providers.userSettingsProvider = new UserSettingsProvider({ entityMetadataProvider: providerNameToInstance(config.entityProviders.usersettings) });
  await providers.userSettingsProvider.initialize();
  providers.queryCache = new QueryCache(providers);
  if (config.campaigns && config.campaigns.provider === 'cosmosdb') {
    const campaignCosmosStore = new CosmosHelper({
      endpoint: config.campaigns.cosmosdb.endpoint,
      key: config.campaigns.cosmosdb.key,
      database: config.campaigns.cosmosdb.database,
      collection: config.campaigns.cosmosdb.collection,
    });
    await campaignCosmosStore.initialize();
    providers.campaignStateProvider = new StatefulCampaignProvider(campaignCosmosStore);
  }
  if (config.session.provider === 'cosmosdb') {
    const cosmosStore = new CosmosSessionStore({
      endpoint: config.session.cosmosdb.endpoint,
      key: config.session.cosmosdb.key,
      database: config.session.cosmosdb.database,
      collection: config.session.cosmosdb.collection,
      ttl: config.session.cosmosdb.ttl,
    });
    await cosmosStore.initialize();
    providers.session = cosmosStore;
  } else if (config.session.provider === 'redis') {
    const redisSessionClient = await connectRedis(config, config.session.redis, 'session');
    providers.sessionRedisClient = redisSessionClient;
  }
  if (config?.diagnostics?.blob?.key) {
    providers.diagnosticsDrop = new BlobCache({
      key: config.diagnostics.blob.key,
      account: config.diagnostics.blob.account,
      container: config.diagnostics.blob.container,
    });
    await providers.diagnosticsDrop.initialize();
  }

  providers.corporateContactProvider = createCorporateContactProviderInstance(config, providers.cacheProvider);
  providers.corporateAdministrationProfile = getCompanySpecificDeployment()?.administrationSection;
  providers.corporateViews = await initializeCorporateViews(providers, rootdir);

  await dynamicStartup(config, providers, rootdir);

  const webhooksConfig = config.github.webhooks;
  if (webhooksConfig && webhooksConfig.provider) {
    let webhooksProvider: IQueueProcessor = null;
    if (webhooksConfig.provider === 'servicebus') {
      const serviceBusConfig = webhooksConfig.serviceBus;
      webhooksProvider = new ServiceBusQueueProcessor(serviceBusConfig);
    } else if (webhooksConfig.provider === 'azurequeues') {
      const queuesConfig = webhooksConfig.azureQueues;
      webhooksProvider = new AzureQueuesProcessor(queuesConfig);
    }
    if (webhooksProvider) {
      await webhooksProvider.initialize();
      providers.webhookQueueProcessor = webhooksProvider;
    }
  }

  try {
    const repositoryMetadataProvider = await createAndInitializeRepositoryMetadataProviderInstance({ entityMetadataProvider: providerNameToInstance(config.entityProviders.repositorymetadata) });
    const operations = new Operations({ providers, repositoryMetadataProvider, github: providers.github });
    await operations.initialize();
    app.set('operations', operations);
    providers.operations = operations;
  } catch (ignoredError2) {
    console.dir(ignoredError2);
    throw ignoredError2;
  }

  if (providers.applicationProfile.startup) {
    debug('Application provider-specific startup...');
    await providers.applicationProfile.startup(providers);
  }
}

function configureGitHubLibrary(cacheProvider: ICacheHelper, config): RestLibrary {
  if (config && config.github && config.github.operations && !config.github.operations.centralOperationsToken) {
    debug('WARNING: no central GitHub operations token is defined, some library operations may not succeed. ref: config.github.operations.centralOperationsToken var: GITHUB_CENTRAL_OPERATIONS_TOKEN');
  }
  const libraryContext = new RestLibrary({
    config,
    cacheProvider,
  });
  return libraryContext;
}

// Asynchronous initialization for the Express app, configuration and data stores.
export default async function initialize(app: IReposApplication, express, rootdir: string, config: any, exception: Error): Promise<IReposApplication> {
  if (!config) {
    throw new Error('No configuration resolved');
  }
  const applicationProfile = config?.web?.app && config.web.app !== 'repos' ? await alternateRoutes(config, app, config.web.app) : DefaultApplicationProfile;
  const web = !(config?.skipModules && config.skipModules.has('web'));
  if (applicationProfile.webServer && !web) {
    applicationProfile.webServer = false;
  }
  const containerPurpose = config && config.isJobInternal ? 'Job' : (applicationProfile.webServer ? 'Web Application' : 'Application');
  debug(`${containerPurpose} name: ${applicationProfile.applicationName}`);
  debug(`Environment: ${config?.debug?.environmentName || 'Unknown'}`);
  if (!exception && applicationProfile.validate) {
    try {
      await applicationProfile.validate();
    } catch (error) {
      exception = error;
    }
  }
  app.set('started', new Date());
  app.config = config;
  if (exception) {
    // Once app insights is available, will try to log this exception; display for now.
    console.dir(exception);
  }
  app.set('basedir', rootdir);
  const providers: IProviders = {
    app,
    basedir: rootdir,
    applicationProfile,
  };
  app.set('providers', providers);
  app.providers = providers;
  app.set('runtimeConfig', config);
  providers.healthCheck = healthCheck(app, config);
  if (applicationProfile.webServer) {
    if (!app.startServer) {
      throw new Error(`app.startServer is required for web applications`);
    }
    await app.startServer();
  }
  app.use(RouteCorrelationId);
  providers.insights = appInsights(app, config);
  app.set('appInsightsClient', providers.insights);
  if (!exception && (!config || !config.activeDirectory)) {
    exception = new Error(`config.activeDirectory.clientId and config.activeDirectory.clientSecret are required to initialize KeyVault`);
  }
  app.use('*', (req, res, next) => {
    if (providers.healthCheck.ready) {
      return next();
    }
    return res.send('Service not ready.');
  });
  if (config.containers && config.containers.deployment) {
    debug('Container deployment: HTTP: listening, HSTS: on');
    app.use(RouteHsts);
  } else if (config.containers && config.containers.docker) {
    debug('Docker image: HTTP: listening, HSTS: off');
  } else if (config.webServer.allowHttp) {
    debug('WARNING: Development mode (DEBUG_ALLOW_HTTP): HTTP: listening, HSTS: off');
  } else {
    app.use(RouteSslify);
    app.use(RouteHsts);
  }
  if (!exception) {
    const kvConfig = {
      clientId: config && config.activeDirectory ? config.activeDirectory.clientId : null,
      clientSecret: config && config.activeDirectory ? config.activeDirectory.clientSecret : null,
    };
    providers.config = config;
    let keyEncryptionKeyResolver = null;
    try {
      const keyVaultClient = keyVault(kvConfig);
      keyEncryptionKeyResolver = keyVaultResolver(keyVaultClient);
      app.set('keyEncryptionKeyResolver', keyEncryptionKeyResolver);
      providers.keyEncryptionKeyResolver = keyEncryptionKeyResolver;
      debug('configuration secrets resolved');
    } catch (noKeyVault) {
      if (!kvConfig.clientId && !kvConfig.clientSecret) {
        debug('configuration resolved, no key vault client configured');
      } else {
        console.warn(noKeyVault);
        throw noKeyVault;
      }
    }
    try {
      await initializeAsync(app, express, rootdir, config);
    } catch (initializeError) {
      console.dir(initializeError);
      debug(`Initialization failure: ${initializeError}`);
      exception = initializeError;
    }
  }
  try {
    MiddlewareIndex(app, express, config, rootdir, exception);
  } catch (middlewareError) {
    exception = middlewareError;
  }
  // ROUTES:
  if (!exception) {
    if (applicationProfile.customRoutes) {
      await applicationProfile.customRoutes();
    } else {
      app.use('/', expressRoutes);
    }
  } else {
    console.error(exception);
    const appInsightsClient = providers.insights;
    const crash = (error) => {
      return () => {
        debug('App crashed because of an initialization error.');
        console.log(error.message);
        if (error.stack) {
          console.log(error.stack);
        }
        process.exit(1);
      };
    };
    if (appInsightsClient) {
      appInsightsClient.trackException({
        exception,
        properties: {
          info: 'App crashed while initializing',
        },
      });
      try {
        appInsightsClient.flush({ isAppCrashing: true, callback: crash(exception) });
      } catch (sendError) {
        console.dir(sendError);
        crash(exception)();
      }
    } else {
      crash(exception)();
    }
  }
  await ErrorRoutes(app, exception);
  return app;
}

function createGraphProvider(providers: IProviders, config: any): Promise<IGraphProvider> {
  return new Promise((resolve, reject) => {
    // The graph provider is optional. A graph provider can connect to a
    // corporate directory to validate or lookup employees and other
    // directory members at runtime to gather additional information.
    CreateGraphProviderInstance(providers, config, (providerInitError: Error, provider: IGraphProvider) => {
      if (providerInitError) {
        debug(`No org chart graph provider configured: ${providerInitError.message}`);
        if (config.graph?.require === true) {
          return reject(new Error(`Unable to initialize the graph provider: ${providerInitError.message}`));
        }
      } else {
        return resolve(provider);
      }
      return resolve(null);
    });
  });
}

export function ConnectPostgresPool(postgresConfigSection: any): Promise<PostgresPool> {
  return new Promise((resolve, reject) => {
    try {
      if (postgresConfigSection && postgresConfigSection.user) {
        const pool = new PostgresPool(postgresConfigSection);
        // central
        pool.on('error', (err, client) => {
          pgDebug('POSTGRES POOL ERROR:');
          pgDebug(err);
        });
        pool.on('connect', (client) => {
          pgDebug(`Pool connecting a new client (pool: ${pool.totalCount} clients, ${pool.idleCount} idle, ${pool.waitingCount} waiting)`);
        });
        pool.on('acquire', client => {
          pgDebug(`Postgres client being checked out (pool: ${pool.totalCount} clients, ${pool.idleCount} idle, ${pool.waitingCount} waiting)`);
        });
        pool.on('remove', client => {
          pgDebug(`Postgres client checked back in (pool: ${pool.totalCount} clients, ${pool.idleCount} idle, ${pool.waitingCount} waiting)`);
        });
        // try connecting
        pool.connect((err, client, release) => {
          if (err) {
            const poolError: InnerError = new Error(`There was a problem connecting to the Postgres server`);
            poolError.inner = err;
            return reject(poolError);
          }
          client.query('SELECT NOW()', (err, result) => {
            release();
            if (err) {
              const poolQueryError: InnerError = new Error('There was a problem performing a test query to the Postgres server');
              poolQueryError.inner = err;
              return reject(poolQueryError);
            }
            debug(`connected to Postgres (${postgresConfigSection.host} ${postgresConfigSection.database} as ${postgresConfigSection.user}) and a pool of ${postgresConfigSection.max} clients is available in providers.postgresPool`);
            return resolve(pool);
          })
        })
      } else {
        return resolve(undefined);
      }
    } catch (failProblem) {
      return reject(failProblem);
    }
  });
}

function connectRedis(config: any, redisConfig: any, purpose: string): Promise<redis.RedisClient> {
  const nodeEnvironment = config && config.node ? config.node.environment : null;
  let redisClient: redis.RedisClient = null;
  const redisOptions: RedisOptions = {
    detect_buffers: true,
  };
  if (config.redis.key) {
    redisOptions.auth_pass = config.redis.key;
  }
  if (redisConfig.tls) {
    redisOptions.tls = {
      servername: redisConfig.tls,
    };
  }
  if (!redisConfig.host && !redisConfig.tls) {
    if (nodeEnvironment === 'production') {
      console.warn(`${purpose}: Redis host or TLS host must be provided in production environments`);
      throw new Error(`No ${purpose}.redis.host or ${purpose}.redis.tls`);
    }
    debug(`mocking Redis, in-memory provider in use`);
    redisClient = redisMock.createClient();
  } else {
    debug(`connecting to ${purpose} Redis ${redisConfig.host || redisConfig.tls}`);
    const port = redisConfig.port || (redisConfig.tls ? 6380 : 6379);
    redisClient = redis.createClient(port, redisConfig.host || redisConfig.tls, redisOptions);
  }
  let isFirst = true;
  return new Promise((resolve, reject) => {
    redisClient.on('connect', function () {
      if (isFirst) {
        isFirst = false;
        return resolve(redisClient);
      }
    });
    // NOTE: a timeout would hang the process here
    if (config.redis.key) {
      redisClient.auth(config.redis.key);
      debug(`authenticated to Redis for ${purpose}`);
    } else {
      debug(`connected to Redis for ${purpose} (unauthenticated)`);
    }
  });
}

async function createMailAddressProvider(config: any, providers: IProviders): Promise<IMailAddressProvider> {
  const options = {
    config: config,
    providers: providers,
  };
  return createMailAddressProviderInstance(options);
}

async function dynamicStartup(config: any, providers: IProviders, rootdir: string) {
  const p = config?.startup?.path;
  if (p) {
    try {
      const dynamicInclude = require(path.join(rootdir, p));
      const entrypoint = dynamicInclude && dynamicInclude.default ? dynamicInclude.default : dynamicInclude;
      if (typeof(entrypoint) !== 'function') {
        throw new Error(`Entrypoint ${p} is not a function`);
      }
      const promise = (entrypoint as CompanyStartupEntrypoint).call(null, config, providers, rootdir) as Promise<void>;
      await promise;
      debug(`Company-specific startup complete (${p})`);
    } catch (dynamicLoadError) {
      throw new Error(`config.startup.path=${p} could not successfully load: ${dynamicLoadError}`);
    }
  }
}
