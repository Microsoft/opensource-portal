//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

/*eslint no-console: ["error", { allow: ["log", "warn"] }] */

import bodyParser from 'body-parser';
import compression from 'compression';
import path from 'path';

const debug = require('debug')('startup');

export * from './react';
export * from './links';
export * from './business';
export * from './jsonError';

import {
  hasStaticReactClientApp,
  IProviders,
  stripDistFolderName,
} from '../transitional';
import { StaticClientApp } from './staticClientApp';
import { StaticReactClientApp } from './staticClientApp2';
import { StaticSiteFavIcon, StaticSiteAssets } from './staticSiteAssets';
import ConnectSession from './session';
import passportConfig from './passport-config';
import Onboard from './onboarding';
import viewServices from '../lib/pugViewServices';

import campaign from './campaign';
import officeHyperlinks from './officeHyperlinks';
import rawBodyParser from './rawBodyParser';

import RouteScrubbedUrl from './scrubbedUrl';
import RouteLogger from './logger';
import RouteLocals from './locals';
import RoutePassport from './passport-routes';

export default function initMiddleware(
  app,
  express,
  config,
  dirname,
  initializationError
) {
  config = config || {};
  const appDirectory =
    config && config.typescript && config.typescript.appDirectory
      ? config.typescript.appDirectory
      : stripDistFolderName(dirname);
  const providers = app.get('providers') as IProviders;
  const applicationProfile = providers.applicationProfile;
  if (initializationError) {
    providers.healthCheck.healthy = false;
  }

  app.set('views', path.join(appDirectory, 'views'));
  app.set('view engine', 'pug');

  // const pugCustomLoadPlugin = {
  //   XXresolve(filename, source, loadOptions) {
  //     console.log();
  //   },
  //   read(filename, loadOptions) {
  //     console.log();
  //   }
  // };

  // const pugRenderfile = pug.renderFile;
  // pug.renderFile = function (renderPath, renderOptions, renderCallback) {
  //   if (!renderOptions.plugins) {
  //     renderOptions.plugins = [pugCustomLoadPlugin];
  //     console.log('--added plugins--');
  //   }
  //   return pugRenderfile(renderPath, renderOptions, renderCallback);
  // };

  // const pugCompileFile = pug.compileFile;
  // pug.compileFile = function (renderPath, renderOptions) {
  //   try {
  //     return pugCompileFile(renderPath, renderOptions);
  //   } catch (noFileError) {
  //     console.log();
  //     throw noFileError;
  //   }
  // };

  //app.engine('pug', pug.__express);
  app.set('view cache', process.env.NODE_ENV !== 'development'); // CONSIDER: pull from config instead
  app.disable('x-powered-by');

  app.set('viewServices', viewServices);

  providers.viewServices = viewServices;
  if (applicationProfile.webServer) {
    StaticSiteFavIcon(app);
    app.use(rawBodyParser);
    app.use(bodyParser.json());
    app.use(bodyParser.urlencoded({ extended: false }));
    app.use(compression());
    if (applicationProfile.serveStaticAssets) {
      StaticSiteAssets(app, express);
    }
    if (hasStaticReactClientApp()) {
      StaticReactClientApp(app, express);
    }
    if (applicationProfile.serveClientAssets) {
      StaticClientApp(app, express);
    }
    providers.campaign = campaign(app);
    let passport;
    if (!initializationError) {
      if (config.containers && config.containers.deployment) {
        app.enable('trust proxy');
        debug('proxy: trusting reverse proxy');
      }
      if (applicationProfile.sessions) {
        app.use(ConnectSession(app, config, providers));
        try {
          passport = passportConfig(app, config);
        } catch (passportError) {
          initializationError = passportError;
        }
      }
    }
    app.use(RouteScrubbedUrl);
    app.use(RouteLogger(config));
    app.use(RouteLocals);
    if (!initializationError) {
      if (applicationProfile.sessions) {
        RoutePassport(app, passport, config);
        if (
          config.github.organizations.onboarding &&
          config.github.organizations.onboarding.length
        ) {
          debug('Onboarding helper loaded');
          Onboard(app, config);
        }
      }
      app.use(officeHyperlinks);
    }
  }
  if (initializationError) {
    providers.healthCheck.healthy = false;
    throw initializationError;
  } else {
    providers.healthCheck.ready = true; // Ready to accept traffic
  }
}
