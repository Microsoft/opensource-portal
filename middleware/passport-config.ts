//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import passport from 'passport';

import createAADStrategy from './passport/aadStrategy';
import createGithubStrategy from './passport/githubStrategy';
import serializer from './passport/serializer';

export default function (app, config) {
  const supportedAuth = ['github', 'aad', 'oauth2'];

  if (!supportedAuth.includes(config.authentication.scheme)) {
    throw new Error(
      `Unsupported primary authentication scheme type "${config.authentication.scheme}"`
    );
  }

  // Always set up GitHub strategies
  const githubStrategies = createGithubStrategy(app, config);
  for (const name in githubStrategies) {
    passport.use(name, githubStrategies[name]);
    // Validate and borrow a few parameters from the GitHub passport library, for the default github strategy
    if (name === 'github') {
      const strategy = githubStrategies[name];
      // @ts-ignore
      if (strategy._oauth2 && strategy._oauth2._authorizeUrl) {
        // @ts-ignore
        app.set(
          'runtime/passport/github/authorizeUrl',
          strategy._oauth2._authorizeUrl
        );
      } else {
        throw new Error(
          'The GitHub Passport strategy library may have been updated, it no longer contains the expected Authorize URL property within the OAuth2 object.'
        );
      }
      // @ts-ignore
      if (strategy._scope && strategy._scopeSeparator) {
        // @ts-ignore
        app.set(
          'runtime/passport/github/scope',
          strategy._scope.join(strategy._scopeSeparator)
        );
      } else {
        throw new Error(
          'The GitHub Passport strategy library may have been updated, it no longer contains the expected Authorize URL property within the OAuth2 object.'
        );
      }
    }
  }

  if (config.authentication.scheme == 'aad') {
    const aadStrategies = createAADStrategy(app, config);
    for (const name in aadStrategies) {
      passport.use(name, aadStrategies[name]);
    }
  } else if (config.authentication.scheme == 'oauth2') {
    // Set up oauth2 strategy here
    throw new Error('oauth2 is not currently implemented');
  }

  app.use(passport.initialize());
  app.use(passport.session());

  const serializerOptions = {
    config: config,
    keyResolver: app.get('keyEncryptionKeyResolver'),
  };

  passport.serializeUser(serializer.serialize(serializerOptions));
  passport.deserializeUser(serializer.deserialize(serializerOptions));
  serializer.initialize(serializerOptions, app);

  app.use((req, res, next) => {
    if (
      req.insights &&
      req.insights.properties &&
      config.authentication.scheme === 'aad' &&
      req.user &&
      req.user.azure
    ) {
      req.insights.properties.aadId = req.user.azure.oid;
    }
    next();
  });

  return passport;
}
