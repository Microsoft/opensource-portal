//
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

'use strict';

const async = require('async');

const common = require('./common');
const wrapError = require('../utils').wrapError;

const githubEntityClassification = require('../data/github-entity-classification.json');
const primaryAccountProperties = githubEntityClassification.account.keep;
const secondaryAccountProperties = githubEntityClassification.account.strip;

class Account {
  constructor(entity, operations, getCentralOperationsToken) {
    common.assignKnownFields(this, entity, 'account', primaryAccountProperties, secondaryAccountProperties);

    const privates = _private(this);
    privates.operations = operations;
    privates.getCentralOperationsToken = getCentralOperationsToken;
  }

  // These were previously implemented in lib/user.js; functions may be needed
  // May also need to be in the org- and team- specific accounts, or built as proper objects

  contactName() {
    if (this.link) {
      return this.link.aadname || this.login;
    }
    return this.login;
  }

  contactEmail() {
    return this.link ? this.link.aadupn : null;
  }

  corporateAlias() {
    if (this.link && this.link.aadupn) {
      var email = this.link.aadupn;
      var i = email.indexOf('@');
      if (i >= 0) {
        return email.substring(0, i);
      }
    }
  }

  corporateProfileUrl() {
    const operations = _private(this).operations;
    const config = operations.config;
    const alias = this.corporateAlias();
    const corporateSettings = config.corporate;
    if (alias && corporateSettings && corporateSettings.profile && corporateSettings.profile.prefix) {
      return corporateSettings.profile.prefix + alias;
    }
  }

  avatar(optionalSize) {
    if (this.avatar_url) {
      return this.avatar_url + '&s=' + (optionalSize || 80);
    }
  }

  getProfileCreatedDate() {
    return this.created_at ? new Date(this.created_at) : undefined;
  }

  getProfileUpdatedDate() {
    return this.updated_at ? new Date(this.updated_at) : undefined;
  }

  // End previous functions

  getDetails(options, callback) {
    if (!callback && typeof(options) === 'function') {
      callback = options;
      options = null;
    }
    options = options || {};
    const self = this;
    const token = _private(this).getCentralOperationsToken();
    const operations = _private(this).operations;
    const id = this.id;
    if (!id) {
      return callback(new Error('Must provide a GitHub user ID to retrieve account information.'));
    }
    const parameters = {
      id: id,
    };
    const cacheOptions = {
      maxAgeSeconds: options.maxAgeSeconds || operations.defaults.accountDetailStaleSeconds,
    };
    if (options.backgroundRefresh !== undefined) {
      cacheOptions.backgroundRefresh = options.backgroundRefresh;
    }
    return operations.github.call(token, 'users.getById', parameters, cacheOptions, (error, entity) => {
      if (error) {
        return callback(wrapError(error, `Could not get details about account "${id}".`));
      }
      common.assignKnownFields(self, entity, 'account', primaryAccountProperties, secondaryAccountProperties);
      callback(null, entity);
    });
  }

  removeLink(callback) {
    const operations = _private(this).operations;
    const dataClient = operations.dataClient;
    const id = this.id;
    if (!id) {
      return callback(new Error('No user id known'));
    }
    dataClient.removeLink(id, error => {
      if (error) {
        return callback(error);
      }
      // CONSIDER: if there is a link service and a local
      // link cache, invalidate the local link for the user
      callback();
    });
  }

  terminate(callback) {
    const self = this;
    const insights = _private(self).operations.insights;
    if (insights) {
      insights.trackEvent('UserUnlinkStart', {
        id: self.id,
        login: self.login,
      });
    }
    self.removeManagedOrganizationMemberships(error => {
      // If a removal error occurs, do not remove the link and throw an error,
      // so that the link data and information is still present until cleaned
      if (error && insights) {
        insights.trackException(error);
      }
      if (error) {
        return callback(error);
      }
      self.removeLink(removeError => {
        if (removeError && insights) {
          insights.trackException(removeError);
        }
        if (insights) {
          insights.trackEvent('UserUnlink', {
            id: self.id,
            login: self.login,
          });
        }
        return callback(removeError);
      });
    });
  }

  // TODO: implement getOrganizationMemberships, with caching; reuse below code

  getOperationalOrganizationMemberships(callback) {
    const self = this;
    const operations = _private(self).operations;
    const organizations = operations.organizations;
    // we want to make sure that we have an ID and username
    self.getDetails(getDetailsError => {
      if (getDetailsError) {
        return callback(getDetailsError);
      }
      const username = self.login;
      if (!username) {
        return callback(new Error(`No GitHub username available for user ID ${self.id}`));
      }
      let currentOrganizationMemberships = [];
      async.eachLimit(organizations, 2, (organization, next) => {
        organization.getOperationalMembership(username, (getMembershipError, result) => {
          // getMembershipError is ignored - if there is no membership, that's fine
          if (result && result.state && (result.state === 'active' || result.state === 'pending')) {
            currentOrganizationMemberships.push(organization);
          }
          return next(/* we do not pass the error */);
        });
      }, error => {
        return callback(error ? error : null, error ? null : currentOrganizationMemberships);
      });
    });
  }

  removeManagedOrganizationMemberships(callback) {
    const self = this;
    self.getOperationalOrganizationMemberships((error, organizations) => {
      const username = self.login;
      if (error) {
        return callback(error);
      }
      async.eachLimit(organizations, 1, (organization, next) => {
        // CONSIDER: Should attempts be made to remove all no matter what, even
        // if an error happens along the way?
        organization.removeMember(username, next);
      }, callback);
    });
  }
}

module.exports = Account;

const privateSymbol = Symbol();
function _private(self) {
  if (self[privateSymbol] === undefined) {
    self[privateSymbol] = {};
  }
  return self[privateSymbol];
}
