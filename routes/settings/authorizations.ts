//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import { ReposAppRequest, getProviders } from '../../transitional';
import { ICorporateLink } from '../../business/corporateLink';
import { Operations } from '../../business';

interface IRequestWithAuthorizations extends ReposAppRequest {
  authorizations?: any;
}

function createValidator(
  operations: Operations,
  link: ICorporateLink,
  token: string
) {
  return async function (): Promise<any> {
    let data = null;
    let valid = true;
    let headers = null;
    let critical = false;
    let message = null;
    try {
      data = await operations.getAuthenticatedAccount(token);
      // NOTE: We use strings while GitHub does not
      if (data.id != /* loose */ link.thirdPartyId) {
        critical = true;
        valid = false;
        message = `This token is for a different user, "${data.login}", instead of "${link.thirdPartyUsername}".`;
      } else if (data.login != link.thirdPartyUsername) {
        message = `Your username may have changed. It once was "${link.thirdPartyUsername}" but is now "${data.login}". Your ID remains the same.`;
      }
      headers = data && data.extraFields ? data.extraFields.headers : null;
    } catch (infoError) {
      valid = false;
      if (
        infoError.statusCode === 401 &&
        infoError.message === 'Bad credentials'
      ) {
        message = 'GitHub token revoked or expired';
        critical = true;
      } else {
        message = infoError.message;
      }
    }
    return {
      valid,
      message,
      critical,
      rateLimitRemaining:
        headers && headers['x-ratelimit-remaining']
          ? headers['x-ratelimit-remaining'] + ' remaining API tokens'
          : undefined,
    };
  };
}

router.use((req: IRequestWithAuthorizations, res, next) => {
  // This is a lightweight, temporary implementation of authorization management to help clear
  // stored session tokens for apps like GitHub, VSTS, etc.
  const { operations } = getProviders(req);
  const link = req.individualContext.link;
  const authorizations = [];
  if (req.individualContext.webContext.tokens.gitHubReadToken) {
    authorizations.push({
      validator: createValidator(
        operations,
        link,
        req.individualContext.webContext.tokens.gitHubReadToken
      ),
      property: 'githubToken',
      title: 'GitHub Application: Public App Token',
      text:
        'A GitHub token, authorizing this site, is stored. This token only has rights to read your public profile and validate that you are the authorized user of the GitHub account.',
      mitigations: [
        {
          title: 'Review your GitHub authorized applications',
          url: 'https://github.com/settings/applications',
          mitigation: 'Review your authorized GitHub applications',
        },
      ],
    });
  }
  if (req.individualContext.webContext.tokens.gitHubWriteOrganizationToken) {
    authorizations.push({
      validator: createValidator(
        operations,
        link,
        req.individualContext.webContext.tokens.gitHubWriteOrganizationToken
      ),
      property: 'githubTokenIncreasedScope',
      title: 'GitHub Application: Organization Read/Write Token',
      text:
        'A GitHub token, authorizing this site, is stored. The token has a scope to read and write your organization membership. This token is used to automate organization invitation and joining functionality without requiring manual steps.',
      mitigations: [],
    });
  }
  req.authorizations = authorizations;
  next();
});

router.get('/', (req: IRequestWithAuthorizations, res) => {
  const ghi = req.individualContext.getGitHubIdentity();
  let sghi = req.individualContext.getSessionBasedGitHubIdentity();
  if (sghi && ghi && sghi.id === ghi.id) {
    sghi = null;
  }
  req.individualContext.webContext.render({
    view: 'settings/authorizations',
    title: 'Account authorizations',
    state: {
      authorizations: req.authorizations,
      gitHubAccounts: {
        primary: ghi,
        session: sghi,
      },
    },
  });
});

router.get(
  '/validate',
  asyncHandler(async (req: IRequestWithAuthorizations, res, next) => {
    const authorizations = req.authorizations;
    for (const authorization of authorizations) {
      const validator = authorization.validator;
      const validationResult = await validator();
      authorization.valid = validationResult;
      if (validationResult.critical === true) {
        // TODO: Actually delete this token/authorization
      }
    }
    req.individualContext.webContext.render({
      view: 'settings/authorizations',
      title: 'Account authorizations',
      state: {
        authorizations: authorizations,
      },
    });
  })
);

export default router;
