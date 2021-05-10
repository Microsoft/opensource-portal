//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import moment from 'moment';

import { NoCacheNoBackground, ReposAppRequest } from '../../interfaces';
import { wrapError } from '../../utils';

router.get('/', asyncHandler(async function (req: ReposAppRequest, res, next) {
  const organization = req.organization;
  const onboarding = req.query.onboarding;
  const joining = req.query.joining;

  req.individualContext.webContext.pushBreadcrumb('Multi-factor authentication check');
  const username = req.individualContext.getGitHubIdentity().username;
  try {
    const state = await organization.isMemberSingleFactor(username, NoCacheNoBackground);
    if (state === false && (req.body.validate || onboarding || joining)) {
      let url = organization.baseUrl;
      if (onboarding || joining) {
        let urlSegment = '?' + (onboarding ? 'onboarding' : 'joining') + '=' + (onboarding ? onboarding : joining);
        url = organization.baseUrl + (onboarding ? 'profile-review' : 'teams') + urlSegment;
      }
      return res.redirect(url);
    }
    const title = state === true ? 'Please enable two-factor authentication now' : 'Thanks for using modern security practices';
    req.individualContext.webContext.render({
      view: 'org/2fa',
      title,
      state: {
        twoFactorOff: !state,
        notValidated: (req.query.validate ? true : undefined),
        onboarding: onboarding,
        organization: organization,
        nowString: moment().format('MMMM Do YYYY, h:mm:ss a'),
      },
    });
  } catch (error) {
    return next(wrapError(error, `We were unable to validate your security settings with GitHub. The error GitHub returned: ${error.message || error}`));
  }
}));

export default router;
