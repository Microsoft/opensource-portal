//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import {
  ReposAppRequest,
  UserAlertType,
  ErrorHelper,
  getProviders,
} from '../../transitional';
import { UserSettings } from '../../entities/userSettings';

export interface IRequestWithUserSettings extends ReposAppRequest {
  userSettings?: UserSettings;
}

async function getSettings(req: IRequestWithUserSettings, res, next) {
  const corporateId = req.individualContext.corporateIdentity.id;
  const { userSettingsProvider } = getProviders(req);
  if (!req.userSettings) {
    let settings: UserSettings = null;
    try {
      settings = await userSettingsProvider.getUserSettings(corporateId);
    } catch (notFoundError) {
      if (ErrorHelper.IsNotFound(notFoundError)) {
        // ignore
      } else {
        throw notFoundError;
      }
    }
    if (!settings) {
      settings = new UserSettings();
      settings.corporateId = corporateId;
      await userSettingsProvider.insertUserSettings(settings);
    }
    req.userSettings = settings;
  }
  return next();
}

function view(req: IRequestWithUserSettings, res) {
  const userSettings = req.userSettings;
  req.individualContext.webContext.render({
    view: 'settings/contributionData',
    title: 'Contribution data sharing',
    state: {
      userSettings,
    },
  });
}

router.use(asyncHandler(getSettings));

router.get('/', view);

router.post(
  '/',
  asyncHandler(async function (req: IRequestWithUserSettings, res, next) {
    const isOptIn = !!(req.body.optIn === '1');
    const currentSetting = req.userSettings.contributionShareOptIn;
    req.userSettings.contributionShareOptIn = isOptIn;
    const changed = currentSetting !== isOptIn;
    if (!changed) {
      return next(new Error('No change to sharing setting.'));
    }
    const { userSettingsProvider } = getProviders(req);
    await userSettingsProvider.updateUserSettings(req.userSettings);
    const message = isOptIn
      ? 'You have opted in to sharing of contribution data.'
      : 'You have opted out of sharing contribution data.';
    const title = isOptIn ? 'Opt-in saved' : 'Opt-out';
    req.individualContext.webContext.saveUserAlert(
      message,
      title,
      UserAlertType.Success
    );
    return view(req, res);
  })
);

export default router;
