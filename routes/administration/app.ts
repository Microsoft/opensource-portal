//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import {
  getProviders,
  NoCacheNoBackground,
  ReposAppRequest,
} from '../../transitional';
import GitHubApplication, {
  IGitHubAppInstallation,
} from '../../business/application';
import {
  OrganizationSetting,
  IBasicGitHubAppInstallation,
  SpecialTeam,
} from '../../entities/organizationSettings/organizationSetting';
import { IndividualContext } from '../../user';
import { Operations } from '../../business';
import {
  Organization,
  OrganizationMembershipRole,
  OrganizationMembershipState,
} from '../../business';

router.use(
  '/:appId',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const providers = getProviders(req);
    const appId = Number(req.params.appId);
    const app = providers.operations.getApplicationById(appId);
    if (app) {
      req['githubApplication'] = app;
      return next();
    }
    const notFound = new Error(`App ${req.params.appId} is not configured`);
    notFound['status'] = 404;
    return next(notFound);
  })
);

router.get(
  '/:appId',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const githubApplication = req['githubApplication'] as GitHubApplication;
    const installationIdString = req.query.installation_id;
    const setupAction = req.query.setup_action;
    if (installationIdString && setupAction) {
      return res.redirect(
        `./${githubApplication.id}/installations/${installationIdString}?setup_action=${setupAction}`
      );
    }
    const individualContext = req.individualContext;
    const allInstalls = await githubApplication.getInstallations({
      maxAgeSeconds: 5,
    });
    const { valid, invalid } = GitHubApplication.filterInstallations(
      allInstalls
    );
    individualContext.webContext.render({
      view: 'administration/setup/app',
      title: `Application ${githubApplication.friendlyName}`,
      state: {
        installations: {
          valid,
          invalid,
        },
        app: githubApplication,
      },
    });
  })
);

function getOrganizationConfiguration(config: any, orgName: string) {
  orgName = orgName.toLowerCase();
  if (config.github && config.github.organizations) {
    for (const entry of config.github.organizations) {
      if (entry && entry.name && entry.name.toLowerCase() === orgName) {
        return entry;
      }
    }
  }
}

router.use(
  '/:appId/installations/:installationId',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const githubApplication = req['githubApplication'] as GitHubApplication;
    const installationIdString = req.params.installationId;
    const { config, organizationSettingsProvider } = getProviders(req);
    const installationId = Number(installationIdString);
    const installation = await githubApplication.getInstallation(
      installationId
    );
    const invalidReasons = GitHubApplication.isInvalidInstallation(
      installation
    );
    if (invalidReasons.length) {
      throw new Error(invalidReasons.join(', '));
    }
    const organizationId = installation.account.id;
    const organizationName = installation.account.login;
    let settings: OrganizationSetting = null;
    try {
      settings = await organizationSettingsProvider.getOrganizationSetting(
        organizationId.toString()
      );
    } catch (notFound) {
      /* ignored */
    }
    const staticSettings = getOrganizationConfiguration(
      config,
      organizationName
    );

    req['installationConfiguration'] = {
      staticSettings,
      dynamicSettings: settings,
      installation,
    };
    return next();
  })
);

function isInstallationConfigured(
  settings: OrganizationSetting,
  installation: IGitHubAppInstallation
): boolean {
  if (!settings || !settings.installations) {
    return false;
  }
  for (const install of settings.installations) {
    if (install.installationId === installation.id) {
      return true;
    }
  }
  return false;
}

async function getDynamicSettingsFromLegacySettings(
  operations: Operations,
  staticSettings: any,
  installation: IGitHubAppInstallation,
  individualContext: IndividualContext
): Promise<[OrganizationSetting, Organization]> {
  const settings = OrganizationSetting.CreateFromStaticSettings(staticSettings);

  if (installation.target_type !== 'Organization') {
    throw new Error(
      `Unsupported GitHub App target of ${installation.target_type}.`
    );
  }
  settings.organizationName = installation.account.login;
  settings.organizationId = installation.account.id;

  const thisInstallation: IBasicGitHubAppInstallation = {
    appId: installation.app_id,
    installationId: installation.id,
  };
  settings.installations.push(thisInstallation);

  settings.updated = new Date();
  settings.setupDate = new Date();
  settings.setupByCorporateDisplayName =
    individualContext.corporateIdentity.displayName;
  settings.setupByCorporateId = individualContext.corporateIdentity.id;
  settings.setupByCorporateUsername =
    individualContext.corporateIdentity.username;

  settings.active = false;

  let organizationDetails = null;
  let unconfiguredOrganization: Organization = null;
  try {
    unconfiguredOrganization = operations.getUnconfiguredOrganization(settings);
    organizationDetails = await unconfiguredOrganization.getDetails();
  } catch (ignoreOrganizationDetailsProblem) {
    throw new Error(
      `Is the app still installed correctly? The app needs to be able to read the organization plan information. ${ignoreOrganizationDetailsProblem.message}`
    );
  }
  if (organizationDetails && organizationDetails.plan) {
    settings.properties['plan'] = organizationDetails.plan.name;
    if (!settings.properties['type']) {
      settings.properties['type'] =
        organizationDetails.plan.name === 'free' ? 'public' : 'publicprivate'; // free SKU or not
    }
  }

  return [settings, unconfiguredOrganization];
}

router.post(
  '/:appId/installations/:installationId',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const hasImportButtonClicked = req.body['adopt-import-settings'];
    const hasCreateButtonClicked = req.body['adopt-new-org'];
    const hasElevationButtonClicked = req.body['elevate-to-owner'];
    const forceDeleteConfig = req.body['force-delete-config'];
    const updateConfig = req.body['update'];
    const activate = req.body['activate'];
    const deactivate = req.body['deactivate'];
    const removeConfiguration = req.body['remove-configuration'];
    const addConfiguration = req.body['configure'];
    let isCreatingNew = hasImportButtonClicked || hasCreateButtonClicked;
    if (
      !hasImportButtonClicked &&
      !hasElevationButtonClicked &&
      !forceDeleteConfig &&
      !hasCreateButtonClicked &&
      !activate &&
      !deactivate &&
      !removeConfiguration &&
      !addConfiguration &&
      !updateConfig
    ) {
      return next(new Error('No supported POST parameters present'));
    }
    const providers = getProviders(req);
    const githubApplication = req['githubApplication'] as GitHubApplication;
    const individualContext = req.individualContext;
    const login = individualContext.getGitHubIdentity().username;
    const { staticSettings, dynamicSettings, installation: install } = req[
      'installationConfiguration'
    ];
    const installation = install as IGitHubAppInstallation;
    if (dynamicSettings && isCreatingNew) {
      throw new Error(
        'Settings already exist. The organization has already been adopted.'
      );
    }
    const ds = dynamicSettings as OrganizationSetting;
    let displayDynamicSettings = dynamicSettings;
    const organizationName = installation.account.login;
    const organizationSettingsProvider = providers.organizationSettingsProvider;

    if (hasElevationButtonClicked) {
      // Only available for pre-adoption
      const [
        ,
        unconfiguredOrganization,
      ] = await getDynamicSettingsFromLegacySettings(
        providers.operations,
        staticSettings,
        installation,
        individualContext
      );
      try {
        const result = await unconfiguredOrganization.addMembership(login, {
          role: OrganizationMembershipRole.Admin,
        });
        if (result?.state === OrganizationMembershipState.Pending) {
          return res.send(
            `You need to accept the membership now at: https://github.com/${unconfiguredOrganization.name}`
          );
        } else {
          return res.send('OK. Elevation should be allset.');
        }
      } catch (error) {
        return next(error);
      }
    } else if (isCreatingNew) {
      // create a new settings entity object!
      const [settings] = await getDynamicSettingsFromLegacySettings(
        providers.operations,
        staticSettings,
        installation,
        individualContext
      );
      await organizationSettingsProvider.createOrganizationSetting(settings);
      displayDynamicSettings = settings;
    } else if (dynamicSettings) {
      let goUpdate = false;
      if (activate) {
        ds.active = true;
        goUpdate = true;
      } else if (deactivate) {
        ds.active = false;
        goUpdate = true;
      }
      if (addConfiguration) {
        if (isInstallationConfigured(ds, installation)) {
          return next(
            new Error(
              `installation ${installation.id} is already dynamically configured for app ${ds.organizationName}`
            )
          );
        }
        ds.installations.push({
          appId: installation.app_id,
          installationId: installation.id,
        });
        goUpdate = true;
      } else if (removeConfiguration) {
        if (!isInstallationConfigured(ds, installation)) {
          return next(
            new Error(
              `installation ${installation.id} is not dynamically configured for app ${ds.organizationName}`
            )
          );
        }
        ds.installations = ds.installations.filter(
          (install) => install.installationId !== installation.id
        );
        goUpdate = true;
      } else if (updateConfig) {
        const newFeatureFlag = req.body['add-feature-flag'] as string;
        const changeProperty = req.body['change-property'] as string;
        if (newFeatureFlag) {
          if (ds.features.includes(newFeatureFlag)) {
            throw new Error(
              `The feature flag ${newFeatureFlag} already is added to this organization`
            );
          }
          const isBang = newFeatureFlag.startsWith('!');
          if (isBang) {
            const actualKey = newFeatureFlag.substr(1);
            ds.features = ds.features.filter((val) => val !== actualKey);
          } else {
            ds.features.push(newFeatureFlag);
          }
          goUpdate = true;
        } else if (changeProperty) {
          const i = changeProperty.indexOf(':');
          if (i < 0) {
            throw new Error('Must fit format key:value');
          }
          const removeTeamMoniker = '!team:';
          const addTeamMoniker = 'team:';
          if (
            changeProperty.startsWith(removeTeamMoniker) ||
            changeProperty.startsWith(addTeamMoniker)
          ) {
            const isTeamAdd = changeProperty.startsWith(addTeamMoniker);
            let changeTeamPropertyValue = isTeamAdd
              ? '+' + changeProperty
              : changeProperty;
            changeTeamPropertyValue = changeTeamPropertyValue.substr(
              removeTeamMoniker.length
            );
            const colonIndex = changeTeamPropertyValue.indexOf(':');
            const teamType = changeTeamPropertyValue.substr(0, colonIndex);
            const teamId = changeTeamPropertyValue.substr(colonIndex + 1);
            if (
              !['systemAdmin', 'systemWrite', 'systemRead', 'sudo'].includes(
                teamType
              )
            ) {
              // explicitly now allowing globalSudo to be set here
              throw new Error(`Unsupported team type: ${teamType}`);
            }
            let specialTeamType: SpecialTeam = null;
            switch (teamType) {
              case 'systemAdmin':
                specialTeamType = SpecialTeam.SystemAdmin;
                break;
              case 'systemWrite':
                specialTeamType = SpecialTeam.SystemWrite;
                break;
              case 'systemRead':
                specialTeamType = SpecialTeam.SystemRead;
                break;
              case 'sudo':
                specialTeamType = SpecialTeam.Sudo;
                break;
              // case 'globalSudo':
              // specialTeamType = SpecialTeam.GlobalSudo;
              // explicitly now allowing globalSudo to be set here
              // break;
              default:
                throw new Error('Unsupported team type');
            }
            ds.specialTeams = ds.specialTeams.filter(
              (notThisTeam) => notThisTeam.teamId !== Number(teamId)
            );
            if (isTeamAdd) {
              ds.specialTeams.push({
                teamId: Number(teamId),
                specialTeam: specialTeamType,
              });
            }
          }
          const key = changeProperty.substr(0, i).trim();
          const value = changeProperty.substr(i + 1).trim();
          if (value.length) {
            ds.properties[key] = value;
          } else {
            delete ds.properties[key];
          }
          goUpdate = true;
        }
      }
      if (goUpdate) {
        dynamicSettings.updated = new Date();
        await organizationSettingsProvider.updateOrganizationSetting(
          dynamicSettings
        );
      } else if (forceDeleteConfig) {
        await organizationSettingsProvider.deleteOrganizationSetting(ds);
      }
    }

    individualContext.webContext.render({
      view: 'administration/setup/installation',
      title: `${organizationName} installation of application ${githubApplication.id}`,
      state: {
        organizationName,
        dynamicSettings: displayDynamicSettings,
        staticSettings,
        installation,
        app: githubApplication,
        installationConfigured: isInstallationConfigured(
          dynamicSettings,
          installation
        ),
      },
    });
  })
);

router.get(
  '/:appId/installations/:installationId',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const githubApplication = req['githubApplication'] as GitHubApplication;
    const providers = getProviders(req);
    const individualContext = req.individualContext;
    const { staticSettings, dynamicSettings, installation } = req[
      'installationConfiguration'
    ];
    const organizationName = installation.account.login;
    const [proposedDynamicSettings, unconfiguredOrganization] = !dynamicSettings
      ? await getDynamicSettingsFromLegacySettings(
          providers.operations,
          staticSettings,
          installation,
          individualContext
        )
      : [null, null];
    const installationConfigured = isInstallationConfigured(
      dynamicSettings,
      installation
    );
    let isUserOwner = null;
    let userCheckError = null;
    if (!installationConfigured && unconfiguredOrganization) {
      try {
        const login = individualContext.getGitHubIdentity().username;
        const userMembership = await unconfiguredOrganization.getMembership(
          login,
          NoCacheNoBackground
        );
        if (userMembership?.role === OrganizationMembershipRole.Admin) {
          isUserOwner = true;
        } else if (userMembership?.role == OrganizationMembershipRole.Member) {
          isUserOwner = false;
        }
      } catch (error) {
        console.warn(error);
        userCheckError = error;
      }
    }
    individualContext.webContext.render({
      view: 'administration/setup/installation',
      title: `${organizationName} installation of application ${githubApplication.id}`,
      state: {
        organizationName,
        dynamicSettings,
        staticSettings,
        isUserOwner,
        userCheckError,
        installation,
        proposedDynamicSettings,
        app: githubApplication,
        installationConfigured,
      },
    });
  })
);

export default router;
