//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
import asyncHandler from 'express-async-handler';
const router: Router = Router();

import { getProviders } from '../../transitional';
import { sortByCaseInsensitive } from '../../utils';
import GitHubApplication from '../../business/application';
import { OrganizationSetting } from '../../entities/organizationSettings/organizationSetting';
import { ReposAppRequest, UserAlertType } from '../../interfaces';

interface IByOrgViewAppInstallation {
  app: GitHubApplication;
  installationId?: number;
}

enum OrgStatus {
  Active = 'Active',
  Adopted = 'Adopted',
  NotAdopted = 'NotAdopted',
}

interface IByOrgView {
  organizationName: string;
  status: OrgStatus;
  appInstallations: Map<number, IByOrgViewAppInstallation>;
  dynamicSettings: OrganizationSetting,
  configuredInstallations: number[],
  id?: number;
}

router.post('/', asyncHandler(async function (req: ReposAppRequest, res, next) {
  const providers = getProviders(req);
  const { deletesettingsorgname } = req.body;
  if (!deletesettingsorgname) {
    return next(new Error('Not able to complete this operation'));
  }
  const organizationSettingsProvider = providers.organizationSettingsProvider;
  const allOrgs = await organizationSettingsProvider.queryAllOrganizations();
  const thisOne = allOrgs.filter(org => org.organizationName.toLowerCase() === deletesettingsorgname.toLowerCase());
  if (thisOne.length === 1) {
    const org = thisOne[0];
    await organizationSettingsProvider.deleteOrganizationSetting(org);
    req.individualContext.webContext.saveUserAlert('Removed the org.', org.organizationName, UserAlertType.Danger);
    res.redirect('/administration/apps');
    // after the redirect, delete any caching for the org...
    if (providers.queryCache) {
      const { queryCache } = providers;
      queryCache.removeOrganizationById(String(org.organizationId));
    }
    return;
  } else {
    return next(new Error('Org not found with settings'));
  }
}));

router.get('/', asyncHandler(async function (req: ReposAppRequest, res, next) {
  const providers = getProviders(req);
  const operations = providers.operations;
  const apps = providers.operations.getApplications();
  const individualContext = req.individualContext;
  const organizationSettingsProvider = providers.organizationSettingsProvider;
  const byOrg = new Map<string, IByOrgView>();
  function getOrg(name: string) {
    let o = byOrg.get(name);
    if (!o) {
      o = {
        organizationName: name,
        id: undefined,
        status: OrgStatus.NotAdopted,
        appInstallations: new Map(),
        dynamicSettings: null,
        configuredInstallations: [],
      };
      for (const app of apps) {
        o.appInstallations.set(app.id, null);
      }
      byOrg.set(name, o);
    }
    return o;
  }
  for (const app of apps) {
    const appInstalls = await app.getInstallations({ maxAgeSeconds: 5 });
    const { valid } = GitHubApplication.filterInstallations(appInstalls);
    for (const vi of valid) {
      const organizationName = vi.account.login;
      const o = getOrg(organizationName.toLowerCase());
      o.appInstallations.set(app.id, {
        app,
        installationId: vi.id,
      });
      o.id = Number(vi.target_id);
      if (!o.dynamicSettings && vi.target_type === 'Organization') {
        try {
          o.dynamicSettings = await organizationSettingsProvider.getOrganizationSetting(vi.target_id.toString());
        } catch (ignore) { /* ignored */ }
        if (o.dynamicSettings) {
          o.configuredInstallations = o.dynamicSettings.installations.map(install => install.installationId);
          o.status = OrgStatus.Adopted;
        }
        if (o.dynamicSettings && o.dynamicSettings.active === true) {
          o.status = OrgStatus.Active;
        }
      }
    }
  }
  for (const organization of operations.organizations.values()) {
    const anOrg = getOrg(organization.name.toLowerCase());
    anOrg.id = organization.id;
  }
  const orgs = byOrg;
  const orgNames = Array.from(byOrg.keys()).sort(sortByCaseInsensitive);
  individualContext.webContext.render({
    view: 'administration/setup/apps',
    title: `GitHub Applications`,
    state: {
      apps,
      orgNames,
      orgs,
    },
  });
}));

export default router;
