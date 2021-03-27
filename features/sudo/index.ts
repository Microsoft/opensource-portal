//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import type { ICorporateLink, Organization } from '../../business';
import getCompanySpecificDeployment from '../../middleware/companySpecificDeployment';
import { IProviders } from '../../transitional';

export interface IOrganizationSudo {
  isSudoer(githubLogin: string, link?: ICorporateLink): Promise<boolean>;
}

export interface IPortalSudo {
  isSudoer(githubLogin: string, link?: ICorporateLink): Promise<boolean>;
}

export abstract class OrganizationSudo implements IOrganizationSudo {
  constructor(
    protected providers: IProviders,
    protected organization: Organization
  ) {}
  abstract isSudoer(
    githubLogin: string,
    link?: ICorporateLink
  ): Promise<boolean>;

  protected isSudoEnvironmentOff() {
    // Optional helper for debugging, if the provider chooses to respect the deployment environment
    const config = this.providers.config;
    if (config?.sudo?.organization?.off) {
      console.warn(
        'DEBUG WARNING: Organization sudo support is turned off in the current environment'
      );
      return true;
    }
    return false;
  }
}

export { OrganizationFeatureSecurityGroupProperty } from './securityGroup';

export * from './portal';

import { OrganizationSudoNoop } from './noop';
import { OrganizationSudoSecurityGroup } from './securityGroup';
import { OrganizationSudoGitHubTeams } from './teams';

export function createOrganizationSudoInstance(
  providers: IProviders,
  organization: Organization
): IOrganizationSudo {
  const override = getCompanySpecificDeployment();
  let instance = override?.features?.organizationSudo?.tryCreateInstance(
    providers,
    organization
  );
  if (instance) {
    return instance;
  }

  const config = providers.config;
  const defaultProviderName = config?.sudo?.organization?.defaultProviderName;

  let providerName = defaultProviderName;

  const allowUniqueProvidersByOrganization =
    config?.sudo?.organization?.allowUniqueProvidersByOrganization;
  if (allowUniqueProvidersByOrganization) {
    const name = getProviderNameForOrganization(organization);
    if (name) {
      providerName = name;
    }
  }

  instance = createProviderInstance(providerName, providers, organization);
  return instance;
}

function getProviderNameForOrganization(organization: Organization) {
  if (OrganizationSudoSecurityGroup.forOrganization(organization)) {
    return OrganizationSudoSecurityGroup.providerName;
  }
}

function createProviderInstance(
  providerName: string,
  providers: IProviders,
  organization: Organization
) {
  switch (providerName) {
    case 'noop':
      return new OrganizationSudoNoop(providers, organization);
    case 'githubteams':
      return new OrganizationSudoGitHubTeams(providers, organization);
    case OrganizationSudoSecurityGroup.providerName:
      return new OrganizationSudoSecurityGroup(providers, organization);
    default:
      throw new Error(
        `OrganizationSudo: unsupported or unconfigured provider name=${providerName}`
      );
  }
}
