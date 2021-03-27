//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { OrganizationSudo } from '.';
import { ICorporateLink, Organization } from '../../business';
import { ErrorHelper, IProviders } from '../../transitional';

export const OrganizationFeatureSecurityGroupProperty = 'orgsudosecuritygroup';

export class OrganizationSudoSecurityGroup extends OrganizationSudo {
  static providerName = 'securitygroup';

  static forOrganization(organization: Organization) {
    if (!organization.hasDynamicSettings) {
      return null;
    }
    const settings = organization.getDynamicSettings();

    // Security group flips on security groups
    const val = settings.getProperty(
      OrganizationFeatureSecurityGroupProperty
    ) as string;
    if (val) {
      return OrganizationSudoSecurityGroup.providerName;
    }
  }

  constructor(providers: IProviders, organization: Organization) {
    super(providers, organization);
  }

  async isSudoer(githubLogin: string, link?: ICorporateLink): Promise<boolean> {
    const { insights } = this.providers;
    if (this.isSudoEnvironmentOff()) {
      return false;
    }
    if (!link || !link.corporateId) {
      return false;
    }
    const corporateId = link.corporateId;
    const organization = this.organization;
    const settings = organization.getDynamicSettings();
    if (!settings) {
      return false;
    }
    const securityGroupId = settings.getProperty(
      OrganizationFeatureSecurityGroupProperty
    ) as string;
    if (!securityGroupId) {
      return false;
    }
    const { graphProvider } = this.providers;
    if (!graphProvider) {
      throw new Error('No graph provider configured');
    }
    try {
      if (await graphProvider.isUserInGroup(corporateId, securityGroupId)) {
        insights?.trackEvent({
          name: 'OrganizationSudoAuthorized',
          properties: {
            corporateId,
            securityGroupId,
            organizationName: organization.name,
          },
        });
        return true;
      }
    } catch (error) {
      if (ErrorHelper.IsNotFound(error)) {
        // security groups do get deleted and should not bring down any system in that case
        return false;
      }
      console.warn(error);
      insights?.trackException({
        exception: error,
        properties: {
          eventName: 'OrganizationSudoSecurityGroupError',
          className: 'OrganizationSudoSecurityGroup',
          callName: 'isUserInGroup',
          organizationName: organization.name,
        },
      });
      return false;
    }
    return false;
  }
}
