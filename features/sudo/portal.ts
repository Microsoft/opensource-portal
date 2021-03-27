//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { IPortalSudo } from '.';
import { ICorporateLink, Organization } from '../../business';
import getCompanySpecificDeployment from '../../middleware/companySpecificDeployment';
import { ErrorHelper, IProviders } from '../../transitional';

abstract class PortalSudoBase {
  constructor(private providers: IProviders) {}
  protected isOff() {
    const config = this.providers.config;
    if (config?.sudo?.portal?.off) {
      console.warn(
        'DEBUG WARNING: Portal sudo support is turned off in the current environment'
      );
      return true;
    }
    return false;
  }

  protected forceAlways() {
    const config = this.providers.config;
    if (config?.sudo?.portal?.force) {
      console.warn(
        'DEBUG WARNING: Portal sudo is turned on for all users in the current environment'
      );
      return true;
    }
    return false;
  }
}

class PortalSudoPrimaryOrganization
  extends PortalSudoBase
  implements IPortalSudo {
  private _org: Organization;
  private _providers: IProviders;

  constructor(providers: IProviders) {
    super(providers);
    this._providers = providers;
  }

  isSudoer(githubLogin: string, link?: ICorporateLink): Promise<boolean> {
    if (this.isOff()) {
      return Promise.resolve(false);
    }
    if (this.forceAlways()) {
      return Promise.resolve(true);
    }
    if (this._org === undefined) {
      const operations = this._providers.operations;
      const primaryOrganizationName = operations.getPrimaryOrganizationName();
      this._org = primaryOrganizationName
        ? operations.getOrganization(primaryOrganizationName)
        : ((false as any) as Organization);
    }
    return this._org
      ? this._org.isSudoer(githubLogin, link)
      : Promise.resolve(false);
  }
}

class PortalSudoSecurityGroup extends PortalSudoBase implements IPortalSudo {
  private _providers: IProviders;
  private _groupId: string;

  constructor(providers: IProviders) {
    super(providers);
    if (!providers.graphProvider) {
      throw new Error('No graph provider instance available');
    }
    this._providers = providers;
    const securityGroupId = providers.config.sudo?.portal?.securityGroup?.id;
    if (!securityGroupId) {
      throw new Error('No configured security group ID');
    }
    this._groupId = securityGroupId;
  }

  async isSudoer(githubLogin: string, link?: ICorporateLink): Promise<boolean> {
    if (this.isOff()) {
      return false;
    }
    if (this.forceAlways()) {
      return true;
    }
    if (!link || !link.corporateId) {
      return false;
    }
    const insights = this._providers.insights;
    try {
      if (
        await this._providers.graphProvider.isUserInGroup(
          link.corporateId,
          this._groupId
        )
      ) {
        insights?.trackEvent({
          name: 'PortalSudoAuthorized',
          properties: {
            corporateId: link.corporateId,
            securityGroupId: this._groupId,
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
          eventName: 'PortalSudoSecurityGroupError',
          className: 'PortalSudoSecurityGroup',
          callName: 'isUserInGroup',
          corporateId: link.corporateId,
          securityGroupId: this._groupId,
        },
      });
      return false;
    }
  }
}

export function createPortalSudoInstance(providers: IProviders): IPortalSudo {
  const override = getCompanySpecificDeployment();
  let instance = override?.features?.portalSudo?.tryCreateInstance(providers);
  if (instance) {
    return instance;
  }

  const config = providers.config;
  const providerName = config?.sudo?.portal?.providerName;
  instance = createProviderInstance(providerName, providers);
  return instance;
}

function createProviderInstance(
  providerName: string,
  providers: IProviders
): IPortalSudo {
  switch (providerName) {
    case null:
    case '':
    case 'none': {
      return {
        isSudoer: () => {
          return Promise.resolve(false);
        },
      };
    }
    case 'primaryorg': {
      return new PortalSudoPrimaryOrganization(providers);
    }
    case 'securityGroup':
    case 'securitygroup': {
      return new PortalSudoSecurityGroup(providers);
    }
    default:
      throw new Error(
        `PortalSudo: unsupported or unconfigured provider name=${providerName}`
      );
  }
}
