//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { OrganizationSudo } from '.';
import { ICorporateLink, Organization } from '../../business';
import { IProviders } from '../../transitional';

export class OrganizationSudoNoop extends OrganizationSudo {
  constructor(
    protected providers: IProviders,
    protected organization: Organization
  ) {
    super(providers, organization);
  }

  isSudoer(githubLogin: string, link?: ICorporateLink): Promise<boolean> {
    return Promise.resolve(false);
  }
}
