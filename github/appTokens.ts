//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { request } from '@octokit/request';
import { createAppAuth } from '@octokit/auth-app';
import {
  AuthOptions,
  Authentication,
} from '@octokit/auth-app/dist-types/types';
import { AuthInterface } from '@octokit/types';

import { AppPurpose } from '.';
import { IAuthorizationHeaderValue } from '../transitional';

interface IInstallationToken {
  installationId: number;
  organizationName: string;
  requested: Date;

  expires: Date;
  headerValue: string;
}

type OctokitAuthFunction = AuthInterface<[AuthOptions], Authentication>;

const InstallationTokenLifetimeMilliseconds = 1000 * 60 * 60;
const ValidityOffsetAfterNowMilliseconds = 1000 * 120; // how long to require validity in the future

export class GitHubAppTokens {
  #privateKey: string;
  private _appId: number;
  public purpose: AppPurpose;
  private _appAuth: OctokitAuthFunction;
  private _installationAuth = new Map<number, OctokitAuthFunction>();
  private _tokensByInstallation = new Map<number, IInstallationToken[]>();
  private _baseUrl: string;

  static CreateFromBase64EncodedFileString(
    purpose: AppPurpose,
    friendlyName: string,
    applicationId: number,
    fileContents: string,
    baseUrl?: string
  ): GitHubAppTokens {
    let keyContents = Buffer.from(fileContents, 'base64')
      .toString('utf8')
      .replace(/\r\n/g, '\n');
    return new GitHubAppTokens(
      purpose,
      friendlyName,
      applicationId,
      keyContents,
      baseUrl
    );
  }

  static CreateFromString(
    purpose: AppPurpose,
    friendlyName: string,
    applicationId: number,
    value: string,
    baseUrl?: string
  ): GitHubAppTokens {
    return new GitHubAppTokens(
      purpose,
      friendlyName,
      applicationId,
      value,
      baseUrl
    );
  }

  constructor(
    purpose: AppPurpose,
    public friendlyName: string,
    appId: number,
    privateKey: string,
    baseUrl?: string
  ) {
    this.#privateKey = privateKey;
    this._appId = appId;
    this._baseUrl = baseUrl;
    this._appAuth = createAppAuth({
      appId,
      privateKey,
      request: request.defaults({
        baseUrl,
      }),
    });
    this.purpose = purpose;
  }

  async getAppAuthenticationToken() {
    const details = await this._appAuth({ type: 'app' });
    const token = details.token;
    return token;
  }

  private getOrCreateInstallationAuthFunction(installationId: number) {
    let auth = this._installationAuth.get(installationId);
    if (!auth) {
      auth = createAppAuth({
        appId: this._appId,
        privateKey: this.#privateKey,
        installationId,
        request: request.defaults({
          baseUrl: this._baseUrl,
        }),
      });
      this._installationAuth.set(installationId, auth);
    }
    return auth;
  }

  async getInstallationToken(
    installationId: number,
    organizationName: string
  ): Promise<IAuthorizationHeaderValue> {
    const now = new Date();
    const requiredValidityPeriod = new Date(
      now.getTime() + ValidityOffsetAfterNowMilliseconds
    );
    const latestToken = this.getLatestValidToken(
      installationId,
      requiredValidityPeriod
    );
    if (latestToken) {
      return { value: latestToken.headerValue, purpose: this.purpose };
    }
    try {
      const requestedToken = await this.requestInstallationToken(
        installationId,
        organizationName
      );
      this.getInstallationTokens(installationId).push(requestedToken);
      return { value: requestedToken.headerValue, purpose: this.purpose };
    } catch (error) {
      console.warn(
        `Error retrieving installation token ID ${installationId} for organization ${organizationName}`
      );
      throw error;
    }
  }

  private async requestInstallationToken(
    installationId: number,
    organizationName: string
  ): Promise<IInstallationToken> {
    try {
      const requested = new Date();
      const installationAppAuth = this.getOrCreateInstallationAuthFunction(
        installationId
      );
      const installationTokenDetails = await installationAppAuth({
        type: 'installation',
      });
      const installationToken = installationTokenDetails.token;
      const headerValue = `token ${installationToken}`;
      const expiresFromDetails = (installationTokenDetails as any).expiresAt;
      const expires = expiresFromDetails
        ? new Date(expiresFromDetails)
        : new Date(requested.getTime() + InstallationTokenLifetimeMilliseconds);
      const wrapped: IInstallationToken = {
        installationId,
        organizationName,
        requested,
        headerValue,
        expires,
      };
      return wrapped;
    } catch (getTokenError) {
      console.dir(getTokenError);
      throw getTokenError;
    }
  }

  private getLatestValidToken(
    installationId: number,
    timeTokenMustBeValid: Date
  ): IInstallationToken {
    let tokens = this.getInstallationTokens(installationId);
    const count = tokens.length;
    tokens = tokens
      .filter(tokenValidFilter.bind(null, timeTokenMustBeValid))
      .sort(sortByLatestToken);
    if (tokens.length !== count) {
      this.replaceInstallationTokens(installationId, tokens);
    }
    return tokens.length > 0 ? tokens[0] : null;
  }

  private getInstallationTokens(installationId: number): IInstallationToken[] {
    let bin = this._tokensByInstallation.get(installationId);
    if (!bin) {
      bin = [];
      this._tokensByInstallation.set(installationId, bin);
    }
    return bin;
  }

  private replaceInstallationTokens(
    installationId: number,
    arr: IInstallationToken[]
  ) {
    this._tokensByInstallation.set(installationId, arr);
  }
}

function sortByLatestToken(a: IInstallationToken, b: IInstallationToken) {
  if (a > b) {
    return -1;
  } else if (a < b) {
    return 1;
  }
  return 0;
}

function tokenValidFilter(
  timeTokenMustBeValid: Date,
  token: IInstallationToken
) {
  const isValid = token.expires > timeTokenMustBeValid;
  if (!isValid) {
    console.log(
      `invalid or expired token being removed: expires=${token.expires} install_id=${token.installationId} org=${token.organizationName}`
    );
    return false;
  }
  return true;
}
