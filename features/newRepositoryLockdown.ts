//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Operations, ICachedEmployeeInformation } from '../business';
import { Organization } from '../business';
import { Repository, GitHubCollaboratorAffiliationQuery } from '../business';
import { Team } from '../business';
import { ICorporateLink } from '../business/corporateLink';
import { IMail } from '../lib/mailProvider';
import { IRepositoryMetadataProvider } from '../entities/repositoryMetadata/repositoryMetadataProvider';
import {
  RepositoryMetadataEntity,
  GitHubRepositoryVisibility,
  RepositoryLockdownState,
} from '../entities/repositoryMetadata/repositoryMetadata';
import { IndividualContext } from '../user';
import { daysInMilliseconds } from '../utils';

import moment from 'moment';

const botBracket = '[bot]';

interface IMailToRemoveAdministrativeLock {
  organization: Organization;
  repository: Repository;
  linkToClassifyRepository: string;
  linkToDeleteRepository: string;
  mailAddress?: string;
}

interface IMailToLockdownRepo {
  username: string;
  log: string[];
  organization: Organization;
  repository: Repository;
  linkToClassifyRepository: string;
  linkToDeleteRepository: string;
  mailAddress?: string;
  link?: ICorporateLink;
  isForkAdministratorLocked: boolean;
}

export interface INewRepositoryLockdownSystemOptions {
  operations: Operations;
  organization: Organization;
  repository: Repository;
  repositoryMetadataProvider: IRepositoryMetadataProvider;
}

export default class NewRepositoryLockdownSystem {
  organization: Organization;
  operations: Operations;
  repository: Repository;
  repositoryMetadataProvider: IRepositoryMetadataProvider;

  constructor(options: INewRepositoryLockdownSystemOptions) {
    this.organization = options.organization;
    this.operations = options.operations;
    this.repository = options.repository;
    this.repositoryMetadataProvider = options.repositoryMetadataProvider;
  }

  private populateRepositoryMetadata(
    entity: RepositoryMetadataEntity,
    username: string,
    userId: number,
    link: ICorporateLink,
    transferSourceRepositoryLogin: string
  ) {
    entity.createdByThirdPartyUsername = username;
    entity.createdByThirdPartyId = userId.toString();
    if (link) {
      entity.createdByCorporateDisplayName = link.corporateDisplayName;
      entity.createdByCorporateId = link.corporateId;
      entity.createdByCorporateUsername = link.corporateUsername;
    }
    if (transferSourceRepositoryLogin) {
      entity.transferSource = transferSourceRepositoryLogin;
    }
    return entity;
  }

  static async ValidateUserCanSelfDeleteRepository(
    repository: Repository,
    metadata: RepositoryMetadataEntity,
    individualContext: IndividualContext,
    daysAfterCreateToAllowSelfDelete: number
  ): Promise<void> {
    if (
      (!individualContext.corporateIdentity ||
        !individualContext.corporateIdentity.id ||
        !metadata.createdByCorporateId) &&
      (!individualContext.getGitHubIdentity() ||
        !individualContext.getGitHubIdentity().id)
    ) {
      throw new Error(
        'The authenticated user or the linked identity of the repo creator did not have a corporate ID available'
      );
    }
    if (
      (metadata.createdByCorporateId &&
        individualContext.corporateIdentity.id !==
          metadata.createdByCorporateId) ||
      (metadata.createdByThirdPartyId &&
        individualContext.getGitHubIdentity()?.id !==
          metadata.createdByThirdPartyId)
    ) {
      throw new Error(
        'Only the original linked user who first created this repository can classify the repository'
      );
    }
    // any lockdown state is permitted for self-deletes
    const isLockedForkOrNotSetupYet =
      metadata.lockdownState === RepositoryLockdownState.AdministratorLocked ||
      metadata.lockdownState === RepositoryLockdownState.Locked;
    const isWindowOk =
      new Date() <=
      new Date(
        new Date(repository.created_at).getTime() +
          daysInMilliseconds(daysAfterCreateToAllowSelfDelete)
      );
    if (!isWindowOk && !isLockedForkOrNotSetupYet) {
      throw new Error(
        `The ${repository.name} repo was created ${moment(
          repository.created_at
        ).fromNow()}. Repos can only be deleted by their creator ${daysAfterCreateToAllowSelfDelete} days after being created.`
      );
    }
  }

  static async ValidateUserCanConfigureRepository(
    metadata: RepositoryMetadataEntity,
    individualContext: IndividualContext
  ): Promise<void> {
    if (
      (!individualContext.corporateIdentity ||
        !individualContext.corporateIdentity.id ||
        !metadata.createdByCorporateId) &&
      (!individualContext.getGitHubIdentity() ||
        !individualContext.getGitHubIdentity().id)
    ) {
      throw new Error(
        'The authenticated user or the linked identity of the repo creator did not have a corporate ID available'
      );
    }
    if (
      (metadata.createdByCorporateId &&
        individualContext.corporateIdentity.id !==
          metadata.createdByCorporateId) ||
      (metadata.createdByThirdPartyId &&
        individualContext.getGitHubIdentity()?.id !==
          metadata.createdByThirdPartyId)
    ) {
      throw new Error(
        'Only the original linked user who first created this repository can classify the repository'
      );
    }
    if (!metadata.lockdownState) {
      throw new Error('The repository has not been locked down');
    }
    if (metadata.lockdownState === RepositoryLockdownState.Unlocked) {
      throw new Error('The repository has already been unlocked');
    }
    if (
      metadata.lockdownState === RepositoryLockdownState.AdministratorLocked
    ) {
      throw new Error(
        'This repository is locked and requires administrator approval.'
      );
    }
    if (metadata.lockdownState === RepositoryLockdownState.ComplianceLocked) {
      throw new Error(
        'This repository is locked because compliance information is missing.'
      );
    }
    if (metadata.lockdownState !== RepositoryLockdownState.Locked) {
      throw new Error(
        `Unsupported repository lockdown state ${metadata.lockdownState}`
      );
    }
  }

  async removeAdministrativeLock(): Promise<void> {
    if (!this.organization.isNewRepositoryLockdownSystemEnabled()) {
      throw new Error('lockdown system not enabled');
    }
    const repositoryMetadata = await this.repositoryMetadataProvider.getRepositoryMetadata(
      this.repository.id.toString()
    );
    if (
      repositoryMetadata.lockdownState !==
      RepositoryLockdownState.AdministratorLocked
    ) {
      throw new Error(
        `Repository's current lockdown state is not administrative. It is: ${repositoryMetadata.lockdownState}`
      );
    }
    repositoryMetadata.lockdownState = RepositoryLockdownState.Locked;
    await this.repositoryMetadataProvider.updateRepositoryMetadata(
      repositoryMetadata
    );
    let mailSentToCreator = false;
    const lockdownMailContent: IMailToRemoveAdministrativeLock = {
      organization: this.organization,
      repository: this.repository,
      linkToClassifyRepository:
        this.organization.absoluteBaseUrl +
        `wizard?existingreponame=${this.repository.name}&existingrepoid=${this.repository.id}`,
      linkToDeleteRepository: this.repository.absoluteBaseUrl + 'delete',
      mailAddress: null,
    };
    try {
      lockdownMailContent.mailAddress = await this.operations.getMailAddressFromCorporateUsername(
        repositoryMetadata.createdByCorporateUsername
      );
      const repoName = this.repository.name;
      const companyName = this.operations.config.brand.companyName;
      let managerInfo: ICachedEmployeeInformation = null;
      let reasonInfo = `This mail was sent to: ${lockdownMailContent.mailAddress}`;
      try {
        managerInfo = await this.operations.getCachedEmployeeManagementInformation(
          repositoryMetadata.createdByCorporateId
        );
        if (managerInfo && managerInfo.managerMail) {
          reasonInfo += ` and manager ${managerInfo.managerMail}`;
        }
      } catch (managerInfoError) {
        console.dir(managerInfoError);
      }
      const mailToCreator: IMail = {
        to: lockdownMailContent.mailAddress,
        subject: `Your repo was approved, please complete its setup: ${repoName}`,
        content: await this.operations.emailRender('newrepolockremoved', {
          reason: `Your new repo was approved. Additional actions are now required to gain access to continue to use it after classification.
                    ${reasonInfo}.`,
          headline: 'Repo approved',
          notification: 'action',
          app: `${companyName} GitHub`,
          isMailToCreator: true,
          lockdownMailContent,
        }),
      };
      if (managerInfo && managerInfo.managerMail) {
        mailToCreator.cc = managerInfo.managerMail;
      }
      await this.operations.sendMail(mailToCreator);
      mailSentToCreator = true;
    } catch (noLinkOrEmail) {
      console.dir(noLinkOrEmail);
    }
    const notifyMailAddress = this.operations.getRepositoriesNotificationMailAddress();
    const operationsMails = notifyMailAddress ? [notifyMailAddress] : [];
    if (operationsMails && operationsMails.length) {
      try {
        const subject = `Repo approved by an administrator - ${this.organization.name}/${this.repository.name}`;
        const mailToOperations: IMail = {
          to: operationsMails,
          subject,
          content: await this.operations.emailRender('newrepolockremoved', {
            reason: `An administrator has approved this repo, removing an administrative lock. As the operations contact for this system, you are receiving this e-mail.
                      This mail was sent to: ${operationsMails.join(', ')}`,
            headline: `Administrative lock removed: ${this.organization.name}/${this.repository.name}`,
            notification: 'information',
            app: `${this.operations.config.brand.companyName} GitHub`,
            isMailToOperations: true,
            lockdownMailContent,
            mailSentToCreator,
          }),
        };
        await this.operations.sendMail(mailToOperations);
      } catch (mailIssue) {
        console.dir(mailIssue);
      }
    }
    const insights = this.operations.insights;
    if (insights) {
      insights.trackMetric({ name: 'LockedRepoAdminUnlocks', value: 1 });
    }
  }

  async deleteLockedRepository(
    onlyDeleteIfAdministrativeLocked: boolean,
    deletedByUser: boolean
  ): Promise<void> {
    if (!this.organization.isNewRepositoryLockdownSystemEnabled()) {
      throw new Error('lockdown system not enabled');
    }
    const repositoryMetadata = await this.repositoryMetadataProvider.getRepositoryMetadata(
      this.repository.id.toString()
    );
    if (
      onlyDeleteIfAdministrativeLocked &&
      repositoryMetadata.lockdownState !==
        RepositoryLockdownState.AdministratorLocked
    ) {
      throw new Error(
        `Repository's current lockdown state is not administrative. It is: ${repositoryMetadata.lockdownState}`
      );
    }
    const targetType = this.repository.fork ? 'Fork' : 'Repo';
    const repoName = this.repository.name;
    await this.repository.delete();
    try {
      const mailAddress = await this.operations.getMailAddressFromCorporateUsername(
        repositoryMetadata.createdByCorporateUsername
      );
      let managerInfo: ICachedEmployeeInformation = null;
      let reasonInfo = `This mail was sent to: ${mailAddress}`;
      try {
        managerInfo = await this.operations.getCachedEmployeeManagementInformation(
          repositoryMetadata.createdByCorporateId
        );
        if (managerInfo && managerInfo.managerMail) {
          reasonInfo += ` and manager ${managerInfo.managerMail}`;
        }
      } catch (managerInfoError) {
        console.dir(managerInfoError);
      }
      const companyName = this.operations.config.brand.companyName;
      const mailToCreator: IMail = {
        to: mailAddress,
        subject: `${targetType} deleted by ${
          deletedByUser
            ? repositoryMetadata.createdByCorporateUsername
            : 'operations'
        }: ${this.repository.organization.name}/${repoName}`,
        content: await this.operations.emailRender('lockedrepodeleted', {
          reason: `The ${targetType.toLowerCase()} was deleted. ${reasonInfo}.`,
          headline: `${targetType} deleted`,
          notification: 'information',
          app: `${companyName} GitHub`,
          isMailToCreator: true,
          deletedByUser,
          isFork: this.repository.fork,
          creator: repositoryMetadata.createdByCorporateUsername,
          repository: this.repository,
        }),
      };
      if (managerInfo && managerInfo.managerMail) {
        mailToCreator.cc = managerInfo.managerMail;
      }
      await this.operations.sendMail(mailToCreator);
    } catch (noLinkOrEmail) {
      console.dir(noLinkOrEmail);
    }
    const operationsMails = [
      this.operations.getRepositoriesNotificationMailAddress(),
    ];
    if (operationsMails && operationsMails.length) {
      try {
        const mailToOperations: IMail = {
          to: operationsMails,
          subject: `${targetType} deleted by ${
            deletedByUser
              ? repositoryMetadata.createdByCorporateUsername
              : 'operations'
          }: ${this.repository.organization.name}/${repoName}`,
          content: await this.operations.emailRender('lockedrepodeleted', {
            reason: `A decision has been made to delete this repo.
                      This mail was sent to operations at: ${operationsMails.join(
                        ', '
                      )}`,
            headline: `${targetType} deleted`,
            isFork: this.repository.fork,
            notification: 'information',
            deletedByUser,
            app: `${this.operations.config.brand.companyName} GitHub`,
            isMailToOperations: true,
            creator: repositoryMetadata.createdByCorporateUsername,
            repository: this.repository,
          }),
        };
        await this.operations.sendMail(mailToOperations);
      } catch (mailIssue) {
        console.dir(mailIssue);
      }
    }
    const insights = this.operations.insights;
    if (insights) {
      insights.trackMetric({ name: 'LockedRepoDeletes', value: 1 });
      insights.trackMetric({
        name: deletedByUser
          ? 'LockedRepoUserDeletes'
          : 'LockedRepoAdminDeletes',
        value: 1,
      });
    }
  }

  async lockdownIfNecessary(
    action: 'created' | 'transferred',
    username: string,
    thirdPartyId: number,
    transferSourceRepositoryLogin: string
  ): Promise<boolean> {
    const lockdownLog: string[] = [];
    // reconfirm that the new repository system is enabled for this organization
    if (!this.organization.isNewRepositoryLockdownSystemEnabled()) {
      return false;
    }
    const lockdownForks = this.organization.isForkLockdownSystemEnabled();
    const lockdownTransfers = this.organization.isTransferLockdownSystemEnabled();
    lockdownLog.push(
      `Confirmed that the ${this.organization.name} organization has opted in to the new repository lockdown system`
    );
    if (lockdownForks) {
      lockdownLog.push(
        'Confirmed that the additional fork lockdown feature is enabled for this org'
      );
    }
    if (lockdownTransfers) {
      lockdownLog.push(
        'Confirmed that the additional transfer lockdown feature is enabled for this org'
      );
    }
    let isTransfer = action === 'transferred';
    if (isTransfer && !lockdownTransfers) {
      return false; // no need to do special transfer logic
    }
    if (isTransfer && transferSourceRepositoryLogin) {
      try {
        const sourceOrganization = this.operations.getOrganization(
          transferSourceRepositoryLogin
        );
        // BUSINESS RULE: If the organization is configured in the system, no need to lock it down...
        // CONSIDER: should there be a feature flag for this behavior, to allow managed-to-managed org transfers without lockdown?
        // CONSIDER: notify operations that a transfer happened
        return false;
      } catch (notManagedOrganization) {
        /* ignore */
      }
    }
    const lowercaseUsername = username.toLowerCase();
    // any repository created by a bot *is ok* and will not be locked down. If this is an issue, having an approved list of permitted bots to create repos would be one way to approach this loophole. Non-bot users cannot have brackets in their names.
    if (lowercaseUsername.includes(botBracket)) {
      // CONSIDER: send operations an e-mail FYI when a bot account is used?
      return false;
    }
    lockdownLog.push(
      `Confirmed that the repository was not ${action} by a bot`
    );
    // a repository created by one of the operations accounts in the allowed list is OK and will not be locked down
    const systemAccounts = new Set(
      this.operations.systemAccountsByUsername.map((username) =>
        username.toLowerCase()
      )
    );
    if (systemAccounts.has(lowercaseUsername)) {
      return false;
    }
    lockdownLog.push(
      `Confirmed that the repository was not ${action} by any of the system accounts: ${Array.from(
        systemAccounts.values()
      ).join(', ')}`
    );
    await this.lockdownRepository(lockdownLog, systemAccounts);
    let link: ICorporateLink = null;
    try {
      link = await this.operations.getLinkByThirdPartyId(
        thirdPartyId.toString()
      );
    } catch (noLinkError) {
      lockdownLog.push(
        `No corporate link available for the GitHub username ${username} that created the repository`
      );
    }
    let isForkAdministratorLocked = false;
    try {
      // Repository metadata is used to lock down the security of the repository system. Only
      // a complete system administrator or the initial creator of a repository is able to
      // complete the initial repository setup process.
      let repositoryMetadata: RepositoryMetadataEntity = null;
      try {
        repositoryMetadata = await this.repositoryMetadataProvider.getRepositoryMetadata(
          this.repository.id.toString()
        );
      } catch (doesNotExist) {
        // ignore: 404 is standard here
      }
      let lockdownState = RepositoryLockdownState.Locked;
      if (action === 'created' && this.repository.fork && lockdownForks) {
        lockdownState = RepositoryLockdownState.AdministratorLocked;
        isForkAdministratorLocked = true;
        lockdownLog.push(
          'The repository is a fork and will be administrator locked'
        );
      }
      if (repositoryMetadata) {
        lockdownLog.push(
          `Repository metadata already exists for repository ID ${this.repository.id}`
        );
        const updateMetadata = this.populateRepositoryMetadata(
          repositoryMetadata,
          username,
          thirdPartyId,
          link,
          transferSourceRepositoryLogin
        );
        await this.repositoryMetadataProvider.updateRepositoryMetadata(
          updateMetadata
        );
        lockdownLog.push(
          `Updated the repository metadata with username and link information`
        );
      } else {
        repositoryMetadata = this.populateRepositoryMetadata(
          new RepositoryMetadataEntity(),
          username,
          thirdPartyId,
          link,
          transferSourceRepositoryLogin
        );
        repositoryMetadata.created = new Date();
        repositoryMetadata.lockdownState = lockdownState;
        repositoryMetadata.repositoryId = this.repository.id.toString();
        repositoryMetadata.repositoryName = this.repository.name;
        repositoryMetadata.organizationName = this.organization.name;
        repositoryMetadata.organizationId = this.organization.id.toString();
        repositoryMetadata.initialRepositoryDescription = this.repository.description;
        repositoryMetadata.initialRepositoryVisibility = this.repository.private
          ? GitHubRepositoryVisibility.Private
          : GitHubRepositoryVisibility.Public;
        await this.repositoryMetadataProvider.createRepositoryMetadata(
          repositoryMetadata
        );
        lockdownLog.push(
          `Created the initial repository metadata indicating the repo was created by ${username}`
        );
      }
    } catch (metadataSystemError) {
      console.dir(metadataSystemError);
      lockdownLog.push(
        `While writing repository metadata an error: ${metadataSystemError.message}`
      );
    }
    let mailSentToCreator = false;
    const operationsMails = [
      this.operations.getRepositoriesNotificationMailAddress(),
    ];
    const lockdownMailContent: IMailToLockdownRepo = {
      username,
      log: lockdownLog,
      organization: this.organization,
      repository: this.repository,
      linkToDeleteRepository: this.repository.absoluteBaseUrl + 'delete',
      linkToClassifyRepository: `${this.organization.absoluteBaseUrl}wizard?existingreponame=${this.repository.name}&existingrepoid=${this.repository.id}`,
      mailAddress: null,
      link,
      isForkAdministratorLocked,
    };
    lockdownLog.push(
      `The repo can be unlocked at ${lockdownMailContent.linkToClassifyRepository}`
    );
    const repoActionType = this.repository.fork
      ? 'forked'
      : isTransfer
      ? 'transferred'
      : 'created';
    const stateVerb = isTransfer ? 'transferred' : 'new';
    if (link) {
      try {
        const mailAddress =
          link.corporateMailAddress ||
          (await this.operations.getMailAddressFromCorporateUsername(
            link.corporateUsername
          ));
        const repoName = this.repository.name;
        const subject = isForkAdministratorLocked
          ? `Your new fork requires administrator approval: ${repoName} (${username})`
          : `Please complete the setup of your ${stateVerb} GitHub repository ${repoName} (${username})`;
        if (mailAddress) {
          lockdownMailContent.mailAddress = mailAddress;
          const companyName = this.operations.config.brand.companyName;
          let managerInfo: ICachedEmployeeInformation = null;
          let reasonInfo = `This mail was sent to: ${mailAddress}`;
          try {
            const providers = this.operations.providers;
            let shouldTryNotifyManager = true;
            if (providers?.customizedNewRepositoryLogic) {
              // this is a hack around the new repo custom logic
              const customContext = providers.customizedNewRepositoryLogic.createContext(
                {} /* "request" */
              );
              shouldTryNotifyManager = providers.customizedNewRepositoryLogic.shouldNotifyManager(
                customContext,
                link.corporateId
              );
            }
            if (shouldTryNotifyManager) {
              managerInfo = await this.operations.getCachedEmployeeManagementInformation(
                link.corporateId
              );
              if (managerInfo && managerInfo.managerMail) {
                reasonInfo += ` and manager ${managerInfo.managerMail}`;
              }
            }
          } catch (managerInfoError) {
            console.dir(managerInfoError);
          }
          const mailToCreator: IMail = {
            to: mailAddress,
            subject,
            content: await this.operations.emailRender('newrepolockdown', {
              reason: `You just ${repoActionType} a repository on GitHub and have additional actions required to gain access to continue to use it after classification.
                        ${reasonInfo}.`,
              headline: isForkAdministratorLocked
                ? 'Fork approval required'
                : `Setup your ${stateVerb} repository`,
              notification: isForkAdministratorLocked
                ? 'action'
                : 'information',
              app: `${companyName} GitHub`,
              isMailToCreator: true,
              lockdownMailContent,
              isForkAdministratorLocked,
              action,
              operationsMail: operationsMails.join(','),
              transferSourceRepositoryLogin,
            }),
          };
          if (managerInfo && managerInfo.managerMail) {
            mailToCreator.cc = managerInfo.managerMail;
          }
          await this.operations.sendMail(mailToCreator);
          lockdownLog.push(
            `sent an e-mail to the person who ${repoActionType} the repository ${mailAddress} (corporate username: ${link.corporateUsername})`
          );
          mailSentToCreator = true;
        } else {
          lockdownLog.push(
            `no e-mail address available for the corporate username ${link.corporateUsername}`
          );
        }
      } catch (noLinkOrEmail) {
        console.dir(noLinkOrEmail);
      }
    }
    if (operationsMails) {
      try {
        const subject = isForkAdministratorLocked
          ? `New fork ${this.organization.name}/${this.repository.name} requires approval - forked by ${username}`
          : `Repository ${repoActionType}: ${this.organization.name}/${this.repository.name} (by ${username})`;
        const mailToOperations: IMail = {
          to: operationsMails,
          subject,
          content: await this.operations.emailRender('newrepolockdown', {
            reason: `A user just ${repoActionType} this repository directly on GitHub. As the operations contact for this system, you are receiving this e-mail.
                      This mail was sent to: ${operationsMails.join(', ')}`,
            headline: isForkAdministratorLocked
              ? `Fork ${this.organization.name}/${this.repository.name} by ${username}`
              : `Repo (${stateVerb}) ${this.organization.name}/${this.repository.name} ${repoActionType} by ${username}`,
            notification: 'information',
            app: `${this.operations.config.brand.companyName} GitHub`,
            isMailToOperations: true,
            lockdownMailContent,
            transferSourceRepositoryLogin,
            action,
            mailSentToCreator,
            isForkAdministratorLocked,
          }),
        };
        await this.operations.sendMail(mailToOperations);
        lockdownLog.push(
          `sent an e-mail to the operations contact(s): ${operationsMails.join(
            ', '
          )}`
        );
      } catch (mailIssue) {
        console.dir(mailIssue);
      }
    }
    const insights = this.operations.insights;
    if (insights) {
      insights.trackMetric({ name: 'LockedRepos', value: 1 });
      let metricName = isForkAdministratorLocked
        ? 'LockedForks'
        : 'LockedDirectRepos';
      if (isTransfer) {
        metricName = 'LockedTransfers';
      }
      insights.trackMetric({ name: metricName, value: 1 });
    }
    console.dir(lockdownLog);
    return true;
  }

  async lockdownRepository(
    log: string[],
    systemAccounts: Set<string>
  ): Promise<void> {
    try {
      const specialPermittedTeams = new Set([
        ...this.organization.specialRepositoryPermissionTeams.admin,
        ...this.organization.specialRepositoryPermissionTeams.write,
        ...this.organization.specialRepositoryPermissionTeams.read,
      ]);
      const teamPermissions = await this.repository.getTeamPermissions();
      for (const tp of teamPermissions) {
        if (specialPermittedTeams.has(tp.team.id)) {
          log.push(
            `Special permitted team id=${tp.team.id} name=${tp.team.name} will continue to have repository access`
          );
        } else {
          await this.tryDropTeam(this.repository, tp.team, log);
        }
      }
      const collaborators = await this.repository.getCollaborators({
        affiliation: GitHubCollaboratorAffiliationQuery.Direct,
      });
      for (const collaborator of collaborators) {
        if (systemAccounts.has(collaborator.login.toLowerCase())) {
          log.push(
            `System account ${collaborator.login} will continue to have repository access`
          );
        } else {
          await this.tryDropCollaborator(
            this.repository,
            collaborator.login,
            log
          );
        }
      }
      log.push('Lockdown of permissions complete');
    } catch (lockdownError) {
      log.push(
        `Error while locking down the repository: ${lockdownError.message}`
      );
    }
  }

  async tryDropTeam(
    repository: Repository,
    team: Team,
    log: string[]
  ): Promise<void> {
    try {
      await repository.removeTeamPermission(team.id);
      log.push(
        `Lockdown removed team id=${team.id} name=${team.name} from the repository ${repository.name} in organization ${repository.organization.name}`
      );
    } catch (lockdownError) {
      log.push(
        `Error while removing team id=${team.id} name=${team.name} permission from the repository ${repository.name} in organization ${repository.organization.name}: ${lockdownError.message}`
      );
    }
  }

  async tryDropCollaborator(
    repository: Repository,
    login: string,
    log: string[]
  ): Promise<void> {
    try {
      await repository.removeCollaborator(login);
      log.push(
        `Lockdown removed collaborator login=${login} from the repository ${repository.name} in organization ${repository.organization.name}`
      );
    } catch (lockdownError) {
      log.push(
        `Error while removing collaborator login=${login} from the repository ${repository.name} in organization ${repository.organization.name}: ${lockdownError.message}`
      );
    }
  }
}
