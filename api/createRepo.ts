//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// NOTE: this file at this time is Microsoft-specific and needs to be updated
// and refactored to be useful by others. There are values stored in
// configuration that can be used instead of the hardcoded values within.

import _ from 'lodash';
import { jsonError } from '../middleware';
import {
  getProviders,
  ICustomizedNewRepoProperties,
  ICustomizedNewRepositoryLogic,
  INewRepositoryContext,
  ReposAppRequest,
} from '../transitional';
import { ICreateRepositoryResult, Organization } from '../business';
import {
  RepositoryMetadataEntity,
  GitHubRepositoryVisibility,
  GitHubRepositoryPermission,
  RepositoryLockdownState,
} from '../entities/repositoryMetadata/repositoryMetadata';
import RenderHtmlMail from '../lib/emailRender';

import {
  RepoWorkflowEngine,
  IRepositoryWorkflowOutput,
  IApprovalPackage,
} from '../routes/org/repoWorkflowEngine';
import { IMailProvider } from '../lib/mailProvider';
import { IndividualContext } from '../user';
import NewRepositoryLockdownSystem from '../features/newRepositoryLockdown';
import { ICachedEmployeeInformation } from '../business';
import { Repository } from '../business';
import { ICorporateLink } from '../business';
import getCompanySpecificDeployment from '../middleware/companySpecificDeployment';

const supportedLicenseExpressions = [
  'mit',
  '(mit and cc-by-4.0)',
  'cc-by-4.0',
  'other',
];

const hardcodedApprovalTypes = [
  'NewReleaseReview',
  'ExistingReleaseReview',
  'SmallLibrariesToolsSamples',
  'Migrate',
  'Exempt',
];

export interface ICreateRepositoryApiResult {
  github: any;
  name: string;
  repositoryId: number;
  organizationName: string;
  notified?: string[];
}

export enum CreateRepositoryEntrypoint {
  Api = 'api',
  Client = 'client',
}

export interface IReposAppRequestWithCreateResponse extends ReposAppRequest {
  repoCreateResponse?: ICreateRepositoryApiResult;
}

export async function CreateRepository(
  req,
  logic: ICustomizedNewRepositoryLogic,
  createContext: INewRepositoryContext,
  bodyOverride: unknown,
  entrypoint: CreateRepositoryEntrypoint,
  individualContext?: IndividualContext
): Promise<ICreateRepositoryApiResult> {
  if (!req.organization) {
    throw jsonError(new Error('No organization available in the route.'), 400);
  }
  const providers = getProviders(req);
  const {
    operations,
    mailProvider,
    repositoryMetadataProvider,
    insights,
  } = providers;
  const ourFields = [
    'ms.onBehalfOf',
    'ms.license',
    'ms.approval',
    'ms.approval-url',
    'ms.justification',
    'ms.notify',
    'ms.administrators',
    'ms.teams',
    'ms.template',
    'ms.project-type',
  ];
  const properties = {};
  const parameters = bodyOverride || req.body;
  ourFields.forEach((fieldName) => {
    if (parameters[fieldName] !== undefined) {
      properties[fieldName] = parameters[fieldName];
      delete parameters[fieldName];
    }
  });
  const msProperties = {
    onBehalfOf: properties['ms.onBehalfOf'] || req.headers['ms-onbehalfof'],
    justification:
      properties['ms.justification'] || req.headers['ms-justification'],
    license: properties['ms.license'] || req.headers['ms-license'],
    approvalType: properties['ms.approval'] || req.headers['ms-approval'],
    approvalUrl:
      properties['ms.approval-url'] || req.headers['ms-approval-url'],
    notify: properties['ms.notify'] || req.headers['ms-notify'],
    administrators:
      properties['ms.administrators'] || req.headers['ms-administrators'],
    teams: properties['ms.teams'] || req.headers['ms-teams'],
    template: properties['ms.template'] || req.headers['ms-template'],
    projectType:
      properties['ms.project-type'] || req.headers['ms-project-type'],
  };
  // Validate licenses when present
  let msLicense = msProperties.license;
  if (msLicense) {
    msLicense = msLicense.toLowerCase();
    if (supportedLicenseExpressions.indexOf(msLicense) < 0) {
      throw jsonError(
        new Error('The provided license expression is not currently supported'),
        422
      );
    }
  }
  if (
    msProperties.administrators &&
    !Array.isArray(msProperties.administrators)
  ) {
    throw jsonError(
      new Error('Administrators must be an array of logins'),
      422
    );
  }
  // Validate approval types
  const msApprovalType = msProperties.approvalType;
  if (!msApprovalType) {
    throw jsonError(
      new Error('Missing corporate approval type information'),
      422
    );
  }
  if (hardcodedApprovalTypes.indexOf(msApprovalType) < 0) {
    throw jsonError(
      new Error('The provided approval type is not supported'),
      422
    );
  }
  // Validate specifics of what is in the approval
  switch (msApprovalType) {
    case 'NewReleaseReview':
    case 'ExistingReleaseReview':
      if (!msProperties.approvalUrl) {
        throw jsonError(
          new Error(
            'Approval URL for the release review is required when using the release review approval type'
          ),
          422
        );
      }
      break;
    case 'SmallLibrariesToolsSamples':
      break;
    case 'Exempt':
      if (!msProperties.justification) {
        throw jsonError(
          new Error(
            'Justification is required when using the exempted approval type'
          ),
          422
        );
      }
      break;
    default:
      throw jsonError(
        new Error('The requested approval type is not currently supported.'),
        422
      );
  }
  parameters.org = req.organization.name;
  const organization = operations.getOrganization(parameters.org);
  const existingRepoId = req.body.existingrepoid;
  let metadata: RepositoryMetadataEntity;
  let response: any = null;
  if (existingRepoId && !organization.isNewRepositoryLockdownSystemEnabled()) {
    throw jsonError(
      new Error(
        `Repository ID ${existingRepoId} provided for a repository within the ${organization.name} org that is not configured for existing repository classification`
      ),
      422
    );
  }
  let repository: Repository = null;
  if (!existingRepoId) {
    providers.insights?.trackEvent({
      name: 'ApiRepoTryCreateForOrg',
      properties: {
        parameterName: parameters.name,
        description: parameters.description,
        private: parameters.private,
        org: parameters.org,
        entrypoint,
      },
    });
    let createResult: ICreateRepositoryResult = null;
    try {
      createResult = await organization.createRepository(
        parameters.name,
        parameters
      );
      if (createResult && createResult.repository) {
        repository = organization.repositoryFromEntity(createResult.repository);
      }
    } catch (error) {
      providers.insights?.trackEvent({
        name: 'ApiRepoCreateForOrgGitHubFailure',
        properties: {
          parameterName: parameters.name,
          private: parameters.private,
          org: parameters.org,
          parameters: JSON.stringify(parameters),
          entrypoint,
        },
      });
      if (error && error.innerError) {
        const inner = error.innerError;
        req.insights.trackException({
          exception: inner,
          properties: {
            event: 'ApiRepoCreateGitHubErrorInside',
            message: inner && inner.message ? inner.message : inner,
            status: inner && inner.status ? inner.status : '',
            statusCode: inner && inner.statusCode ? inner.statusCode : '',
          },
        });
      }
      req.insights.trackException({
        exception: error,
        properties: {
          event: 'ApiRepoCreateGitHubError',
          message: error && error.message ? error.message : error,
        },
      });
      throw jsonError(error, error.status || 500);
    }
    response = createResult.response;
    providers.insights?.trackEvent({
      name: 'ApiRepoCreateForOrg',
      properties: {
        parameterName: parameters.name,
        description: parameters.description,
        private: parameters.private,
        org: parameters.org,
        entrypoint,
        result: JSON.stringify(response),
      },
    });
    // strip an internal "cost" part off our response object
    delete response.cost;
    // from this point on any errors should roll back
    const repoCreateResponse: ICreateRepositoryApiResult = {
      github: response,
      name: response && response.name ? response.name : undefined,
      organizationName: repository.organization.name,
      repositoryId: response?.id ? Number(response.id) : undefined,
    };
    req.repoCreateResponse = repoCreateResponse;
    metadata = new RepositoryMetadataEntity();
    // Store create metadata
    metadata.created = new Date();
    metadata.createdByThirdPartyUsername = msProperties.onBehalfOf;
    if (individualContext && individualContext.corporateIdentity.id) {
      metadata.createdByCorporateId = individualContext.corporateIdentity.id;
      metadata.createdByCorporateUsername =
        individualContext.corporateIdentity.username;
      metadata.createdByCorporateDisplayName =
        individualContext.corporateIdentity.displayName;
    }
    // TODO: we also want to store corporate manager information, eventually
    try {
      const account = await operations.getAccountByUsername(
        metadata.createdByThirdPartyUsername
      );
      metadata.createdByThirdPartyId = account.id.toString();
    } catch (noAvailableUsername) {
      providers.insights?.trackEvent({
        name: 'ApiRepoCreateInvalidUsername',
        properties: {
          username: metadata.createdByThirdPartyUsername,
          error: noAvailableUsername.message,
          encodedError: JSON.stringify(noAvailableUsername),
        },
      });
    }
    metadata.repositoryName = response.name;
    metadata.repositoryId = response.id;
    metadata.initialRepositoryDescription = response.description;
    metadata.initialRepositoryVisibility = response.private
      ? GitHubRepositoryVisibility.Private
      : GitHubRepositoryVisibility.Public;
    metadata.organizationName = req.organization.name.toLowerCase();
    if (organization.id) {
      metadata.organizationId = organization.id.toString();
    }
  } else {
    // Locked down new repo, this is a classification call
    try {
      metadata = await repositoryMetadataProvider.getRepositoryMetadata(
        existingRepoId
      );
    } catch (existingError) {
      if (existingError.status && existingError.status === 404) {
        throw new Error(
          `The existing repository with id=${existingRepoId} cannot be classified as it was not processed as a new repository`
        );
      } else {
        throw existingError;
      }
    }
    // Verify that the active user is the same person who created it
    if (!individualContext) {
      throw new Error(
        'Existing repository reclassification requires an authenticated identity'
      );
    }
    NewRepositoryLockdownSystem.ValidateUserCanConfigureRepository(
      metadata,
      individualContext
    );
    // CONSIDER: or a org sudo user or a portal administrator
    const repositoryByName = organization.repository(metadata.repositoryName);
    const response = await repositoryByName.getDetails();
    if (response.id != /* loose */ existingRepoId) {
      throw new Error(
        `The ID of the repo ${metadata.repositoryName} does not match ${existingRepoId}`
      );
    }
    repository = repositoryByName;
    metadata.lockdownState = RepositoryLockdownState.Unlocked;
    const repoCreateResponse: ICreateRepositoryApiResult = {
      github: response,
      name: response && response.name ? response.name : undefined,
      repositoryId: Number(existingRepoId),
      organizationName: repository.organization.name,
    };
    req.repoCreateResponse = repoCreateResponse;
  }
  metadata.releaseReviewJustification = msProperties.justification;
  metadata.initialLicense = msProperties.license;
  metadata.releaseReviewType = msProperties.approvalType;
  metadata.releaseReviewUrl = msProperties.approvalUrl;
  metadata.initialTemplate = msProperties.template;
  if (msProperties.administrators) {
    metadata.initialAdministrators = msProperties.administrators;
  }
  metadata.projectType = msProperties.projectType;
  metadata.initialCorrelationId = req.correlationId;
  // team permissions
  const teamTypes = ['pull', 'push', 'admin'];
  const typeValues = [
    GitHubRepositoryPermission.Pull,
    GitHubRepositoryPermission.Push,
    GitHubRepositoryPermission.Admin,
  ];
  downgradeBroadAccessTeams(organization, msProperties.teams || {});
  for (let i = 0; msProperties.teams && i < teamTypes.length; i++) {
    const teamType = teamTypes[i];
    const enumValue = typeValues[i];
    const idList = msProperties.teams[teamType];
    if (idList && idList.length) {
      for (let j = 0; j < idList.length; j++) {
        metadata.initialTeamPermissions.push({
          permission: enumValue,
          teamId: idList[j],
        });
      }
    }
  }
  let entityId = existingRepoId || null;
  try {
    if (!entityId) {
      entityId = await repositoryMetadataProvider.createRepositoryMetadata(
        metadata
      );
    }
  } catch (insertRequestError) {
    const err = jsonError(
      new Error(
        `Rolling back, problems creating repo metadata for ${metadata.repositoryName} and repo ${metadata.repositoryId}`
      ),
      500
    );
    req.insights.trackException({
      exception: insertRequestError,
      properties: {
        event: 'ApiRepoCreateRollbackError',
        message:
          insertRequestError && insertRequestError.message
            ? insertRequestError.message
            : insertRequestError,
      },
    });
    if (!req.organization || !metadata || !metadata.repositoryName) {
      throw insertRequestError; // if GitHub never returned
    }
    const newlyCreatedRepo = (req.organization as Organization).repository(
      metadata.repositoryName
    );
    await newlyCreatedRepo.delete();
    throw err;
  }
  if (existingRepoId) {
    try {
      await repositoryMetadataProvider.updateRepositoryMetadata(metadata);
    } catch (updateError) {
      throw updateError;
    }
  }
  const approvalPackage: IApprovalPackage = {
    id: entityId,
    repositoryMetadata: metadata,
    createResponse: response,
    isUnlockingExistingRepository: existingRepoId,
    isFork: response ? response.fork : false,
    isTransfer: metadata && metadata.transferSource ? true : false,
    // TEMPORARY:
    createEntrypoint: entrypoint,
    renameDefaultBranchTo: null,
    renameDefaultBranchExcludeIfApiCall: null,
  };
  // TEMPORARY FEATURE: through end of 2020, default branch work
  const config = providers.config;
  if (config?.github?.newRepos?.renameDefaultBranchFlag === true) {
    approvalPackage.renameDefaultBranchTo =
      config.github.newRepos.renameDefaultBranchTo;
    approvalPackage.renameDefaultBranchExcludeIfApiCall =
      config.github.newRepos.renameDefaultBranchExcludeApiCreates;
  }
  // END TEMPORARY FEATURE: default branch work
  // req.approvalRequest['ms.approvalId'] = requestId; // TODO: is this ever used?
  const repoWorkflow = new RepoWorkflowEngine(
    req.organization as Organization,
    approvalPackage
  );
  let output = [];
  try {
    output = await generateAndRunSecondaryTasks(repoWorkflow);
  } catch (rollbackNeededError) {
    console.dir(rollbackNeededError);
  }
  if (!req.repoCreateResponse) {
    req.repoCreateResponse = { tasks: null };
  }
  req.repoCreateResponse.tasks = output;
  if (logic?.afterRepositoryCreated) {
    try {
      await logic.afterRepositoryCreated(
        createContext,
        individualContext?.corporateIdentity?.id,
        req.repoCreateResponse
      );
    } catch (ignoredCustomError) {
      insights?.trackException({ exception: ignoredCustomError });
    }
  }
  if (msProperties.notify && mailProvider) {
    try {
      let createdUserLink: ICorporateLink = individualContext?.link || null;
      if (!createdUserLink && providers.linkProvider) {
        try {
          createdUserLink = await providers.linkProvider.getByThirdPartyId(
            repoWorkflow.request.createdByThirdPartyId
          );
        } catch (linkError) {
          console.log(
            `Ignored link error during new repo notification: ${linkError}`
          );
        }
      }
      await sendEmail(
        req,
        logic,
        createContext,
        mailProvider,
        req.apiKeyRow,
        req.correlationId,
        output,
        repoWorkflow.request,
        msProperties,
        existingRepoId,
        repository,
        createdUserLink
      );
    } catch (mailSendError) {
      insights?.trackException({ exception: mailSendError });
      console.dir(mailSendError);
    }
  }
  return req.repoCreateResponse;
}

async function generateAndRunSecondaryTasks(
  repoWorkflow: RepoWorkflowEngine
): Promise<IRepositoryWorkflowOutput[]> {
  const results = await repoWorkflow.executeNewRepositoryChores();
  // NOTE: no longer failing with any errors
  return results;
}

function downgradeBroadAccessTeams(organization, teams) {
  const broadAccessTeams = new Set(organization.broadAccessTeams);
  if (teams.admin && Array.isArray(teams.admin)) {
    _.remove(teams.admin, (teamId) => {
      if (broadAccessTeams.has(teamId)) {
        if (!teams.pull) {
          teams.pull = [];
        }
        teams.pull.push(teamId);
        return true;
      }
      return false;
    });
  }
  if (teams.pull && Array.isArray(teams.pull)) {
    teams.pull = _.uniq(teams.pull); // deduplicate
  }
}

async function sendEmail(
  req: IReposAppRequestWithCreateResponse,
  logic: ICustomizedNewRepositoryLogic,
  createContext: INewRepositoryContext,
  mailProvider: IMailProvider,
  apiKeyRow,
  correlationId: string,
  repoCreateResults,
  approvalRequest: RepositoryMetadataEntity,
  msProperties,
  existingRepoId: any,
  repository: Repository,
  createdUserLink: ICorporateLink
): Promise<void> {
  const { config, insights, operations, viewServices } = getProviders(req);
  const deployment = getCompanySpecificDeployment();
  const emailTemplate =
    deployment?.views?.email?.repository?.new || 'newRepository';
  const excludeNotificationsValue =
    config.notifications?.reposNotificationExcludeForUsers;
  let excludeNotifications = [];
  if (excludeNotificationsValue) {
    excludeNotifications = excludeNotificationsValue.split(',');
  }
  if (
    approvalRequest.createdByCorporateUsername &&
    excludeNotifications &&
    excludeNotifications.includes(
      approvalRequest.createdByCorporateUsername.toLowerCase()
    )
  ) {
    return;
  }
  const emails =
    msProperties?.notify && (msProperties.notify as string).split(',');
  let targetType = repoCreateResults.fork ? 'Fork' : 'Repo';
  if (!repoCreateResults.fork && approvalRequest.transferSource) {
    targetType = 'Transfer';
  }
  let managerInfo: ICachedEmployeeInformation = null;
  if (approvalRequest.createdByCorporateId) {
    try {
      managerInfo = await operations.getCachedEmployeeManagementInformation(
        approvalRequest.createdByCorporateId
      );
    } catch (ignoreError) {
      console.dir(ignoreError);
    }
  }
  let headline = `${targetType} ready`;
  const serviceShortName =
    apiKeyRow && apiKeyRow.service ? apiKeyRow.service : undefined;
  let subject = serviceShortName
    ? `${
        approvalRequest.repositoryName
      } ${targetType.toLowerCase()} created by ${serviceShortName}`
    : `${approvalRequest.repositoryName} ${targetType.toLowerCase()} created`;
  if (existingRepoId) {
    subject = `${
      approvalRequest.repositoryName
    } ${targetType.toLowerCase()} ready`;
  }
  const displayHostname = req.hostname;
  const approvalScheme =
    displayHostname === 'localhost' && config.webServer.allowHttp === true
      ? 'http'
      : 'https';
  const reposSiteBaseUrl = `${approvalScheme}://${displayHostname}/`;
  if (repository) {
    try {
      await repository.getDetails();
    } catch (getDetailsErrorIgnored) {
      console.dir(getDetailsErrorIgnored);
    }
  }
  let additionalViewProperties: ICustomizedNewRepoProperties = null;
  try {
    if (createContext && logic) {
      additionalViewProperties = await logic.getNewMailViewProperties(
        createContext,
        repository
      );
    }
  } catch (err) {
    insights?.trackException({ exception: err });
    console.warn(err);
  }
  const mail = {
    to: [...emails, ...additionalViewProperties?.to],
    cc: additionalViewProperties?.cc,
    bcc: additionalViewProperties?.bcc,
    subject,
    correlationId,
    content: undefined,
  };
  if (managerInfo && managerInfo.managerMail) {
    let shouldSend = true;
    if (createContext && logic) {
      shouldSend = logic.shouldNotifyManager(
        createContext,
        approvalRequest.createdByCorporateId
      );
    }
    if (shouldSend) {
      if (mail.cc) {
        mail.cc.push(managerInfo.managerMail);
      } else {
        mail.cc = [managerInfo.managerMail];
      }
    }
  }
  const app = config.brand?.companyName
    ? `${config.brand.companyName} GitHub`
    : 'GitHub';
  const contentOptions = Object.assign(
    additionalViewProperties?.viewProperties ||
      {} /* allow a custom provider to override */,
    {
      reason: `You are receiving this e-mail because the new repository request included the e-mail notification address(es) ${msProperties.notify}, or, you are the manager of the person who created the repo.`,
      headline,
      notification: 'information',
      app,
      correlationId,
      approvalRequest, // old name
      repositoryMetadataEntity: approvalRequest,
      repository,
      organization: repository ? repository.organization : null,
      createdUserLink,
      existingRepoId,
      results: repoCreateResults,
      version: config.logging.version,
      managerInfo,
      reposSiteUrl: reposSiteBaseUrl,
      liveReposSiteUrl: config.urls ? config.urls.repos : null,
      api: serviceShortName, // when used by the client single-page app, this is not considered an API call
      service: serviceShortName,
      serviceOwner: apiKeyRow ? apiKeyRow.owner : undefined,
      serviceDescription: apiKeyRow ? apiKeyRow.description : undefined,
      viewServices,
      isNotBootstrap: true,
    }
  );
  try {
    mail.content = await RenderHtmlMail(
      config.typescript.appDirectory,
      emailTemplate,
      contentOptions
    );
  } catch (renderError) {
    req.insights.trackException({
      exception: renderError,
      properties: {
        content: contentOptions,
        eventName: 'ApiRepoCreateMailRenderFailure',
      },
    });
    throw renderError;
  }
  const customData = {
    content: contentOptions,
    receipt: null,
    eventName: undefined,
  };
  const additionalMail = { ...mail };
  try {
    insights?.trackEvent({
      name: 'ApiRepoSendMail',
      properties: {
        to: JSON.stringify(mail.to || ''),
        cc: JSON.stringify(mail.to || ''),
        bcc: JSON.stringify(mail.to || ''),
      },
    });
    customData.receipt = await mailProvider.sendMail(mail);
    insights?.trackEvent({
      name: 'ApiRepoCreateMailSuccess',
      properties: customData,
    });
    req.repoCreateResponse.notified = emails;
  } catch (mailError) {
    customData.eventName = 'ApiRepoCreateMailFailure';
    insights?.trackException({ exception: mailError, properties: customData });
  }
  // send to operations, too
  delete additionalMail.cc;
  const notifyMailAddress = operations.getRepositoriesNotificationMailAddress();
  const operationsMails = notifyMailAddress ? [notifyMailAddress] : [];
  if (operationsMails && operationsMails.length) {
    additionalMail.to = operationsMails;
    contentOptions.reason = `You are receiving this e-mail as the operations contact address(es) ${operationsMails.join(
      ', '
    )}. A repo has been created or classified.`;
    try {
      additionalMail.content = await RenderHtmlMail(
        config.typescript.appDirectory,
        emailTemplate,
        contentOptions
      );
    } catch (renderError) {
      console.dir(renderError);
      return;
    }
    try {
      await mailProvider.sendMail(additionalMail);
    } catch (ignoredError) {
      console.dir(ignoredError);
      return;
    }
  }
}
