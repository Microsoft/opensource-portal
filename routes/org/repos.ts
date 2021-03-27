//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import _ from 'lodash';
import asyncHandler from 'express-async-handler';
import express from 'express';
import moment from 'moment';

import lowercaser from '../../middleware/lowercaser';

import {
  Collaborator,
  GitHubCollaboratorAffiliationQuery,
  ICorporateLink,
  ITemporaryCommandOutput,
  Organization,
  OrganizationMember,
  Repository,
  TeamPermission,
} from '../../business';

import {
  ReposAppRequest,
  IProviders,
  UserAlertType,
  CreateError,
  NoCacheNoBackground,
  getProviders,
  ErrorHelper,
} from '../../transitional';
import { RepositoryMetadataEntity } from '../../entities/repositoryMetadata/repositoryMetadata';
import {
  AddRepositoryPermissionsToRequest,
  getContextualRepositoryPermissions,
  IContextualRepositoryPermissions,
} from '../../middleware/github/repoPermissions';

import routeAdministrativeLock from './repoAdministrativeLock';
import NewRepositoryLockdownSystem from '../../features/newRepositoryLockdown';
import { IGraphEntry } from '../../lib/graphProvider';
import { IMail } from '../../lib/mailProvider';
import { IndividualContext } from '../../user';

const router = express.Router();

import RouteReposPager from '../reposPager';

interface ILocalRequest extends ReposAppRequest {
  repository?: Repository;
  repositoryMetadata?: RepositoryMetadataEntity;
  repoPermissions?: any;
}

interface IFindRepoCollaboratorsExcludingTeamsResult {
  collaborators: Collaborator[];
  memberCollaborators: Collaborator[];
  outsideCollaborators: Collaborator[];
}

interface ICalculateRepoPermissionsResult
  extends IFindRepoCollaboratorsExcludingTeamsResult {
  permissions: TeamPermission[];
}

const teamsFilterType = {
  systemTeamsExcluded: 'systemTeamsExcluded',
  systemTeamsOnly: 'systemTeamsOnly',
};

router.use(function (req: ReposAppRequest, res, next) {
  req.individualContext.webContext.pushBreadcrumb('Repositories');
  req.reposContext = {
    section: 'repos',
    organization: req.organization,
    pivotDirectlyToOtherOrg: '/repos/', // hack
  };
  req.reposPagerMode = 'org';
  next();
});

router.get(
  '/',
  lowercaser(['sort', 'language', 'type', 'tt']),
  RouteReposPager
);

function sliceCollaboratorsForView(collaborators) {
  // Slices to the highest permission level for a collaborator
  const collabView = {
    readers: [],
    writers: [],
    administrators: [],
  };
  collaborators.forEach((collab) => {
    const permission = collab.permissions;
    const destination = permission.admin
      ? collabView.administrators
      : permission.push
      ? collabView.writers
      : permission.pull
      ? collabView.readers
      : null;
    if (destination) {
      destination.push(collab);
    }
  });
  return collabView;
}

function slicePermissionsForView(permissions) {
  const perms = {};
  permissions.forEach((permission) => {
    const level = permission.permission;
    if (!level) {
      throw new Error(
        'Invalid operation: no permission associated with the permission entity'
      );
    }
    if (!perms[level]) {
      perms[level] = [];
    }
    perms[level].push(permission);
  });
  return perms;
}

async function calculateRepoPermissions(
  organization: Organization,
  repository: Repository
): Promise<ICalculateRepoPermissionsResult> {
  const teamPermissions = await repository.getTeamPermissions();
  const owners = await organization.getOwners();
  const {
    collaborators,
    outsideCollaborators,
    memberCollaborators,
  } = await findRepoCollaboratorsExcludingOwners(repository, owners);
  for (let teamPermission of teamPermissions) {
    try {
      teamPermission.resolveTeamMembers();
    } catch (ignoredError) {
      /* ignored */
    }
  }
  return {
    permissions: teamPermissions,
    collaborators,
    outsideCollaborators,
    memberCollaborators,
  };
}

export async function findRepoCollaboratorsExcludingOwners(
  repository: Repository,
  owners: OrganizationMember[]
): Promise<IFindRepoCollaboratorsExcludingTeamsResult> {
  const ownersMap = new Map<number, OrganizationMember>();
  for (let i = 0; i < owners.length; i++) {
    ownersMap.set(Number(owners[i].id), owners[i]);
  }
  const collaborators = await repository.getCollaborators({
    affiliation: GitHubCollaboratorAffiliationQuery.Direct,
  });
  const outsideCollaborators = await repository.getCollaborators({
    affiliation: GitHubCollaboratorAffiliationQuery.Outside,
  });
  function filterOutOwners(collaborator: Collaborator) {
    const id = Number(collaborator.id);
    return !ownersMap.has(id);
  }
  const collaboratorsWithoutOwners = _.filter(collaborators, filterOutOwners);
  const outsideCollaboratorIds = new Set<number>(
    outsideCollaborators.map((oc) => Number(oc.id))
  );
  return {
    collaborators: collaboratorsWithoutOwners,
    outsideCollaborators,
    memberCollaborators: collaboratorsWithoutOwners.filter(
      (collab) => outsideCollaboratorIds.has(Number(collab.id)) === false
    ),
  };
}

router.use(
  '/:repoName',
  asyncHandler(async function (req: ILocalRequest, res, next) {
    const repoName = req.params.repoName;
    const organization = req.organization;
    const repository = organization.repository(repoName);
    await repository.getDetails();
    req.repository = repository;
    req.repositoryMetadata = await repository.getRepositoryMetadata();
    return next();
  })
);

router.use('/:repoName/administrativeLock', routeAdministrativeLock);

router.use(
  '/:repoName/delete',
  asyncHandler(async function (req: ILocalRequest, res, next) {
    const individualContext = req.individualContext;
    const repository = req.repository;
    const organization = req.organization;
    if (!organization.isNewRepositoryLockdownSystemEnabled) {
      return next(
        new Error('This endpoint is not available as configured in this app.')
      );
    }
    const daysAfterCreateToAllowSelfDelete = 21; // could be a config setting if anyone cares
    try {
      const metadata = await repository.getRepositoryMetadata();
      await NewRepositoryLockdownSystem.ValidateUserCanSelfDeleteRepository(
        repository,
        metadata,
        individualContext,
        daysAfterCreateToAllowSelfDelete
      );
    } catch (noExistingMetadata) {
      if (noExistingMetadata.status === 404) {
        throw new Error(
          'This repository does not have any metadata available regarding who can setup it up. No further actions available.'
        );
      }
      throw noExistingMetadata;
    }
    return next();
  })
);

router.get(
  '/:repoName/delete',
  asyncHandler(async function (req: ILocalRequest, res, next) {
    return req.individualContext.webContext.render({
      title: 'Delete the repo you created',
      view: 'repos/delete',
      state: {
        repo: req.repository,
      },
    });
  })
);

router.post(
  '/:repoName/delete',
  asyncHandler(async function (req: ILocalRequest, res, next) {
    // NOTE: this code is also duplicated for now in the client/internal/* folder
    // CONSIDER: de-duplicate
    const { operations, repositoryMetadataProvider } = getProviders(req);
    const { organization, repository } = req;
    const lockdownSystem = new NewRepositoryLockdownSystem({
      operations,
      organization,
      repository,
      repositoryMetadataProvider,
    });
    await lockdownSystem.deleteLockedRepository(
      false /* delete for any reason */,
      true /* deleted by the original user instead of ops */
    );
    req.individualContext.webContext.saveUserAlert(
      `You deleted your repo, ${repository.full_name}.`,
      'Repo deleted',
      UserAlertType.Success
    );
    return res.redirect(organization.baseUrl);
  })
);

export interface IRenameOutput {
  message: string;
  output: ITemporaryCommandOutput[];
}

router.post(
  '/:repoName/defaultBranch',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: ILocalRequest, res, next) {
    try {
      const targetBranchName = req.body.targetBranchName || 'main';
      const providers = getProviders(req);
      const activeContext = (req.individualContext ||
        req.apiContext) as IndividualContext;
      const repoPermissions = getContextualRepositoryPermissions(req);
      const repository = req.repository as Repository;
      const outcome = await renameRepositoryDefaultBranchEndToEnd(
        providers,
        activeContext,
        repoPermissions,
        repository,
        targetBranchName,
        false
      );
      req.individualContext.webContext.render({
        view: 'repos/repoBranchRenamed',
        title: outcome.message,
        state: {
          output: outcome.output,
          repository,
        },
      });
    } catch (error) {
      return next(error);
    }
  })
);

export async function renameRepositoryDefaultBranchEndToEnd(
  providers: IProviders,
  activeContext: IndividualContext,
  repoPermissions: IContextualRepositoryPermissions,
  repository: Repository,
  targetBranchName: string,
  waitForRefresh: boolean
): Promise<IRenameOutput> {
  const corporateUsername = activeContext.corporateIdentity.username;
  if (!corporateUsername) {
    throw CreateError.InvalidParameters('no corporate username in the session');
  }
  if (!targetBranchName) {
    throw CreateError.InvalidParameters('invalid target branch name');
  }
  if (!repoPermissions) {
    throw CreateError.InvalidParameters('no repo permissions');
  }
  if (!repoPermissions.allowAdministration) {
    throw CreateError.NotAuthorized(
      'You do not have administrative permission on this repository'
    );
  }
  await repository.getDetails();
  function finishUp(): Promise<void> {
    return new Promise((resolve) => {
      triggerRenameNotification(
        providers,
        repository,
        corporateUsername,
        targetBranchName,
        output
      )
        .then((ok) => {
          /* ignore */
        })
        .catch((error) => {
          console.error(`Notify rename trigger: ${error}`);
        });
      repository
        .getDetails(NoCacheNoBackground)
        .then((ok) => {
          return resolve();
        })
        .catch((error) => {
          console.error(`Background refresh error: ${error}`);
          return resolve();
        });
    });
  }
  const output = await repository.renameDefaultBranch(targetBranchName);
  if (waitForRefresh) {
    await finishUp();
  } else {
    process.nextTick(() => {
      finishUp()
        .then((ok) => {
          /* ignore */
        })
        .catch((error) => {
          /* ignore */
        });
    });
  }
  return {
    message: `Branch renamed to ${targetBranchName} for ${repository.name}`,
    output,
  };
}

router.post(
  '/:repoName',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: ILocalRequest, res, next) {
    const repoPermissions = req.repoPermissions;
    if (!repoPermissions.allowAdministration) {
      return next(
        new Error(
          'You do not have administrative permission on this repository'
        )
      );
    }
    // only supporting the 'take public' operation now
    const takePublic = req.body['make-repo-public'];
    if (!takePublic) {
      return next(new Error('Unsupported operation'));
    }
    const repository = req.repository as Repository;
    await repository.editPublicPrivate({ private: false });
    req.individualContext.webContext.saveUserAlert(
      `${repository.full_name} is now public.`,
      'Repository publish',
      UserAlertType.Success
    );
    await repository.getDetails(NoCacheNoBackground);
    return res.redirect(
      `/${repository.organization.name}/repos/${repository.name}?published`
    );
  })
);

router.get(
  '/:repoName',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: ILocalRequest, res, next) {
    const { linkProvider, config, graphProvider } = getProviders(req);
    const repoPermissions = req.repoPermissions;
    const referer = req.headers.referer as string;
    const fromReposPage =
      referer && (referer.endsWith('repos') || referer.endsWith('repos/'));
    const organization = req.organization;
    const repository = req.repository;
    const repositoryMetadataEntity = req.repositoryMetadata;
    const organizationSupportsUpdatesApp = await organization.supportsUpdatesApp();
    let repo = null;
    try {
      repo = decorateRepoForView(await repository.getDetails());
    } catch (repoDetailsError) {
      console.dir(repoDetailsError);
      throw repoDetailsError;
    }
    // const { permissions, collaborators, outsideCollaborators } = await calculateRepoPermissions(organization, repository);
    // const systemTeams = combineAllTeams(organization.specialRepositoryPermissionTeams);
    // const teamBasedPermissions = consolidateTeamPermissions(permissions, systemTeams);
    const title = `${repository.name} - Repository`;
    const details = await repository.organization.getDetails();
    organization.id = details.id;
    const activeUserCorporateId = req.individualContext.corporateIdentity.id;
    let createdUserLink: ICorporateLink = null;
    const createdByThirdPartyId =
      repositoryMetadataEntity && repositoryMetadataEntity.createdByThirdPartyId
        ? repositoryMetadataEntity.createdByThirdPartyId
        : null;
    if (createdByThirdPartyId) {
      try {
        createdUserLink = await linkProvider.getByThirdPartyId(
          createdByThirdPartyId
        );
      } catch (linkError) {
        if (!ErrorHelper.IsNotFound(linkError)) {
          console.dir(linkError);
        }
      }
    }
    const createdByCorporateId =
      repositoryMetadataEntity && repositoryMetadataEntity.createdByCorporateId
        ? repositoryMetadataEntity.createdByCorporateId
        : null;
    if (!createdUserLink && createdByCorporateId) {
      try {
        const results = await linkProvider.queryByCorporateId(
          createdByCorporateId
        );
        if (results && results.length === 1) {
          createdUserLink = results[0];
        }
      } catch (linkError) {
        console.dir(linkError);
      }
    }
    const createdByThirdPartyUsername =
      repositoryMetadataEntity &&
      repositoryMetadataEntity.createdByThirdPartyUsername
        ? repositoryMetadataEntity.createdByThirdPartyUsername
        : null;
    if (!createdUserLink && createdByThirdPartyUsername) {
      try {
        createdUserLink = await linkProvider.getByThirdPartyUsername(
          createdByThirdPartyUsername
        );
      } catch (linkError) {
        console.dir(linkError);
      }
    }
    let currentManagementChain: IGraphEntry[] = null;
    try {
      if (createdUserLink && createdUserLink.corporateId) {
        currentManagementChain = (
          await graphProvider.getManagementChain(createdUserLink.corporateId)
        ).reverse();
      }
    } catch (ignoreError) {
      console.dir(ignoreError);
    }
    req.individualContext.webContext.render({
      view: 'repos/repo',
      title,
      state: {
        activeUserCorporateId,
        createdUserLink,
        organization,
        reposSubView: 'default',
        repoPermissions,
        entity: repository.getEntity(),
        currentManagementChain,
        repo,
        repository,
        // permissions: slicePermissionsForView(filterSystemTeams(teamsFilterType.systemTeamsExcluded, systemTeams, permissions)),
        // systemPermissions: slicePermissionsForView(filterSystemTeams(teamsFilterType.systemTeamsOnly, systemTeams, permissions)),
        // collaborators: sliceCollaboratorsForView(collaborators),
        // collaboratorsArray: collaborators,
        // outsideCollaboratorsSlice: sliceCollaboratorsForView(outsideCollaborators),
        // outsideCollaborators: outsideCollaborators,
        // reposDataAgeInformation: ageInformation ? ageInformation : undefined,
        // teamBasedPermissions,
        fromReposPage,
        organizationSupportsUpdatesApp,
        repositoryMetadataEntity,
      },
    });
  })
);

router.get(
  '/:repoName/defaultBranch',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: ILocalRequest, res, next) {
    const referer = req.headers.referer as string;
    const fromReposPage =
      referer && (referer.endsWith('repos') || referer.endsWith('repos/'));
    const organization = req.organization;
    const repoPermissions = req.repoPermissions;
    const repository = req.repository;
    const repositoryMetadataEntity = req.repositoryMetadata;
    await repository.getDetails();
    const title = `${repository.name} - Default Branch Name`;
    const details = await repository.organization.getDetails();
    const organizationSupportsUpdatesApp = await organization.supportsUpdatesApp();
    organization.id = details.id;
    req.individualContext.webContext.render({
      view: 'repos/defaultBranch',
      title,
      state: {
        organization,
        organizationSupportsUpdatesApp,
        repo: decorateRepoForView(repository),
        reposSubView: 'defaultBranch',
        repository,
        fromReposPage,
        repoPermissions,
        repositoryMetadataEntity,
      },
    });
  })
);

export interface IRepositoryPermissionsView {}

export async function calculateGroupedPermissionsViewForRepository(
  repository: Repository
): Promise<any> {
  const organization = repository.organization;
  const {
    permissions, // TeamPermission[]
    collaborators, // Collaborator[]
    outsideCollaborators, // Collaborator[]
  } = await calculateRepoPermissions(organization, repository);
  const systemTeams = combineAllTeams(
    organization.specialRepositoryPermissionTeams
  ); // number[]
  const teamBasedPermissions = consolidateTeamPermissions(
    permissions,
    systemTeams
  ); // busted?
  const groupedPermissions = slicePermissionsForView(
    filterSystemTeams(
      teamsFilterType.systemTeamsExcluded,
      systemTeams,
      permissions
    )
  );
  /*
    admin: TeamPermission[],
  */
  const systemPermissions = slicePermissionsForView(
    filterSystemTeams(teamsFilterType.systemTeamsOnly, systemTeams, permissions)
  );
  /*
    admin: TeamPermission[],
  */
  const groupedCollaborators = sliceCollaboratorsForView(collaborators);
  /*
  administrators: [],
  readers: [],
  writers: [],
  */
  const groupedOutsideCollaborators = sliceCollaboratorsForView(
    outsideCollaborators
  );
  /*
  administrators: [],
  readers: [],
  writers: [],
  */

  // const teamSets: aggregateTeamsToSets(aggregate.teams),
  // repoPermissions,
  const view = {
    teamBasedPermissions,
    systemTeams,
    groupedPermissions,
    systemPermissions,
    groupedCollaborators,
    groupedOutsideCollaborators,
  };
  return view;
}

router.get(
  '/:repoName/permissions',
  asyncHandler(AddRepositoryPermissionsToRequest),
  asyncHandler(async function (req: ILocalRequest, res, next) {
    const referer = req.headers.referer as string;
    const fromReposPage =
      referer && (referer.endsWith('repos') || referer.endsWith('repos/'));
    const organization = req.organization;
    const repoPermissions = req.repoPermissions;
    const repository = req.repository;
    const repositoryMetadataEntity = req.repositoryMetadata;
    const uc = req.individualContext.aggregations;
    const aggregate = await uc.getAggregatedOverview();
    await repository.getDetails();
    const {
      permissions,
      collaborators,
      outsideCollaborators,
    } = await calculateRepoPermissions(organization, repository);
    const systemTeams = combineAllTeams(
      organization.specialRepositoryPermissionTeams
    );
    const teamBasedPermissions = consolidateTeamPermissions(
      permissions,
      systemTeams
    );
    const title = `${repository.name} - Repository`;
    const details = await repository.organization.getDetails();
    organization.id = details.id;
    req.individualContext.webContext.render({
      view: 'repos/permissions',
      title,
      state: {
        organization,
        repo: decorateRepoForView(repository),
        reposSubView: 'permissions',
        repository,
        permissions: slicePermissionsForView(
          filterSystemTeams(
            teamsFilterType.systemTeamsExcluded,
            systemTeams,
            permissions
          )
        ),
        systemPermissions: slicePermissionsForView(
          filterSystemTeams(
            teamsFilterType.systemTeamsOnly,
            systemTeams,
            permissions
          )
        ),
        collaborators: sliceCollaboratorsForView(collaborators),
        collaboratorsArray: collaborators,
        outsideCollaboratorsSlice: sliceCollaboratorsForView(
          outsideCollaborators
        ),
        outsideCollaborators: outsideCollaborators,
        // reposDataAgeInformation: ageInformation ? ageInformation : undefined,
        fromReposPage,
        teamSets: aggregateTeamsToSets(aggregate.teams),
        repoPermissions,
        teamBasedPermissions,
        repositoryMetadataEntity,
      },
    });
  })
);

router.get(
  '/:repoName/history',
  asyncHandler(async function (req: ILocalRequest, res, next) {
    const { auditLogRecordProvider } = getProviders(req);
    const referer = req.headers.referer as string;
    const fromReposPage =
      referer && (referer.endsWith('repos') || referer.endsWith('repos/'));
    const organization = req.organization;
    const repository = req.repository;
    const repositoryMetadataEntity = req.repositoryMetadata;
    await repository.getDetails();
    const history = await auditLogRecordProvider.queryAuditLogForRepositoryOperations(
      repository.id.toString()
    );
    const title = `${repository.name} - History`;
    req.individualContext.webContext.render({
      view: 'repos/history',
      title,
      state: {
        organization,
        repo: decorateRepoForView(repository),
        reposSubView: 'history',
        repository,
        fromReposPage,
        repositoryMetadataEntity,
        history,
      },
    });
  })
);

function consolidateTeamPermissions(permissions, systemTeams) {
  const systemTeamsSet = new Set(systemTeams);
  const filtered = {
    // id -> [] array of teams
    admin: new Map(),
    push: new Map(),
    pull: new Map(),
  };
  for (let i = 0; i < permissions.length; i++) {
    const teamPermission = permissions[i];
    const permission = teamPermission.permission;
    const members = teamPermission.members;
    const team = teamPermission.team;
    const isSystemTeam = systemTeamsSet.has(team.id);
    if (members && !isSystemTeam /* skip system teams */) {
      for (let j = 0; j < members.length; j++) {
        const member = members[j];
        const map = filtered[permission];
        if (map) {
          let entry = map.get(member.id);
          if (!entry) {
            entry = {
              user: member,
              teams: [],
            };
            map.set(member.id, entry);
          }
          entry.teams.push(team);
        }
      }
    }
  }
  const expanded = {
    readers: Array.from(filtered.pull.values()),
    writers: Array.from(filtered.push.values()),
    administrators: Array.from(filtered.admin.values()),
  };
  return expanded.readers.length === 0 &&
    expanded.writers.length === 0 &&
    expanded.administrators.length === 0
    ? null
    : expanded;
}

function combineAllTeams(systemTeams) {
  const allTypes = Object.getOwnPropertyNames(systemTeams);
  const set = new Set();
  allTypes.forEach((type) => {
    const values = systemTeams[type];
    if (Array.isArray(values)) {
      for (let i = 0; i < values.length; i++) {
        set.add(values[i]);
      }
    }
  });
  return Array.from(set);
}

function filterSystemTeams(filterType, systemTeams, teams) {
  if (
    filterType !== teamsFilterType.systemTeamsExcluded &&
    filterType !== teamsFilterType.systemTeamsOnly
  ) {
    throw new Error(
      'Invalid, unsupported teamsFilterType value for filterType'
    );
  }
  const systemSet = new Set(systemTeams);
  return _.filter(teams, (permission) => {
    const team = permission.team;
    const isSystem = systemSet.has(team.id);
    return filterType === teamsFilterType.systemTeamsOnly
      ? isSystem
      : !isSystem;
  });
}

function decorateRepoForView(repo) {
  // This should just be a view service of its own at some point
  fromNow(repo, ['created_at', 'updated_at', 'pushed_at']);
  return repo;
}

function sanitizeReviewer(entry) {
  if (entry && entry.displayName) {
    let name = entry.displayName as string;
    const projectNameEnd = name.indexOf(']\\');
    if (projectNameEnd >= 0) {
      name = name.substr(projectNameEnd + 2);
    }
    const idIndex = name.indexOf(' <');
    if (idIndex >= 0) {
      name = name.substr(0, idIndex);
    }
    entry.displayName = name;
  }
  return entry;
}

function fromNow(object, property) {
  if (Array.isArray(property)) {
    property.forEach((prop) => {
      fromNow(object, prop);
    });
    return;
  }
  if (!object.moment) {
    object.moment = {};
  }
  let value = object[property];
  if (value) {
    object.moment[property] = moment(value).fromNow();
    return object.moment[property];
  }
}

function aggregateTeamsToSets(teams) {
  const sets = {
    maintained: teamsToSet(teams.maintainer),
    member: teamsToSet(teams.member),
  };
  return sets;
}

function teamsToSet(teams) {
  const set = new Set();
  if (teams) {
    teams.forEach((team) => {
      set.add(team.id);
    });
  }
  return set;
}

// function requireAdministration(req, res, next) {
//   const repoPermissions = req.repoPermissions;
//   if (!repoPermissions) {
//     return next(new Error('Not configured for repo permissions'));
//   }
//   if (repoPermissions.allowAdministration === true) {
//     return next();
//   }
//   return next(new Error('You are not authorized to administer this repository.'));
// }

async function triggerRenameNotification(
  providers: IProviders,
  repository: Repository,
  corporateUsername: string,
  targetBranchName: string,
  output: ITemporaryCommandOutput[]
): Promise<void> {
  const {
    config,
    insights,
    operations,
    mailAddressProvider,
    viewServices,
  } = providers;
  insights.trackMetric({ name: 'RenameDefaultBranchs', value: 1 });
  insights.trackEvent({
    name: 'RenameDefaultBranch',
    properties: {
      orgName: repository.organization.name,
      repoName: repository.name,
      targetBranchName,
    },
  });
  const mailAddress = await mailAddressProvider.getAddressFromUpn(
    corporateUsername
  );
  const emailTemplate = 'repoDefaultBranchRenamed';
  const mail: IMail = {
    to: [mailAddress],
    cc: [operations.getInfrastructureNotificationsMail()],
    subject: `${repository.organization.name}/${repository.name} default branch is now ${targetBranchName}`,
    content: undefined,
  };
  const contentOptions = {
    reason: `You are receiving this e-mail as a transaction record from your action to rename the default branch of this repository you administer.`,
    headline: `${targetBranchName} branch`,
    notification: 'information',
    app: config.brand?.companyName
      ? `${config.brand.companyName} GitHub`
      : 'GitHub',
    output,
    repository,
    organization: repository.organization,
    viewServices,
  };
  try {
    mail.content = await operations.emailRender(emailTemplate, contentOptions);
    await operations.sendMail(mail);
  } catch (mailError) {
    console.warn(mailError);
    insights.trackException({
      exception: mailError,
      properties: {
        repositoryName: repository.full_name,
        organizationName: repository.organization.name,
        eventName: 'SendRenameDefaultBranchMail',
      },
    });
  }
}

export default router;
