//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import express from 'express';
import asyncHandler from 'express-async-handler';
const router = express.Router();

import { getProviders, IProviders, ReposAppRequest } from '../transitional';

import { requirePortalAdministrationPermission } from '../middleware/business/administration';
import { PostgresLinkProvider } from '../lib/linkProviders/postgres/postgresLinkProvider';
import { Operations, UnlinkPurpose } from '../business';
import { ICorporateLink } from '../business';
import { Organization } from '../business';
import { Account } from '../business';
import { ILinkProvider } from '../lib/linkProviders';

// - - - Middleware: require that the user isa portal administrator to continue
router.use(requirePortalAdministrationPermission);
// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

// These functions are not pretty.

enum OperationsAction {
  DestroyLink,
  MarkAsServiceAccount,
  UnmarkServiceAccount,
}

enum UserQueryByType {
  ByGitHubId,
  ByGitHubUsername,
  ByCorporateUsername,
}

interface IUserInformationQuery {
  queryByValue: string;
  queryByType: UserQueryByType;

  link?: ICorporateLink;
  noLinkButKnownThirdPartyId?: string;

  collaboratorRepositories?: string[];

  orgs?: Organization[];
  gitHubUserInfo?: any;
  renamedGitHubUserOutcome?: UserQueryOutcomeRenamedThirdPartyUsername;
  deletedGitHubUserOutcome?: any;
  realtimeGraph?: any;
  realtimeGraphError?: any;
  managerInfo?: any;
}

class UserQueryOutcomeRenamedThirdPartyUsername {
  public username: string;
  public message: string;

  constructor(
    newUsername: string,
    knownPreviousUsername?: string,
    differentMessage?: string
  ) {
    this.username = newUsername;
    if (!knownPreviousUsername) {
      this.message = `The username was renamed to ${newUsername}`;
    } else {
      this.message = `The user previously known as ${knownPreviousUsername} is now ${newUsername}`;
    }
    if (differentMessage) {
      this.message = differentMessage;
    }
  }
}

router.get('/', function (req: ReposAppRequest, res) {
  req.individualContext.webContext.render({
    view: 'organization/index',
    title: 'Organization Dashboard',
  });
});

async function queryByGitHubLogin(
  providers: IProviders,
  login: string
): Promise<IUserInformationQuery> {
  const { operations } = providers;
  const query: IUserInformationQuery = {
    queryByType: UserQueryByType.ByGitHubUsername,
    queryByValue: login,
  };
  let gitHubAccountInfo = null;
  try {
    gitHubAccountInfo = await operations.getAccountByUsername(login);
  } catch (error) {
    // They may have renamed their GitHub username, but the ID is the same as it was before...
    if (error && error.statusCode === 404) {
      const linkByOldName = (await getLinkByThirdPartyUsername(
        providers,
        login
      )) as ICorporateLink;
      if (linkByOldName && linkByOldName.thirdPartyId) {
        const anotherTryGitHubId = linkByOldName.thirdPartyId;
        query.link = linkByOldName;
        gitHubAccountInfo = await getGitHubAccountInformationById(
          operations,
          anotherTryGitHubId
        );
        query.gitHubUserInfo = gitHubAccountInfo;
        error = null;
      }
    }
    if (error) {
      throw error;
    }
  }
  if (!query.link && gitHubAccountInfo && gitHubAccountInfo.login) {
    try {
      query.link = await getLinkByThirdPartyUsername(
        providers,
        gitHubAccountInfo.login
      );
    } catch (queryByLoginAttempt) {
      query.noLinkButKnownThirdPartyId = gitHubAccountInfo.id;
      if (queryByLoginAttempt.status == 404 /* loose*/) {
        console.warn('Not linked');
      } else {
        console.dir(queryByLoginAttempt);
      }
    }
  }
  return loadInformation(providers, query);
}

function getLinkByThirdPartyUsername(
  providers: IProviders,
  login: string
): Promise<ICorporateLink> {
  const linkProvider = providers.linkProvider;
  return linkProvider.getByThirdPartyUsername(login);
}

async function queryByGitHubId(
  providers: IProviders,
  thirdPartyId: string
): Promise<IUserInformationQuery> {
  const { linkProvider } = providers;
  const link = await linkProvider.getByThirdPartyId(thirdPartyId);
  const query: IUserInformationQuery = {
    queryByType: UserQueryByType.ByGitHubId,
    queryByValue: thirdPartyId,
  };
  if (link) {
    query.link = link;
  } else {
    query.noLinkButKnownThirdPartyId = thirdPartyId;
  }
  return loadInformation(providers, query);
}

async function queryByCorporateUsername(
  providers: IProviders,
  upn: string
): Promise<IUserInformationQuery> {
  const linkProvider = providers.linkProvider;
  const links = await linkProvider.queryByCorporateUsername(upn);
  if (!links || links.length !== 1) {
    if (!links || links.length <= 0) {
      throw new Error(
        `No links were identified for the corporate username ${upn}`
      );
    } else {
      const ids = links.map((link) => link['id']);
      throw new Error(
        `Too many links (more than one) exist for the corporate username ${upn}. Individual Link IDs: ${ids.join(
          ', '
        )}`
      );
    }
  }
  const query: IUserInformationQuery = {
    queryByType: UserQueryByType.ByCorporateUsername,
    queryByValue: upn,
    link: links[0],
  };
  return loadInformation(providers, query);
}

async function loadInformation(
  providers: IProviders,
  query: IUserInformationQuery
): Promise<IUserInformationQuery> {
  // Input: query type and value; pre-queried and set single link, if present
  const { operations } = providers;
  const corporateAadId = query.link ? query.link.corporateId : null;
  if (corporateAadId) {
    try {
      const info = await operations.validateCorporateAccountCanLink(
        corporateAadId
      );
      query.realtimeGraph = info.graphEntry;
    } catch (graphError) {
      query.realtimeGraphError = graphError;
    }
    try {
      query.managerInfo = await operations.getCachedEmployeeManagementInformation(
        corporateAadId
      );
    } catch (managerError) {
      console.dir(managerError);
    }
  }

  // Get user account information from GitHub
  let thirdPartyId = query.link
    ? query.link.thirdPartyId
    : query.noLinkButKnownThirdPartyId;
  if (query.gitHubUserInfo && query.gitHubUserInfo.id) {
    // In the scenario that they have renamed their account, this may come up...
    thirdPartyId = query.gitHubUserInfo.id;
  }
  let thirdPartyUsername: string = null;
  let account: Account = null;
  try {
    if (thirdPartyId) {
      account = await getGitHubAccountInformationById(operations, thirdPartyId);
      query.gitHubUserInfo = account;
      const login = account.login;
      if (query.link && login !== query.link.thirdPartyUsername) {
        query.renamedGitHubUserOutcome = new UserQueryOutcomeRenamedThirdPartyUsername(
          login,
          query.link.thirdPartyUsername
        );
      }
      thirdPartyUsername = login;

      const { queryCache } = operations.providers;
      if (queryCache && queryCache.supportsRepositoryCollaborators) {
        const result = await queryCache.userCollaboratorRepositories(
          thirdPartyId
        );
        const collaboratorRepositories = [];
        for (const { repository } of result) {
          try {
            await repository.getDetails();
            collaboratorRepositories.push(repository.full_name);
          } catch (ignoreError) {
            console.dir(ignoreError);
          }
        }
        query.collaboratorRepositories = collaboratorRepositories;
      }
    }
  } catch (ignoreGetAccountError) {
    if (
      ignoreGetAccountError &&
      ignoreGetAccountError.status == /* loose compare */ '404'
    ) {
      thirdPartyUsername = query.link ? query.link.thirdPartyUsername : null;
      if (thirdPartyUsername) {
        let deletedAccountError = null;
        let moreInfo = null;
        try {
          moreInfo = await operations.getAccountByUsername(thirdPartyUsername);
        } catch (deletedAccountCatch) {
          if (
            deletedAccountCatch &&
            deletedAccountCatch.status == /* loose compare */ '404'
          ) {
            deletedAccountError = deletedAccountCatch;
            query.deletedGitHubUserOutcome = `The GitHub account '${thirdPartyUsername}' (ID ${thirdPartyId}) has been deleted`;
          } else {
            throw deletedAccountError;
          }
        }
        query.gitHubUserInfo = moreInfo;
        if (moreInfo && moreInfo.id != /* loose compare */ thirdPartyId) {
          const newId = moreInfo.id;
          query.renamedGitHubUserOutcome = new UserQueryOutcomeRenamedThirdPartyUsername(
            thirdPartyUsername,
            thirdPartyUsername,
            `The original GitHub username this user linked with, ${thirdPartyUsername}, exists. However, the user ID is different now. It was ${thirdPartyId} and now the ID is ${newId}. They most likely deleted their old account or have two-factor problems.`
          );
        }
      }
    } else {
      console.warn(ignoreGetAccountError);
    }
  }

  // Learn about all the org memberships for the username
  if (thirdPartyUsername && account) {
    const loginMemberships = await account.getOperationalOrganizationMemberships();
    query.orgs = loginMemberships;
  }

  return query;
}

async function getGitHubAccountInformationById(
  operations: Operations,
  id: string
): Promise<Account> {
  const account = operations.getAccount(id);
  await account.getDetails();
  return account;
}

router.get('/whois/id/:githubid', function (req: ReposAppRequest, res, next) {
  const thirdPartyId = req.params.githubid;
  const providers = getProviders(req);
  queryByGitHubId(providers, thirdPartyId)
    .then((query) => {
      req.individualContext.webContext.render({
        view: 'organization/whois/result',
        title: `Whois by GitHub ID: ${thirdPartyId}`,
        state: {
          info: query.gitHubUserInfo,
          realtimeGraph: query.realtimeGraph,
          postUrl: `/organization/whois/id/${thirdPartyId}`,
          // new-style
          query,
        },
      });
    })
    .catch(next);
});

enum IDValueType {
  ID,
  Username,
}

interface IIDValue {
  type: IDValueType;
  value: string;
}

router.get(
  '/whois/link/:linkid',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const linkId = req.params.linkid;
    const { linkProvider: lp } = getProviders(req);
    const linkProvider = lp as PostgresLinkProvider;
    const link = await linkProvider.getByPostgresLinkId(linkId);
    return req.individualContext.webContext.render({
      view: 'organization/whois/linkEditorPage',
      title: `Link ${linkId}`,
      state: {
        query: {
          link,
        },
      },
    });
  })
);

router.post(
  '/whois/link/:linkid',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const linkId = req.params.linkid;
    const isLinkDelete = req.body['delete-link'];
    req.body['isServiceAccount'] = req.body['isServiceAccount'] === 'yes';
    const keys = [
      'corporateId',
      'corporateUsername',
      'corporateDisplayName',
      'thirdPartyId',
      'thirdPartyUsername',
      'thirdPartyAvatar',
      'isServiceAccount',
      'serviceAccountMail',
    ];
    for (const key of keys) {
      if (!isLinkDelete && !req.body[key]) {
        return next(new Error(`Must provide a value for ${key}`));
      }
      break;
    }
    const { linkProvider: lp } = getProviders(req);
    const linkProvider = lp as PostgresLinkProvider;
    const link = await linkProvider.getByPostgresLinkId(linkId);
    const messages = [`Link ID ${linkId}`];
    let hadUpdates = false;
    for (const key of keys) {
      // loose comparisons
      if (!isLinkDelete && link[key] != req.body[key]) {
        messages.push(
          `${key}: value has been updated from "${link[key]}" to "${req.body[key]}"`
        );
        link[key] = req.body[key];
        hadUpdates = true;
      }
    }
    const renderOutput = function () {
      req.individualContext.webContext.render({
        view: 'organization/whois/linkUpdate',
        title: `Updating link ${linkId}`,
        state: {
          messages,
          linkId,
        },
      });
    };
    if (isLinkDelete) {
      messages.push(`Deleting link ${linkId}`);
      try {
        await linkProvider.deleteLink(link);
        messages.push('Link deleted OK');
      } catch (error) {
        messages.push(error.toString());
      }
      return renderOutput();
    }
    if (hadUpdates) {
      messages.push('Updating values');
      await linkProvider.updateLink(link);
      return renderOutput();
    } else {
      messages.push('No link values changed, it was not updated');
      return renderOutput();
    }
  })
);

router.post(
  '/whois/link/',
  asyncHandler(async function (req: ReposAppRequest, res, next) {
    const { operations } = getProviders(req);
    const allowAdministratorManualLinking =
      operations?.config?.features?.allowAdministratorManualLinking;
    if (!allowAdministratorManualLinking) {
      return next(new Error('The manual linking feature is not enabled'));
    }

    // set isServiceAccount to true only if it contains the value "yes", otherwise use false
    req.body['isServiceAccount'] = req.body['isServiceAccount'] === 'yes';

    // create link object with the values received from the request
    const link: ICorporateLink = {
      corporateId: req.body['corporateId'],
      corporateUsername: req.body['corporateUsername'],
      corporateDisplayName: req.body['corporateDisplayName'],
      thirdPartyId: req.body['thirdPartyId'],
      thirdPartyUsername: req.body['thirdPartyUsername'],
      thirdPartyAvatar: req.body['thirdPartyAvatar'],
      isServiceAccount: req.body['isServiceAccount'],
      serviceAccountMail: req.body['serviceAccountMail'],
      // these both values are currently not transferred, but required by the link object
      corporateMailAddress: '',
      corporateAlias: '',
    };

    const messages = [];
    // Add only the non empty strings to the message log
    for (const [key, value] of Object.entries(link)) {
      if (value) {
        messages.push(`${key}: value has been set to "${value}"`);
      }
    }

    const linkProvider = operations.providers
      .linkProvider as PostgresLinkProvider;

    // try to create link, if it fails it will directly throw into the users face
    const linkId = await linkProvider.createLink(link);
    // Add the created link id to the messages
    messages.push(`Link ID ${linkId}`);

    // render the output
    req.individualContext.webContext.render({
      view: 'organization/whois/linkUpdate',
      title: `Updating link ${linkId}`,
      state: {
        messages,
        linkId,
      },
    });
  })
);

router.post('/whois/id/:githubid', function (req: ReposAppRequest, res, next) {
  const thirdPartyId = req.params.githubid;
  const markAsServiceAccount = req.body['mark-as-service-account'];
  const unmarkServiceAccount = req.body['unmark-service-account'];
  const providers = getProviders(req);
  let action = OperationsAction.DestroyLink;
  if (markAsServiceAccount) {
    action = OperationsAction.MarkAsServiceAccount;
  } else if (unmarkServiceAccount) {
    action = OperationsAction.UnmarkServiceAccount;
  }
  const idValue: IIDValue = {
    type: IDValueType.ID,
    value: thirdPartyId,
  };
  destructiveLogic(providers, idValue, action, req, res, next)
    .then((state) => {
      if (state.independentView) {
        return;
      }
      req.individualContext.webContext.render({
        view: 'organization/whois/drop',
        title: `Dropped link by ID ${thirdPartyId}`,
        state,
      });
    })
    .catch((error) => {
      return next(error);
    });
});

router.get('/whois/aad/:upn', function (req: ReposAppRequest, res, next) {
  const upn = req.params.upn;
  const providers = getProviders(req);
  queryByCorporateUsername(providers, upn)
    .then((query) => {
      req.individualContext.webContext.render({
        view: 'organization/whois/result',
        title: `Whois by AAD UPN: ${upn}`,
        state: {
          upn,
          info: query.gitHubUserInfo,
          realtimeGraph: query.realtimeGraph,

          // new-style
          query,
        },
      });
    })
    .catch(next);
});

router.get(
  '/whois/github/:username',
  function (req: ReposAppRequest, res, next) {
    const login = req.params.username;
    const providers = getProviders(req);
    queryByGitHubLogin(providers, login)
      .then((query) => {
        req.individualContext.webContext.render({
          view: 'organization/whois/result',
          title: `Whois: ${login}`,
          state: {
            info: query.gitHubUserInfo,
            realtimeGraph: query.realtimeGraph,
            // new-style
            query,
          },
        });
      })
      .catch(next);
  }
);

router.post(
  '/whois/github/:username',
  function (req: ReposAppRequest, res, next) {
    const username = req.params.username;
    const markAsServiceAccount = req.body['mark-as-service-account'];
    const unmarkServiceAccount = req.body['unmark-service-account'];
    const providers = getProviders(req);
    let action = OperationsAction.DestroyLink;
    if (markAsServiceAccount) {
      action = OperationsAction.MarkAsServiceAccount;
    } else if (unmarkServiceAccount) {
      action = OperationsAction.UnmarkServiceAccount;
    }
    const identifier: IIDValue = {
      type: IDValueType.Username,
      value: username,
    };
    destructiveLogic(providers, identifier, action, req, res, next)
      .then((state) => {
        if (state.independentView) {
          return;
        }
        req.individualContext.webContext.render({
          view: 'organization/whois/drop',
          title: `Dropped ${username}`,
          state,
        });
      })
      .catch((error) => {
        return next(error);
      });
  }
);

async function destructiveLogic(
  providers: IProviders,
  identifier: IIDValue,
  action: OperationsAction,
  req,
  res,
  next
): Promise<any> {
  const { operations } = providers;
  let usernameInfo = null;
  let state = {
    results: null,
    entity: null,
    messages: [],
    independentView: false,
  };
  let thirdPartyUsername =
    identifier.type === IDValueType.Username ? identifier.value : null;
  let thirdPartyId =
    identifier.type === IDValueType.ID ? identifier.value : null;
  const dataAsTerminated = req.body.dataTerminated === 'yes';
  try {
    if (!thirdPartyUsername) {
      state.messages.push('Destruction operation not requested on a username');
    } else {
      usernameInfo = await operations.getAccountByUsername(thirdPartyUsername);
      if (thirdPartyId && usernameInfo.id !== thirdPartyId) {
        state.messages.push(
          `The retrieved ID for the username was ${usernameInfo.id} instead of the expected ${thirdPartyId}`
        );
      } else if (!thirdPartyId && usernameInfo.id) {
        thirdPartyId = usernameInfo.id;
      }
    }
  } catch (grabError) {
    state.messages.push(
      `Could not get GitHub account information by USERNAME ${thirdPartyUsername}: ` +
        grabError.toString()
    );
  }
  state.entity = usernameInfo;

  let idInfo = null;
  if (thirdPartyId) {
    try {
      idInfo = await getGitHubAccountInformationById(operations, thirdPartyId);
    } catch (idInfoError) {
      if (idInfoError.status === '404') {
        state.messages.push(
          `The GitHub account by ID does not exist or has been deleted: ${thirdPartyId}: ` +
            idInfoError.toString()
        );
      } else {
        state.messages.push(
          `Could not get GitHub account information by ID ${thirdPartyId}: ` +
            idInfoError.toString()
        );
      }
    }
  }

  let linkQuery = null;
  if (thirdPartyId) {
    try {
      linkQuery = await queryByGitHubId(providers, thirdPartyId);
    } catch (oops) {
      console.dir(oops);
      state.messages.push(
        `Could not find a corporate link by their GitHub user ID of ${thirdPartyId}`
      );
      if (usernameInfo && usernameInfo.login) {
        state.messages.push(
          `Will try next by their GitHub username: ${usernameInfo.login}`
        );
      }
      try {
        linkQuery = await queryByGitHubLogin(providers, thirdPartyUsername);
        state.messages.push(
          `Did find a link by their login on GitHub, ${thirdPartyUsername}. Will terminate this ID.`
        );
      } catch (linkByUsernameError) {
        state.messages.push(
          `Could not find a link by login, ${thirdPartyUsername}. Hmm.`
        );
      }
    }
  }

  // Service Account settings (not so destructive)
  if (
    action === OperationsAction.MarkAsServiceAccount ||
    action === OperationsAction.UnmarkServiceAccount
  ) {
    const linkProvider = operations.providers.linkProvider;
    state.independentView = true; // no rendering on return
    return await modifyServiceAccount(
      linkProvider,
      linkQuery.link,
      action === OperationsAction.MarkAsServiceAccount,
      req,
      res,
      next
    );
  }

  // Account termination
  if (linkQuery && linkQuery.link && !thirdPartyId) {
    thirdPartyId = linkQuery.link.thirdPartyId;
  }
  if (thirdPartyId) {
    const purpose = dataAsTerminated
      ? UnlinkPurpose.Termination
      : UnlinkPurpose.Operations;
    state.results = await operations.terminateLinkAndMemberships(thirdPartyId, {
      purpose,
    });
  } else {
    state.messages.push('Could not terminate the account, no link was found');
  }

  return state;
}

async function modifyServiceAccount(
  linkProvider: ILinkProvider,
  link: ICorporateLink,
  markAsServiceAccount,
  req,
  res,
  next
) {
  link.isServiceAccount = markAsServiceAccount ? true : false;
  try {
    await linkProvider.updateLink(link);
    return res.json(link);
  } catch (updateError) {
    return next(updateError);
  }
}

// ----------------------------------------------------------------------------

router.get('/bulkRepoDelete', (req: ReposAppRequest, res) => {
  req.individualContext.webContext.render({
    view: 'organization/bulkRepoDelete',
    title: 'Bulk repository delete',
  });
});

router.post(
  '/bulkRepoDelete',
  asyncHandler(async (req: ReposAppRequest, res, next) => {
    const { operations } = getProviders(req);
    let repositories = req.body.repositories;
    // TODO: FEATURE FLAG: add a feature flag whether this API is available.
    if (!repositories) {
      return next(new Error('No repositories provided'));
    }
    repositories = repositories.split('\n');
    const log = [];
    for (let repositoryName of repositories) {
      repositoryName = (repositoryName || '').trim();
      if (!repositoryName.length) {
        continue;
      }
      let githubcom = 'github.com';
      let ghi = repositoryName.indexOf(githubcom);
      if (ghi >= 0) {
        let name = repositoryName.substr(ghi + githubcom.length + 1);
        let divider = name.indexOf('/');
        if (divider <= 0) {
          continue;
        }
        let orgName = name.substr(0, divider);
        let repoName = name.substr(divider + 1);
        const repository = operations
          .getOrganization(orgName)
          .repository(repoName);
        try {
          await repository.delete();
          // let metaStatus = more && more.headers ? more.headers.status : null;
          log.push(`${name}: deleted`);
        } catch (deleteError) {
          log.push(`${name}: error: ${deleteError}`);
        }
      } else {
        log.push(
          `Skipping, does not appear to be a GitHub repo URL: ${repositoryName}`
        );
      }
    }
    return res.json(log);
  })
);

export default router;
