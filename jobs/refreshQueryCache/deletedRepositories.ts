import app, { IReposJob, IReposJobResult } from '../../app';
import { Organization } from '../../business/organization';
import { RepositoryCollaboratorCacheEntity } from '../../entities/repositoryCollaboratorCache/repositoryCollaboratorCache';
import { RepositoryTeamCacheEntity } from '../../entities/repositoryTeamCache/repositoryTeamCache';
import { ErrorHelper, IProviders } from '../../transitional';
import { sleep } from '../../utils';

const killBitHours = 8;

app.runJob(byUserJob, {
  defaultDebugOutput: 'qcuser',
  timeoutMinutes: 60 * killBitHours,
  insightsPrefix: 'JobRefreshUserQC',
});

const successDelayMilliseconds = 120;

const realRepositoryIds = new Set<number>();
const knownDeletedRepositoryIds = new Set<number>();

async function doesRepositoryExist(
  i: number,
  organization: Organization,
  repositoryId: number,
  knownRepositoryName: string
) {
  try {
    if (realRepositoryIds.has(repositoryId)) {
      return true;
    } else if (knownDeletedRepositoryIds.has(repositoryId)) {
      return false;
    }
    await organization.getRepositoryById(repositoryId);
    console.log(
      `${i}: repository ${knownRepositoryName} with ID ${repositoryId} in org ${organization.name} exists`
    );
    realRepositoryIds.add(repositoryId);
    await sleep(successDelayMilliseconds); // sleep a little if it is not a deleted repo
    return true;
  } catch (repositoryError) {
    if (ErrorHelper.IsNotFound(repositoryError)) {
      console.log(
        `${i}: repository deleted: ${knownRepositoryName} with ID ${repositoryId}`
      );
      knownDeletedRepositoryIds.add(repositoryId);
      return false;
    } else {
      console.log(repositoryError);
      throw repositoryError;
    }
  }
}

async function processDeletedRepositories(
  providers: IProviders
): Promise<void> {
  const queryCache = providers.queryCache;
  const repositoryTeamCacheProvider = providers.repositoryTeamCacheProvider;
  const checkingAllRepos = true;
  if (checkingAllRepos) {
    let deleted = 0;
    let reposCount = 0;
    try {
      let allRepositories = await queryCache.allRepositories();
      reposCount = allRepositories.length;
      console.log(`Incoming # of repositories cached: ${reposCount}`);
      for (let i = 0; i < allRepositories.length; i++) {
        const repositoryEntry = allRepositories[i];
        const organization = repositoryEntry.repository.organization;
        const organizationId = organization.id;
        const repositoryId = Number(repositoryEntry.repository.id);
        try {
          const existance = await doesRepositoryExist(
            i,
            organization,
            repositoryId,
            repositoryEntry.cacheEntity.repositoryName
          );
          if (existance === true) {
            console.log(
              `${i}: \t\t\trepository ${repositoryEntry.cacheEntity.repositoryName} with ID ${repositoryId} in org ${organization.name} exists`
            );
          } else if (existance === false) {
            console.log(
              `${i}: repository deleted: ${repositoryEntry.cacheEntity.repositoryName} with ID ${repositoryId}, will cleanup`
            );
            try {
              await queryCache.removeRepository(
                String(organizationId),
                String(repositoryId)
              );
              ++deleted;
            } catch (cleanupError) {
              console.log(
                `cleanupError for repository ID ${repositoryId}: ${cleanupError}`
              );
            }
          }
        } catch (error) {
          console.log(error);
        }
      }
    } catch (error) {
      console.dir(error);
      console.log(error);
    }

    console.log(`Incoming # of repositories cached was: ${reposCount}`);
    console.log(`Deleted repositories: ${deleted}`);
  }

  // Team permissions
  const checkingTeamPermissions = true;
  if (checkingTeamPermissions) {
    const allTeamPermissions = await repositoryTeamCacheProvider.queryAllTeams();
    const discoveredRepositoryIds = new Set<number>();
    const repoToTeamPermissions = new Map<
      number,
      RepositoryTeamCacheEntity[]
    >();
    allTeamPermissions.map((tp) => {
      const id = Number(tp.repositoryId);
      discoveredRepositoryIds.add(id);
      let entry = repoToTeamPermissions.get(id);
      if (!entry) {
        entry = [];
        repoToTeamPermissions.set(id, entry);
      }
      entry.push(tp);
    });
    let removedTeamPermissionRepositories = 0;
    const repoIds = Array.from(discoveredRepositoryIds.values()).sort();
    console.log(
      `Team permissions set for ${repoIds.length} repositories across ${allTeamPermissions.length} permission entries`
    );
    for (let i = 0; i < repoIds.length; i++) {
      try {
        const repositoryId = repoIds[i];
        if (realRepositoryIds.has(repositoryId)) {
          continue;
        }
        let deleteTeamPermission = knownDeletedRepositoryIds.has(repositoryId);
        if (!deleteTeamPermission) {
          const entries = repoToTeamPermissions.get(repositoryId);
          if (
            !providers.operations.isOrganizationManagedById(
              Number(entries[0].organizationId)
            )
          ) {
            console.log(`not managed here!`);
          }
          const org = providers.operations.getOrganizationById(
            Number(entries[0].organizationId)
          );
          const repoExists = await doesRepositoryExist(
            i,
            org,
            repositoryId,
            entries[0].repositoryName
          );
          if (!repoExists) {
            deleteTeamPermission = true;
          }
        }
        if (deleteTeamPermission) {
          await repositoryTeamCacheProvider.deleteByRepositoryId(
            String(repositoryId)
          );
          ++removedTeamPermissionRepositories;
        }
      } catch (error) {
        console.dir(error);
        console.log(error);
      }
    }
    console.log(
      `removed team permission repos: ${removedTeamPermissionRepositories}`
    );
  }

  // collaborator permissions
  const repositoryCollaboratorCacheProvider =
    providers.repositoryCollaboratorCacheProvider;
  const allCollaborators = await repositoryCollaboratorCacheProvider.queryAllCollaborators();
  const collaboratorRepositoryIds = new Set<number>();
  const collaboratorPermissionsMap = new Map<
    number,
    RepositoryCollaboratorCacheEntity[]
  >();
  allCollaborators.map((rcce) => {
    const id = Number(rcce.repositoryId);
    collaboratorRepositoryIds.add(id);
    let entry = collaboratorPermissionsMap.get(id);
    if (!entry) {
      entry = [];
      collaboratorPermissionsMap.set(id, entry);
    }
    entry.push(rcce);
  });
  let removedCollaboratorRepositories = 0;
  const collaboratorRepoIds = Array.from(
    collaboratorRepositoryIds.values()
  ).sort();
  console.log(
    `Repository collaborators across ${collaboratorRepoIds.length} repositories across ${allCollaborators.length} collaborator permission entries`
  );
  for (let i = 0; i < collaboratorRepoIds.length; i++) {
    try {
      const repositoryId = collaboratorRepoIds[i];
      if (realRepositoryIds.has(repositoryId)) {
        continue;
      }
      let deleteCollaboratorPermission = knownDeletedRepositoryIds.has(
        repositoryId
      );
      if (!deleteCollaboratorPermission) {
        const entries = collaboratorPermissionsMap.get(repositoryId);
        if (
          !providers.operations.isOrganizationManagedById(
            Number(entries[0].organizationId)
          )
        ) {
          console.log(`not managed here!`);
        }
        const org = providers.operations.getOrganizationById(
          Number(entries[0].organizationId)
        );
        const repoExists = await doesRepositoryExist(
          i,
          org,
          repositoryId,
          entries[0].repositoryId
        );
        if (!repoExists) {
          deleteCollaboratorPermission = true;
        }
      }
      if (deleteCollaboratorPermission) {
        await repositoryCollaboratorCacheProvider.deleteByRepositoryId(
          String(repositoryId)
        );
        ++removedCollaboratorRepositories;
      }
    } catch (error) {
      console.dir(error);
      console.log(error);
    }
  }
  console.log(`removed collaborator repos: ${removedCollaboratorRepositories}`);
}

export default async function byUserJob({
  providers,
  args,
}: IReposJob): Promise<IReposJobResult> {
  await processDeletedRepositories(providers);

  return {};
}
