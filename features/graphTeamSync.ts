//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Team } from '../business/team';
import { IProviders } from '../transitional';

export class GraphTeamSync {
  #team: Team;
  #providers: IProviders;

  constructor(providers: IProviders, team: Team) {
    this.#providers = providers;
    this.#team = team;
  }

  // On the GitHub side, they define:
  // - The team must not have any sub-teams (must be a leaf node)
  // - Can have up to 5 directory groups assigned to the team
  // - Team Maintainers are not touched (TBD need to confirm)
  // - If you suspend sync, it removes all the members

  // Unknowns:
  // - What happens if the sole maint leaves... guess it just becomes an orphan team
  // - How to handle removed directory groups
  // - How much to store, log, show in the UI, or notify people about

  // Value-adds:
  // - When a user links for the first time, kick off a job to evaluate the directory groups they're in
}
