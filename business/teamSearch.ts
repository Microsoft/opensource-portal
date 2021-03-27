//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

const defaultPageSize = 20; // GitHub.com seems to use a value around 33

export default class TeamSearch {
  private teams: any;
  private page: any;
  private tags: any;
  private pageSize: any;
  private set: any;
  private yourTeamsMap: any;
  private phrase: string;
  private sort: string;

  public pageLastTeam: number;
  public pageFirstTeam: number;
  public totalPages: number;
  public totalTeams: number;

  constructor(teams, options) {
    options = options || {};
    this.teams = teams; //teamsWithMembers
    this.pageSize = options.pageSize || defaultPageSize;

    this.phrase = options.phrase;
    this.set = options.set;
    this.yourTeamsMap = options.yourTeamsMap || new Map();
  }

  search(tags, page, sort): Promise<void> {
    this.page = parseInt(page);
    this.tags = tags;
    this.sort = sort
      ? sort.charAt(0).toUpperCase() + sort.slice(1)
      : 'Alphabet';
    return this.filterByType(this.set)
      .filterByPhrase(this.phrase)
      .determinePages()
      ['sortBy' + this.sort]()
      .getPage(this.page);
  }

  determinePages() {
    this.totalPages = Math.ceil(this.teams.length / this.pageSize);
    this.totalTeams = this.teams.length;
    return this;
  }

  getPage(page: number) {
    this.teams = this.teams.slice(
      (page - 1) * this.pageSize,
      (page - 1) * this.pageSize + this.pageSize
    );
    this.pageFirstTeam = 1 + (page - 1) * this.pageSize;
    this.pageLastTeam = this.pageFirstTeam + this.teams.length - 1;
    return this;
  }

  filterByType(setType) {
    let filter = null;
    if (setType === 'your' || setType === 'available') {
      const showIfInSet = setType === 'your';
      filter = (t) => {
        const map = this.yourTeamsMap || new Map();
        return map.has(t.id) === showIfInSet;
      };
    }
    if (filter) {
      this.teams = this.teams.filter(filter);
    }
    return this;
  }

  filterByPhrase(phrase: string) {
    if (phrase) {
      phrase = phrase.toLowerCase();
      this.teams = this.teams.filter((t) => {
        return teamMatchesPhrase(t, phrase);
      });
    }
    return this;
  }

  sortByAlphabet() {
    this.teams.sort((a, b) => {
      let nameA = a.name.toLowerCase();
      let nameB = b.name.toLowerCase();
      if (nameA < nameB) {
        return -1;
      }
      if (nameA > nameB) {
        return 1;
      }
      return 0;
    });
    return this;
  }
}

function teamMatchesPhrase(team, phrase) {
  // Poor man's search, starting with just a raw includes search
  // Assumes that phrase is already lowercase to work
  let string = (
    (team.name || '') +
    (team.description || '') +
    (team.id || '') +
    (team.slug || '')
  ).toLowerCase();
  return string.includes(phrase);
}
