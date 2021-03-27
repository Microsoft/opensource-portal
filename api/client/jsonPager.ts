//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { jsonError } from '../../middleware';

const maxPageSize = 50;
const defaultPageSize = 30;

type Response = {
  json: (obj: any) => void;
};

type Request = {
  query: any;
};

export default class JsonPager<T> {
  pageSize: number;
  page: number;

  // computed:
  total: number;
  lastPage: number;
  begin: number;
  end: number;
  subsetReturnSize: number;

  res: Response;

  constructor(req: Request, res: Response) {
    this.res = res;
    const { query } = req;
    const requestedPageSize = query.pageSize
      ? Number(query.pageSize)
      : defaultPageSize;
    const requestedPage = query.page ? Number(query.page) : 0;
    this.pageSize = Math.min(requestedPageSize, maxPageSize);
    const page = requestedPage || 1;
    if (page < 0 || isNaN(page)) {
      throw jsonError('Invalid page', 400);
    }
    this.page = page;
  }

  slice(array: T[]) {
    this.total = array.length;
    this.lastPage = Math.ceil(this.total / this.pageSize);
    // TODO: this can go past the end, i.e. search while on page 7, it will not return page 1 results
    this.begin = (this.page - 1) * this.pageSize;
    this.end = this.begin + this.pageSize;
    const subset = array.slice(this.begin, this.end);
    this.subsetReturnSize = subset.length;
    return subset;
  }

  sendJson(mappedValues: any[]) {
    if (mappedValues && mappedValues.length !== this.subsetReturnSize) {
      console.warn(
        `The mapped values length ${mappedValues.length} !== ${this.subsetReturnSize} that was computed`
      );
    }
    return this.res.json({
      values: mappedValues,
      total: this.total,
      lastPage: this.lastPage,
      nextPage: this.page + 1, // TODO: should not return if it's the end of the road
      page: this.page,
      pageSize: this.pageSize,
    });
  }

  sliceAndSend(array: T[]) {
    const subset = this.slice(array);
    return this.sendJson(subset);
  }
}
