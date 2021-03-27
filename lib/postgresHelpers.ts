//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

const debug = require('debug')('pg');

import { InnerError } from '../transitional';

export function PostgresPoolQuerySingleRow(
  pool,
  sql: string,
  values: any[],
  callback
) {
  PostgresPoolQuery(pool, sql, values, (error, results) => {
    if (error) {
      return callback(error);
    }
    const len = results.rowCount;
    if (len === 1) {
      return callback(null, results.rows[0]);
    } else if (len === 0) {
      const notFoundError = new Error('The query did not return a result');
      notFoundError['status'] = 404;
      notFoundError['sqlStatement'] = sql;
      notFoundError['sqlValues'] = values;
      return callback(notFoundError);
    }
    const tooManyRows = new Error(
      `Only one row should be returned; ${len} rows were returned`
    );
    tooManyRows['status'] = 412;
    tooManyRows['sqlStatement'] = sql;
    tooManyRows['sqlValues'] = values;
    return callback(tooManyRows);
  });
}

export function PostgresPoolQuerySingleRowAsync(
  pool,
  sql: string,
  values: any[]
): Promise<any> {
  return new Promise((resolve, reject) => {
    PostgresPoolQuerySingleRow(pool, sql, values, (error, results) => {
      return error ? reject(error) : resolve(results);
    });
  });
}

export function PostgresPoolQuery(pool, sql: string, values: any[], callback) {
  if (!pool) {
    throw new Error('No Postgres pool provided');
  }
  if (!callback && typeof values === 'function') {
    callback = values;
    values = [];
  }
  pool.connect(function (connectError, client, release) {
    if (connectError) {
      return callback(connectError);
    }
    debug(sql);
    debug(values);
    client.query(sql, values, function (queryError, results) {
      release();
      if (queryError) {
        const err: InnerError = new Error(
          queryError.message /* Postgres provider never leaks SQL statements thankfully */ ||
            'There was an error querying a database'
        );
        err.inner = queryError;
        if (queryError.position) {
          err['position'] = queryError.position;
        }
        if (queryError.message) {
          err['sqlMessage'] = queryError.message;
          err['sqlStatement'] = sql;
          err['sqlValues'] = values;
        }
        return callback(err);
      }
      return callback(null, results);
    });
  });
}

export function PostgresPoolQueryAsync(
  pool,
  sql: string,
  values: any[]
): Promise<any> {
  return new Promise((resolve, reject) => {
    PostgresPoolQuery(pool, sql, values, (error, results) => {
      if (results && results['rows'] && results['rows'].length !== undefined) {
        debug(`rows: ${results['rows'].length}`);
      }
      return error ? reject(error) : resolve(results);
    });
  });
}
