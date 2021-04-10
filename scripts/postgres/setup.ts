//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import prompt from 'async-prompt';
import escape from 'pg-escape';
import { Client } from 'pg';
import fs from 'fs';
import path from 'path';

initialize();

async function setup() {
  const sql = fs.readFileSync(path.join(__dirname, '../../../pg.sql'), 'utf8');

  const server = await prompt('postgres server: ');
  const adminUsername = await prompt('admin user: ');
  const adminPassword = await prompt.password('admin password: ');
  const client = new Client({
    user: adminUsername,
    host: server,
    database: 'postgres',
    password: adminPassword,
    ssl: true,
  });
  await client.connect();

  const db = await prompt('database name: ');
  try {
    await client.query(`create database ${db}`);
  } catch (createError) {
    console.log(createError.message);
  }

  const newUsername = await prompt('new runtime username: ');
  const newPassword = await prompt.password('new password: ');

  try {
    await client.query(escape(`create user ${newUsername} with password %L`, newPassword));
  } catch (createUserError) {
    console.log(createUserError.message);
  }

  const dbc = new Client({
    user: adminUsername,
    host: server,
    database: db,
    password: adminPassword,
    ssl: true,
  });
  await dbc.connect();
  try {
    await dbc.query(`
      ALTER DEFAULT PRIVILEGES
      FOR USER ${newUsername}
      IN SCHEMA public
      GRANT SELECT, INSERT, UPDATE, DELETE ON TABLES TO ${newUsername};
    `);
  } catch (alterError) {
    console.log(alterError.message);
  }

  console.log(sql);
  try {
    const res = await dbc.query(sql);
    console.log(res);
  } catch (error) {
    console.dir(error);
  }

  console.log('Done with SQL. Granting user rights...');

  try {
    await dbc.query(`
      GRANT SELECT, INSERT, UPDATE, DELETE ON ALL TABLES IN SCHEMA public TO ${newUsername};
    `);
  } catch (alterError) {
    console.log(alterError.message);
  }

  return null;
}

function initialize() {
  return setup().then(ok => {
    console.log('OK');
    process.exit(0);
  }).catch(error => {
    console.error(error);
    console.dir(error);
    process.exit(1);
  });
}
