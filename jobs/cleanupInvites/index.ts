//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import Job from './task';
import app from '../../app';

app.runJob(Job, {
  timeoutMinutes: 90,
  defaultDebugOutput: 'restapi',
  insightsPrefix: 'JobOrganizationInvitationsCleanup',
});
