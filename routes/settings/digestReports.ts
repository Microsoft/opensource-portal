//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
const router: Router = Router();

import { getProviders } from '../../transitional';
import { IndividualContext } from '../../user';
import { RequestWithSystemwidePermissions } from '../../interfaces';

// import { buildConsolidatedMap as buildRecipientMap } from '../../jobs/reports/consolidated';
// const buildRecipientMap = require('../../jobs/reports/consolidated').buildRecipientMap;
// import RedisHelper from '../../lib/caching/redis';

import MiddlewareSystemWidePermissions from '../../middleware/github/systemWidePermissions';

router.use(MiddlewareSystemWidePermissions);

interface IRequestWithDigestReports extends RequestWithSystemwidePermissions {
  availableReports?: any;
}

router.use((req: IRequestWithDigestReports, res, next) => {
  const context = req.individualContext as IndividualContext;

  let upn = context && context.corporateIdentity ? context.corporateIdentity.username : null;
  if (!upn) {
    return next(new Error('Must have an active corporate link'));
  }

  // const systemWidePermissions = req.systemWidePermissions;

  // For performance reasons, this current implementation only works
  // when the Redis server is the same for both reports and the
  // app

  const providers = getProviders(req);
  const config = providers.config;

  const reportConfig = config && config.github && config.github.jobs ? config.github.jobs.reports : {};
  return next(new Error('Digest report storage is not enabled for this environment. Reports are not available to be viewed on-demand.'));

  const availableReports = [];

  // return reportRedisClient.getObjectCompressed(reportConfig.witnessEventKey, (error, consolidatedReport) => {
  //   if (!error && !consolidatedReport) {
  //     error = new Error('No recent report is currently available for your account');
  //   }
  //   if (error) {
  //     return next(error);
  //   }
  //   const generated = consolidatedReport.metadata.startedText || 'recently';
  //   const reportsByRecipient = buildRecipientMap(consolidatedReport);
  //   // Hard-coded
  //   const administratorUpn = 'upn:TBD@TBD';
  //   const administratorReport = reportsByRecipient.get(administratorUpn);
  //   if (systemWidePermissions.allowAdministration && administratorReport) {
  //     availableReports.push({
  //       description: `Microsoft-wide report as of ${generated}`,
  //       id: administratorUpn,
  //       report: administratorReport,
  //     });
  //   }
  //   const reportIndex = `upn:${upn.toLowerCase()}`;
  //   const userReport = reportsByRecipient.get(reportIndex);
  //   if (userReport) {
  //     availableReports.push({
  //       description: `Your administrator's report as of ${generated}`,
  //       report: userReport,
  //       id: reportIndex,
  //     });
  //   }

  req.availableReports = availableReports;
  return next();
});

router.get('/administrator/:id', (req: IRequestWithDigestReports, res, next) => {
  const id = req.params.id;
  const availableReports = req.availableReports;
  for (let i = 0; i < availableReports.length; i++) {
    const availableReport = availableReports[i];
    if (availableReport.id === id) {
      return req.individualContext.webContext.render({
        view: 'settings/digestReportView',
        title: availableReport.description,
        state: {
          reportTitle: availableReport.description,
          github: {
            consolidated: availableReport.report,
          },
        },
      });
    }
  }
  return next(new Error('Not found'));
});

router.get('/', (req: IRequestWithDigestReports, res) => {
  const availableReports = req.availableReports;
  req.individualContext.webContext.render({
    view: 'settings/digestReports',
    title: 'Reports',
    state: {
      availableReports,
    },
  });
});

export default router;
