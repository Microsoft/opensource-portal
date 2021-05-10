//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Router } from 'express';
const router: Router = Router();

import { ReposAppRequest } from '../interfaces';

// Enforcing just a single GitHub account per Active Directory user. With
// mild refactoring, this portal could easily support a session selecting
// which link to work with, too.

// interface IRequestWithSessionPlus extends ReposAppRequest {
//   session?: any;
//   linksForCleanup?: any;
// }

router.use((req: ReposAppRequest, res, next) => {
  // TODO: revisit implementation
  return next();

  // The purpose of link cleanup -was-:
  //
  // 1. IF a user's GitHub account associated with this session is different,
  //    and they only have ONE linked account... ask them whether they want to
  //    swap to the other account. This hasn't worked well in practice...
  // 2. IF they are a special user type allowing multiple accounts, see if they
  //    want to onboard or swap their session to another account.

    // req.linksForCleanup = userLinks;
    // if (userLinks.length === 1 && req.user && req.user.github && req.user.github.id !== userLinks[0].ghid) {
    //   if (req.body.unlink && req.body.confirm) {
    //     return unlink(req, userLinks[0], (unlinkError) => {
    //       if (unlinkError) {
    //         next(unlinkError);
    //       } else {
    //         res.redirect('/');
    //       }
    //     });
    //   }
    //   if (req.body.link && req.session.enableMultipleAccounts === true) {
    //     return link(req, req.body.link, (linkError, linkObject) => {
    //       if (linkError) {
    //         next(linkError);
    //       } else {
    //         req.session.selectedGithubId = linkObject.ghid;
    //         res.redirect('/?onboarding=yes');
    //       }
    //     });
    //   }
    //   return renderChangeAccountPage(req, res, userLinks[0]);
    // }
    // if (userLinks.length < 2) {
    //   return res.redirect('/');
    // }
    // // CONSIDER: Make GitHub user calls to see if these users still exist.
    // // EDGE: user renamed their GitHub account... so we may not have their latest GitHub ID, but
    // // it would not create a duplicate link since the GHID fields would be the same.
    // next();
  // });
});

// function renderChangeAccountPage(req: IRequestWithSessionPlus, res, link) {
//   req.individualContext.webContext.render({
//     view: 'removeothergithubaccount',
//     title: 'Exiting GitHub account found',
//     state: {
//       link: link,
//       confirming: req.body.unlink,
//       hideGitHubAccount: true,
//       allowAdditionalAccountLink: req.session && req.session.enableMultipleAccounts ? req.session.enableMultipleAccounts : false,
//     },
//   });
// }

// function renderCleanupPage(req: IRequestWithSessionPlus, res, idToConfirm, links) {
//   links = links || req.linksForCleanup;
//   let twoColumns = [[], []];
//   for (let i = 0; i < links.length; i++) {
//     if (links[i].joined) {
//       try {
//       links[i].joinedDate = new Date(Math.round(links[i].joined));
//       } catch (dateProblem) {
//         /* */
//       }
//     }
//     twoColumns[i % 2].push(links[i]);
//   }
//   req.individualContext.webContext.render({
//     view: 'multiplegithubaccounts',
//     title: 'GitHub Cleanup',
//     state: {
//       linksForCleanupByColumn: twoColumns,
//       numberToRemove: req.linksForCleanup.length - 1,
//       confirming: idToConfirm,
//       hideGitHubAccount: true,
//     },
//   });
// }

// router.get('/', (req: IRequestWithSessionPlus, res) => {
//   renderCleanupPage(req, res, null, null);
// });

// router.post('/', (req: IRequestWithSessionPlus, res, next) => {
//   let action = 'unlink';
//   let id = req.body.unlink;
//   if (!req.body.unlink && req.session && req.session.enableMultipleAccounts === true && req.body.select) {
//     id = req.body.select;
//     action = 'select';
//   }
//   let link = null;
//   let remainingLinks = [];
//   for (let i = 0; i < req.linksForCleanup.length; i++) {
//     if (req.linksForCleanup[i].ghid === id) {
//       link = req.linksForCleanup[i];
//     } else {
//       remainingLinks.push(req.linksForCleanup[i]);
//     }
//   }
//   if (!link) {
//     return next(new Error(`Could not identify the link for GitHub user ${id}.`));
//   }
//   if (action === 'select') {
//     req.session.selectedGithubId = id;
//     return res.redirect('/');
//   }
//   let isConfirming = req.body.confirm === id;
//   if (!isConfirming) {
//     return renderCleanupPage(req, res, id, null);
//   }
//   unlink(req, link, (unlinkError) => {
//     if (unlinkError) {
//       return next(unlinkError);
//     }
//     if (remainingLinks.length > 1) {
//       renderCleanupPage(req, res, null, remainingLinks);
//     } else {
//       req.individualContext.webContext.saveUserAlert(link.ghu + ' has been unlinked. You now have just one GitHub account link.', 'Link cleanup complete', 'success');
//       res.redirect('/');
//     }
//   });
// });

// function unlink(req, link, callback) {
//   const { operations, insights, redisClient, config, githubLibrary } = getProviders(req);
//   const options = {
//     config,
//     redisClient,
//     githubLibrary,
//     operations,
//     link,
//     insights,
//   };
//   new OpenSourceUserContext(options, function (contextError, unlinkContext) {
//     if (contextError) {
//       return callback(contextError);
//     }
//     const account = operations.getAccount(unlinkContext.id.github);
//     const reason = 'Link-cleanup, voluntary unlinking';
//     account.terminate({ reason: reason }, callback);
//   });
// }

export default router;
