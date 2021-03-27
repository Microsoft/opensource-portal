//
// Copyright (c) Microsoft.
//

const path = require('path');
const pug = require('pug');

export default function RenderHtmlMail(
  basedir,
  viewName,
  options
): Promise<string> {
  return new Promise((resolve, reject) => {
    return renderMailHtml(basedir, viewName, options, (error, html) => {
      return error ? reject(error) : resolve(html);
    });
  });
}

function renderMailHtml(basedir, viewName, options, callback) {
  options = options || {};
  if (!viewName) {
    viewName = 'email';
  }
  if (!options.view) {
    options.view = viewName;
  }
  let html = null;
  try {
    const view = path.join(basedir, `views/email/${viewName}.pug`);
    options.pretty = true;
    options.basedir = basedir;
    html = pug.renderFile(view, options);
  } catch (renderError) {
    return callback(renderError);
  }
  return callback(null, html);
}
