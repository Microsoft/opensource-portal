//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import _ from 'lodash';
import moment from 'moment';
import octicons from 'octicons'; //const octicons = require('octicons');
import fileSize from 'file-size';
import languageMap from 'language-map';

export default {
  _,
  fileSize: function (bytes) {
    return fileSize(bytes).human();
  },
  languageColor: function (name) {
    const language = languageMap[name];
    if (language) {
      return language.color;
    }
  },
  moment,
  octicon: function (name, optionalWidth, classes, optionalAria) {
    const icon = octicons[name];
    if (!icon || typeof icon.toSVG !== 'function') {
      throw new Error(`Missing octicon ${name}`);
    }
    const options: {
      width?: string | number;
      class?: string;
      'aria-label'?: string;
    } = {};
    if (optionalWidth) {
      options.width = optionalWidth;
    }
    if (optionalAria) {
      options['aria-label'] = optionalAria;
    }
    if (classes) {
      options.class = classes;
    }
    return icon.toSVG(options);
  },
  stripJsSingleQuote: function (value) {
    return value.replace(/\\/g, '').replace(/\'/g, "\\'");
  },
};
