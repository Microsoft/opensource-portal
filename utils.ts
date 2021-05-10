//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import { Request, Response } from 'express';
import fs from 'fs';
import path from 'path';
import { URL } from 'url';
import { DateTime } from 'luxon';
import zlib from 'zlib';
import { ReposAppRequest, IAppSession, IReposError } from './interfaces';
import { getProviders } from './transitional';

const compressionOptions = {
  type: 'gzip',
  params: {
    level: zlib.Z_BEST_SPEED,
  },
};

const hardcodedCorporateTimezone = 'America/Los_Angeles';

export function getOffsetMonthRange(offsetMonths?: number) {
  offsetMonths = offsetMonths || 0;
  const now = new Date();
  const start = DateTime.fromObject({ year: now.getFullYear(), month: 1 + offsetMonths + now.getMonth(), zone: hardcodedCorporateTimezone });
  const end = start.plus({ months: 1 });
  return { start: start.toJSDate(), end: end.toJSDate() };
}

export function daysInMilliseconds(days: number): number {
  return 1000 * 60 * 60 * 24 * days;
}

export function getCurrentQuarter() {
  const now = new Date();
  const quarter = Math.floor((now.getMonth() / 3));
  return quarter;
}

export function getQuarterRange(quarterOfYear: number /* zero-based */) {
  const now = new Date();
  const start = new Date(now.getFullYear(), quarterOfYear * 3, 1);
  const end = new Date(now.getFullYear(), start.getMonth() + 3, 0);
  return [start, end];
}

export function stringOrNumberAsString(value: any) {
  if (typeof(value) === 'number') {
    return (value as number).toString();
  } else if (typeof(value) === 'string') {
    return value;
  }
  const typeName = typeof(value);
  throw new Error(`Unsupported type ${typeName} for value ${value} (stringOrNumberAsString)`);
}

export function stringOrNumberArrayAsStringArray(values: any[]) {
  return values.map(val => stringOrNumberAsString(val));
}

export function requireJson(nameFromRoot: string): any {
  // In some situations TypeScript can load from JSON, but for the transition this is better to reach outside the out directory
  let file = path.resolve(__dirname, nameFromRoot);
  // If within the output directory
  if (fs.existsSync(file)) {
    const content = fs.readFileSync(file, 'utf8');
    return JSON.parse(content);
  }
  file = path.resolve(__dirname, '..', nameFromRoot);
  if (!fs.existsSync(file)) {
    throw new Error(`Cannot find JSON file ${file} to read as a module`);
  }
  const content = fs.readFileSync(file, 'utf8');
  console.warn(`JSON as module (${file}) from project root (NOT TypeScript 'dist' folder)`);
  return JSON.parse(content);
}

// ----------------------------------------------------------------------------
// Returns an integer, random, between low and high (exclusive) - [low, high)
// ----------------------------------------------------------------------------
export function randomInteger(low, high) {
  return Math.floor(Math.random() * (high - low) + low);
};

export function safeLocalRedirectUrl(path: string) {
  if (!path) {
    return;
  }
  const url = new URL(path, 'http://localhost');
  if (url.host !== 'localhost') {
    return;
  }
  return url.search ? `${url.pathname}${url.search}` : url.pathname;
}

// ----------------------------------------------------------------------------
// Session utility: Store the referral URL, if present, and redirect to a new
// location.
// ----------------------------------------------------------------------------
interface IStoreReferrerEventDetails {
  method: string;
  reason: string;
  referer?: string;
  redirect?: string;
}

export function storeReferrer(req: ReposAppRequest, res, redirect, optionalReason) {
  const { insights } = getProviders(req);
  const eventDetails : IStoreReferrerEventDetails = {
    method: 'storeReferrer',
    reason: optionalReason || 'unknown reason',
  };
  const session = req.session as IAppSession;
  if (session && req.headers && req.headers.referer && session.referer !== undefined && !req.headers.referer.includes('/signout') && !session.referer) {
    session.referer = req.headers.referer;
    eventDetails.referer = req.headers.referer;
  }
  if (redirect) {
    eventDetails.redirect = redirect;
    insights?.trackEvent({ name: 'RedirectWithReferrer', properties: eventDetails });
    res.redirect(redirect);
  }
};

export function sortByCaseInsensitive(a: string, b: string) {
  let nameA = a.toLowerCase();
  let nameB = b.toLowerCase();
  if (nameA < nameB) {
    return -1;
  }
  if (nameA > nameB) {
    return 1;
  }
  return 0;
}

// ----------------------------------------------------------------------------
// Session utility: store the original URL
// ----------------------------------------------------------------------------
export function storeOriginalUrlAsReferrer(req: Request, res: Response, redirect: string, optionalReason?: string) {
  storeOriginalUrlAsVariable(req, res, 'referer', redirect, optionalReason);
};

export function redirectToReferrer(req, res, url, optionalReason) {
  url = url || '/';
  const alternateUrl = popSessionVariable(req, res, 'referer');
  const eventDetails = {
    method: 'redirectToReferrer',
    reason: optionalReason || 'unknown reason',
  };
  if (req.insights) {
    req.insights.trackEvent({ name: 'RedirectToReferrer', properties: eventDetails });
  }
  res.redirect(alternateUrl || url);
};

export function storeOriginalUrlAsVariable(req, res, variable, redirect, optionalReason) {
  const eventDetails = {
    method: 'storeOriginalUrlAsVariable',
    variable: variable,
    redirect: redirect,
    reason: optionalReason || 'unknown reason',
  };
  if (req.session && req.originalUrl) {
    req.session[variable] = req.originalUrl;
    eventDetails['ou'] = req.originalUrl;
  }
  if (redirect) {
    if (req.insights) {
      req.insights.trackEvent({ name: 'RedirectFromOriginalUrl', properties: eventDetails });
    }
    res.redirect(redirect);
  }
}

export function popSessionVariable(req, res, variableName) {
  if (req.session && req.session[variableName] !== undefined) {
    const url = req.session[variableName];
    delete req.session[variableName];
    return url;
  }
}

// ----------------------------------------------------------------------------
// Provide our own error wrapper and message for an underlying thrown error.
// Useful for the user-presentable version.
// ----------------------------------------------------------------------------
const errorPropertiesToClone = [
  'stack',
  'status',
];

export function wrapError(error, message, userIntendedMessage?: boolean): IReposError {
  const err: IReposError = new Error(message);
  err.innerError = error;
  if (error) {
    for (let i = 0; i < errorPropertiesToClone.length; i++) {
      const key = errorPropertiesToClone[i];
      const value = error[key];
      if (value && typeof value === 'number') {
        // Store as a string
        err[key] = value.toString();
      } else if (value) {
        err[key] = value;
      }
    }
  }
  if (userIntendedMessage === true) {
    err.skipLog = true;
  }
  return err;
};

// ----------------------------------------------------------------------------
// A destructive removal function for an object. Removes a single key.
// ----------------------------------------------------------------------------
export function stealValue(obj, key) {
  if (obj[key] !== undefined) {
    var val = obj[key];
    delete obj[key];
    return val;
  }
};

// ----------------------------------------------------------------------------
// Given a list of string values, check a string, using a case-insensitive
// comparison.
// ----------------------------------------------------------------------------
export function inListInsensitive(list, value) {
  value = value.toLowerCase();
  for (var i = 0; i < list.length; i++) {
    if (list[i].toLowerCase() === value) {
      return true;
    }
  }
  return false;
};

// ----------------------------------------------------------------------------
// Given a list of lowercase values, check whether a value is present.
// ----------------------------------------------------------------------------
export function isInListAnycaseInLowercaseList(list, value) {
  value = value.toLowerCase();
  for (var i = 0; i < list.length; i++) {
    if (list[i] === value) {
      return true;
    }
  }
  return false;
};

// ----------------------------------------------------------------------------
// Given an array of things that have an `id` property, return a hash indexed
// by that ID.
// ----------------------------------------------------------------------------
export function arrayToHashById(inputArray) {
  var hash = {};
  if (inputArray && inputArray.length) {
    for (var i = 0; i < inputArray.length; i++) {
      if (inputArray[i] && inputArray[i].id) {
        hash[inputArray[i].id] = inputArray[i];
      }
    }
  }
  return hash;
};

// ----------------------------------------------------------------------------
// Obfuscate a string value, optionally leaving a few characters visible.
// ----------------------------------------------------------------------------
export function obfuscate(value, lastCharactersShowCount) {
  if (value === undefined || value === null || value.length === undefined) {
    return value;
  }
  var length = value.length;
  lastCharactersShowCount = lastCharactersShowCount || 0;
  lastCharactersShowCount = Math.min(Math.round(lastCharactersShowCount), length - 1);
  var obfuscated = '';
  for (var i = 0; i < length - lastCharactersShowCount; i++) {
    obfuscated += '*';
  }
  for (var j = length - lastCharactersShowCount; j < length; j++) {
    obfuscated += value[j];
  }
  return obfuscated;
};

// ----------------------------------------------------------------------------
// A very basic breadcrumb stack that ties in to an Express request object.
// ----------------------------------------------------------------------------
export function addBreadcrumb(req, breadcrumbTitle, optionalBreadcrumbLink) {
  if (req === undefined || req.baseUrl === undefined) {
    throw new Error('addBreadcrumb: did you forget to provide a request object instance?');
  }
  if (!optionalBreadcrumbLink && optionalBreadcrumbLink !== false) {
    optionalBreadcrumbLink = req.baseUrl;
  }
  if (!optionalBreadcrumbLink && optionalBreadcrumbLink !== false) {
    optionalBreadcrumbLink = '/';
  }
  var breadcrumbs = req.breadcrumbs;
  if (breadcrumbs === undefined) {
    breadcrumbs = [];
  }
  breadcrumbs.push({
    title: breadcrumbTitle,
    url: optionalBreadcrumbLink,
  });
  req.breadcrumbs = breadcrumbs;
};

export function stackSafeCallback(callback, err, item, extraItem) {
  // Works around RangeError: Maximum call stack size exceeded.
  setImmediate(() => {
    callback(err, item, extraItem);
  });
};

export function createSafeCallbackNoParams(cb) {
  return () => {
    exports.stackSafeCallback(cb);
  };
};

export function sleep(milliseconds: number): Promise<void> {
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      process.nextTick(resolve);
    }, milliseconds);
  });
}

export function ParseReleaseReviewWorkItemId(uri: string): string {
  const safeUrl = new URL(uri);
  const id = safeUrl.searchParams.get('id');
  if (id) {
    return id;
  }
  const pathname = safeUrl.pathname;
  const editIndex = pathname.indexOf('edit/');
  if (editIndex >= 0) {
    return pathname.substr(editIndex + 5);
  }
  if (safeUrl.host === 'osstool.microsoft.com') {
    return null; // Very legacy
  }
  throw new Error(`Unable to parse work item information from: ${uri}`);
}

export function readFileToText(filename: string): Promise<string> {
  return new Promise((resolve, reject) => {
    return fs.readFile(filename, 'utf8', (error, data) => {
      return error ? reject(error) : resolve(data);
    });
  });
}

export function writeTextToFile(filename: string, stringContent: string): Promise<void> {
  return new Promise((resolve, reject) => {
    return fs.writeFile(filename, stringContent, 'utf8', error => {
      if (error) {
        console.warn(`Trouble writing ${filename} ${error}`);
      } else {
        console.log(`Wrote ${filename}`);
      }
      return error ? reject(error) : resolve();
    });
  });
}

export function quitInTenSeconds(successful: boolean) {
  console.log(`Quitting process in 10s... exit code=${successful ? 0 : 1}`);
  return setTimeout(() => {
    process.exit(successful ? 0 : 1);
  }, 1000 * 10 /* 10s */);
}

export function gzipString(value: string): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const val = Buffer.from(value);
    zlib.gzip(val, (gzipError, compressed: Buffer) => {
      return gzipError ? reject(gzipError) : resolve(compressed);
    });
  });
}

export function gunzipBuffer(buffer: Buffer): Promise<string> {
  return new Promise((resolve, reject) => {
    zlib.gunzip(buffer, (unzipError, unzipped) => {
      // Fallback if there is a data error (i.e. it's not compressed)
      if (unzipError && (unzipError as any)?.errno === zlib.Z_DATA_ERROR) {
        const originalValue = buffer.toString();
        return resolve(originalValue);
      } else if (unzipError) {
        return reject(unzipError);
      }
      try {
        const unzippedValue = unzipped.toString();
        return resolve(unzippedValue);
      } catch (otherError) {
        return reject(otherError);
      }
    });
  });
}

export function swapMap(map: Map<string, string>): Map<string, string> {
  const rm = new Map<string, string>();
  for (const [key, value] of map.entries()) {
    rm.set(value, key);
  }
  return rm;
}

export function addArrayToSet<T>(set: Set<T>, array: T[]): Set<T> {
  for (const entry of array) {
    set.add(entry);
  }
  return set;
}
