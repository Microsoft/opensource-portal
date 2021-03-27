//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import ColorContrastChecker from 'color-contrast-checker';

const WcagHardcodedBodyFontSize = 12;

export default class WcagColorHelper {
  private static _checker = new ColorContrastChecker();
  private static Bright = new WcagColorHelper('#ffffff');
  private static Dark = new WcagColorHelper('#000000');

  private _hex: string;

  constructor(hexColor: string) {
    this._hex = hexColor.startsWith('#') ? hexColor : `#${hexColor}`;
  }

  static BestForegroundColorAsHex(
    backgroundHex: string,
    defaultForegroundColor?: string
  ) {
    const color = new WcagColorHelper(backgroundHex);
    return color.pickBestForegroundAsHex(defaultForegroundColor);
  }

  asHex() {
    return this._hex;
  }

  pickBestForegroundAsHex(defaultForegroundHexColor?: string) {
    let foreground = defaultForegroundHexColor
      ? new WcagColorHelper(defaultForegroundHexColor)
      : WcagColorHelper.Dark;
    if (this.passes(foreground)) {
      return foreground.asHex();
    }
    if (defaultForegroundHexColor) {
      foreground = WcagColorHelper.Dark;
      if (this.passes(foreground)) {
        return foreground.asHex();
      }
    }
    foreground = WcagColorHelper.Bright;
    if (this.passes(foreground)) {
      return foreground.asHex();
    }
    return null;
  }

  passes(compare: WcagColorHelper) {
    return WcagColorHelper._checker.isLevelAA(
      this._hex,
      compare.asHex(),
      WcagHardcodedBodyFontSize
    );
  }
}
