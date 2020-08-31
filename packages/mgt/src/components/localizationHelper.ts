/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property, LitElement } from 'lit-element';
import { MgtTemplatedComponent } from './templatedComponent';
import { LocalizationService } from './localizationService';
import { TelemetryHandlerOptions } from '@microsoft/microsoft-graph-client';

/**
 * Helper class for SocalizationService
 *
 *
 * @export
 * @class LocalizationHelper
 */
export class LocalizationHelper {
  private static _component: string;
  private static _direction: string;

  public static get direction() {
    return this._direction;
  }
  public static set direction(value: string) {
    value = this._direction;
  }
  /**
   * set component reference and intialize
   *
   * @static
   * @memberof LocalizationHelper
   */
  public static intialize(ref: string) {
    this._component = ref;

    let observer = new MutationObserver(updateDirection);
    let observerOptions = {
      attributes: true,
      attributeFilter: ['dir'],
      characterData: false
    };
    observer.observe(document.documentElement, observerOptions);

    let that = this;
    function updateDirection(observer) {
      observer.disconnect();
      that._direction = 'rtl';
    }
    updateDirection(observer);
  }

  /**
   * set specific string values in component instance to localize
   *
   * @memberof LocalizationHelper
   */
  public static localize(str: string) {
    let service: any = document.querySelector('localization-service');
    let strings = service.strings;

    // if current instance component matches user provided component && string matches
    if (service && strings[this._component] && strings[this._component][str]) {
      return strings[this._component][str];
    }
  }
}
