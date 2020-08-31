/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property, LitElement, html } from 'lit-element';
import { MgtTemplatedComponent } from './templatedComponent';

/**
 * Localizes component strings with definied user agent strings
 *
 *
 * @export
 * @class LocalizationService
 */
@customElement('localization-service')
export class LocalizationService extends MgtTemplatedComponent {
  @property({
    attribute: 'strings',
    type: String
  })
  public strings: string;

  //initialize rtl should happen here similar to msal provider implements intializeProvider()
}
