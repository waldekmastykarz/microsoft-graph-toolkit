/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  accentBaseColor,
  baseLayerLuminance,
  neutralBaseColor,
  StandardLuminance,
  SwatchRGB
} from '@fluentui/web-components';
// @microsoft/fast-colors is a transitive dependency of @fluentui/web-components, no need to explicitly add it to package.json
import { parseColorHexRGB } from '@microsoft/fast-colors';

/**
 * Available predefined themes
 */
type Theme = 'light' | 'dark' | 'default' | 'contrast';

/**
 * Helper function to apply fluent ui theme to an element
 *
 * @export
 * @param {Theme} theme - theme name, if an unknown theme is provided, the light theme will be applied
 * @param {HTMLElement} [element=document.body]
 */
export const applyTheme = (theme: Theme, element: HTMLElement = document.body): void => {
  const settings = getThemeSettings(theme);
  applyColorScheme(settings, element);
};

/**
 * Simple data holder for theme settings
 */
type ColorScheme = {
  /**
   * Hex color string for accent base color
   *
   * @type {string}
   */
  accentBaseColor: string;
  /**
   * Hex color string for neutral base color
   *
   * @type {string}
   */
  neutralBaseColor: string;
  /**
   * Base layer luminance for the theme
   * in the range of 0 to 1
   *
   * @type {number}
   */
  baseLayerLuminance: number;

  designTokenOverrides?: Record<string, string>;
};

/**
 * Helper function to apply fluent ui color scheme to an element
 *
 * @param {ColorScheme} settings
 * @param {HTMLElement} [element=document.body]
 */
const applyColorScheme = (settings: ColorScheme, element: HTMLElement = document.body): void => {
  accentBaseColor.setValueFor(element, SwatchRGB.from(parseColorHexRGB(settings.accentBaseColor)));
  neutralBaseColor.setValueFor(element, SwatchRGB.from(parseColorHexRGB(settings.neutralBaseColor)));
  baseLayerLuminance.setValueFor(element, settings.baseLayerLuminance);
  // put this work on the macro queue so that it happens after promise based/reactive work of setting the base colors above
  if (settings.designTokenOverrides) {
    setTimeout(() => {
      Object.entries(settings.designTokenOverrides).forEach(([key, value]) => {
        element.style.setProperty(key, value);
      });
    });
  }
};

/**
 * Helper function to translate theme name to theme settings
 *
 * @param {Theme} theme
 * @return {*}  {ThemeSettings}
 */
const getThemeSettings = (theme: Theme): ColorScheme => {
  switch (theme) {
    case 'contrast':
      return {
        accentBaseColor: '#7f85f5',
        neutralBaseColor: '#adadad',
        baseLayerLuminance: StandardLuminance.DarkMode
      };
    case 'default': // this is the Teams light theme
      return {
        accentBaseColor: '#5b5fc7',
        neutralBaseColor: '#616161',
        baseLayerLuminance: StandardLuminance.LightMode
      };
    case 'dark': // Both MGT default dark and Teams Dark theme
      return {
        accentBaseColor: '#479ef5',
        neutralBaseColor: '#adadad',
        baseLayerLuminance: StandardLuminance.DarkMode,
        designTokenOverrides: {
          '--accent-fill-rest': '#115ea3',
          '--accent-fill-hover': '#0f6cbd',
          '--accent-fill-active': '#0c3b5e',
          '--accent-fill-focus': '#0f548c',
          '--accent-foreground-rest': '#479EF5',
          '--accent-foreground-hover': '#62abf5',
          '--accent-foreground-active': '#2886de',
          '--accent-foreground-focus': '#479ef5',
          '--accent-stroke-control-rest': '#115ea3',
          '--accent-stroke-control-hover': '#0f6cbd',
          '--accent-stroke-control-active': '#0c3b5e',
          '--accent-stroke-control-focus': '#0f548c',
          // foreground on accents
          '--foreground-on-accent-rest': '#ffffff',
          '--foreground-on-accent-active': '#ffffff',
          '--foreground-on-accent-rest-large': '#ffffff',
          '--foreground-on-accent-hover': '#ffffff',
          '--foreground-on-accent-hover-large': '#ffffff',
          '--foreground-on-accent-active-large': '#ffffff',
          '--foreground-on-accent-focus': '#ffffff',
          '--foreground-on-accent-focus-large': '#ffffff'
        }
      };
    case 'light':
    default:
      return {
        accentBaseColor: '#0f6cbd',
        neutralBaseColor: '#616161',
        baseLayerLuminance: StandardLuminance.LightMode
      };
  }
};
