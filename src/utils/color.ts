/**
 * Color Conversion Utilities
 *
 * Converts between various Office XML color formats and CSS color formats.
 * Handles RGB, HSL, theme colors, and system colors.
 *
 * @module utils/color
 */

/**
 * Color value types in Office XML
 */
export type ColorValue =
  | string // Hex color
  | RGBColor
  | HSLColor
  | ThemeColor
  | SystemColor;

/**
 * RGB color
 */
export interface RGBColor {
  r: number; // 0-255
  g: number; // 0-255
  b: number; // 0-255
  a?: number; // 0-1 (optional alpha)
}

/**
 * HSL color
 */
export interface HSLColor {
  h: number; // 0-360
  s: number; // 0-100
  l: number; // 0-100
  a?: number; // 0-1 (optional alpha)
}

/**
 * Theme color reference
 */
export interface ThemeColor {
  name: string;
  shade?: number; // LumMod in 1000ths (e.g., 50000 = 50% darker)
  tint?: number; // LumOff in 1000ths
}

/**
 * System color
 */
export interface SystemColor {
  name: string;
  lastColor?: string; // Fallback RGB value
}

/**
 * PPTX theme color names
 */
export const THEME_COLORS = {
  bg1: 'background1',
  tx1: 'text1',
  bg2: 'background2',
  tx2: 'text2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hyperlink',
  folHlink: 'followedHyperlink',
} as const;

/**
 * Default theme colors (Office standard)
 */
export const DEFAULT_THEME_COLORS: Record<string, string> = {
  background1: '#FFFFFF',
  text1: '#000000',
  background2: '#FFFFFF',
  text2: '#000000',
  accent1: '#4F81BD',
  accent2: '#C0504D',
  accent3: '#9BBB59',
  accent4: '#8064A2',
  accent5: '#4BACC6',
  accent6: '#F79646',
  hyperlink: '#0000FF',
  followedHyperlink: '#800080',
};

/**
 * System color names
 */
export const SYSTEM_COLORS = {
  windowText: '#000000',
  window: '#FFFFFF',
  windowFrame: '#000000',
  menuText: '#000000',
  menu: '#FFFFFF',
  activeCaption: '#000000',
  inactiveCaption: '#FFFFFF',
} as const;

/**
 * Convert hex color to RGB
 *
 * @param hex - Hex color string (#RGB or #RRGGBB)
 * @returns RGB color object
 */
export function hexToRgb(hex: string): RGBColor | null {
  // Remove # if present
  hex = hex.replace('#', '');

  // Handle 3-character hex
  if (hex.length === 3) {
    hex = hex
      .split('')
      .map((c) => c + c)
      .join('');
  }

  // Validate length
  if (hex.length !== 6) {
    return null;
  }

  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);

  if (isNaN(r) || isNaN(g) || isNaN(b)) {
    return null;
  }

  return { r, g, b };
}

/**
 * Convert RGB to hex color
 *
 * @param rgb - RGB color object
 * @returns Hex color string (#RRGGBB)
 */
export function rgbToHex(rgb: RGBColor): string {
  const toHex = (n: number): string => {
    const hex = Math.round(clamp(n, 0, 255)).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  };

  return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`;
}

/**
 * Convert RGB to CSS color string
 *
 * @param rgb - RGB color object
 * @returns CSS color string (rgb() or rgba())
 */
export function rgbToCss(rgb: RGBColor): string {
  const r = Math.round(clamp(rgb.r, 0, 255));
  const g = Math.round(clamp(rgb.g, 0, 255));
  const b = Math.round(clamp(rgb.b, 0, 255));

  if (rgb.a !== undefined) {
    return `rgba(${r}, ${g}, ${b}, ${clamp(rgb.a, 0, 1)})`;
  }

  return `rgb(${r}, ${g}, ${b})`;
}

/**
 * Convert HSL to RGB
 *
 * @param hsl - HSL color object
 * @returns RGB color object
 */
export function hslToRgb(hsl: HSLColor): RGBColor {
  let { h, s, l } = hsl;
  const a = hsl.a;

  // Normalize values
  h = ((h % 360) + 360) % 360;
  s = clamp(s, 0, 100) / 100;
  l = clamp(l, 0, 100) / 100;

  let r, g, b;

  if (s === 0) {
    // Achromatic (gray)
    r = g = b = l;
  } else {
    const hue2rgb = (p: number, q: number, t: number): number => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };

    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;

    r = hue2rgb(p, q, h / 360 + 1 / 3);
    g = hue2rgb(p, q, h / 360);
    b = hue2rgb(p, q, h / 360 - 1 / 3);
  }

  const rgb: RGBColor = {
    r: Math.round(r * 255),
    g: Math.round(g * 255),
    b: Math.round(b * 255),
  };

  if (a !== undefined) {
    rgb.a = a;
  }

  return rgb;
}

/**
 * Convert RGB to HSL
 *
 * @param rgb - RGB color object
 * @returns HSL color object
 */
export function rgbToHsl(rgb: RGBColor): HSLColor {
  const r = clamp(rgb.r, 0, 255) / 255;
  const g = clamp(rgb.g, 0, 255) / 255;
  const b = clamp(rgb.b, 0, 255) / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const delta = max - min;

  let h = 0;
  let s = 0;
  const l = (max + min) / 2;

  if (delta !== 0) {
    s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min);

    switch (max) {
      case r:
        h = ((g - b) / delta + (g < b ? 6 : 0)) / 6;
        break;
      case g:
        h = ((b - r) / delta + 2) / 6;
        break;
      case b:
        h = ((r - g) / delta + 4) / 6;
        break;
    }
  }

  const hsl: HSLColor = {
    h: Math.round(h * 360),
    s: Math.round(s * 100),
    l: Math.round(l * 100),
  };

  if (rgb.a !== undefined) {
    hsl.a = rgb.a;
  }

  return hsl;
}

/**
 * Parse Office XML color to CSS color
 *
 * @param colorObj - Office XML color object
 * @param themeColors - Custom theme colors (optional)
 * @returns CSS color string
 */
export function parseOfficeColor(
  colorObj: any,
  themeColors?: Record<string, string>
): string {
  if (!colorObj) return '';

  // Check for srgbClr (hex color)
  if (colorObj['a:srgbClr']) {
    const hex = colorObj['a:srgbClr']['val'];
    if (hex) return `#${hex}`;
  }

  // Check for scrgbClr (percent RGB)
  if (colorObj['a:scrgbClr']) {
    const r = parseInt(colorObj['a:scrgbClr']['r'], 10) || 0;
    const g = parseInt(colorObj['a:scrgbClr']['g'], 10) || 0;
    const b = parseInt(colorObj['a:scrgbClr']['b'], 10) || 0;
    return rgbToCss({
      r: (r / 100000) * 255,
      g: (g / 100000) * 255,
      b: (b / 100000) * 255,
    });
  }

  // Check for sysClr (system color)
  if (colorObj['a:sysClr']) {
    const lastClr = colorObj['a:sysClr']['lastClr'];
    if (lastClr) return `#${lastClr}`;

    const sysColorName = colorObj['a:sysClr']['val'];
    if (sysColorName && SYSTEM_COLORS[sysColorName as keyof typeof SYSTEM_COLORS]) {
      return SYSTEM_COLORS[sysColorName as keyof typeof SYSTEM_COLORS];
    }
  }

  // Check for schemeClr (theme color)
  if (colorObj['a:schemeClr']) {
    const schemeName = colorObj['a:schemeClr']['val'];
    let color = resolveThemeColor(schemeName, themeColors);

    // Apply lumMod (luminance modulation - darker)
    if (colorObj['a:schemeClr']['a:lumMod']) {
      const lumMod = parseInt(colorObj['a:schemeClr']['a:lumMod']['val'], 10) || 100000;
      color = applyLuminanceModulation(color, lumMod);
    }

    // Apply lumOff (luminance offset - lighter)
    if (colorObj['a:schemeClr']['a:lumOff']) {
      const lumOff = parseInt(colorObj['a:schemeClr']['a:lumOff']['val'], 10) || 0;
      color = applyLuminanceOffset(color, lumOff);
    }

    return color;
  }

  // Check for prstClr (preset color)
  if (colorObj['a:prstClr']) {
    const presetName = colorObj['a:prstClr']['val'];
    return getPresetColor(presetName);
  }

  return '';
}

/**
 * Resolve theme color name to hex color
 *
 * @param schemeName - Theme color name
 * @param customTheme - Custom theme colors (optional)
 * @returns Hex color string
 */
export function resolveThemeColor(
  schemeName: string,
  customTheme?: Record<string, string>
): string {
  // Check custom theme first
  if (customTheme && customTheme[schemeName]) {
    return customTheme[schemeName];
  }

  // Fall back to default theme colors
  const mappedName = THEME_COLORS[schemeName as keyof typeof THEME_COLORS];
  if (mappedName && DEFAULT_THEME_COLORS[mappedName]) {
    return DEFAULT_THEME_COLORS[mappedName];
  }

  // Fallback to black
  return '#000000';
}

/**
 * Apply luminance modulation (darken)
 *
 * @param color - Hex color
 * @param lumMod - Luminance modulation in 1000ths (0-100000)
 * @returns Darkened hex color
 */
export function applyLuminanceModulation(color: string, lumMod: number): string {
  const rgb = hexToRgb(color);
  if (!rgb) return color;

  const factor = clamp(lumMod / 100000, 0, 1);

  return rgbToHex({
    r: rgb.r * factor,
    g: rgb.g * factor,
    b: rgb.b * factor,
  });
}

/**
 * Apply luminance offset (lighten)
 *
 * @param color - Hex color
 * @param lumOff - Luminance offset in 1000ths (0-100000)
 * @returns Lightened hex color
 */
export function applyLuminanceOffset(color: string, lumOff: number): string {
  const rgb = hexToRgb(color);
  if (!rgb) return color;

  const offset = clamp(lumOff / 100000, 0, 1);

  return rgbToHex({
    r: rgb.r + (255 - rgb.r) * offset,
    g: rgb.g + (255 - rgb.g) * offset,
    b: rgb.b + (255 - rgb.b) * offset,
  });
}

/**
 * Get preset color by name
 *
 * @param presetName - Preset color name
 * @returns Hex color string
 */
export function getPresetColor(presetName: string): string {
  const presetColors: Record<string, string> = {
    aliceBlue: '#F0F8FF',
    antiqueWhite: '#FAEBD7',
    aqua: '#00FFFF',
    aquamarine: '#7FFFD4',
    azure: '#F0FFFF',
    beige: '#F5F5DC',
    bisque: '#FFE4C4',
    black: '#000000',
    blanchedAlmond: '#FFEBCD',
    blue: '#0000FF',
    blueViolet: '#8A2BE2',
    brown: '#A52A2A',
    burlyWood: '#DEB887',
    cadetBlue: '#5F9EA0',
    chartreuse: '#7FFF00',
    chocolate: '#D2691E',
    coral: '#FF7F50',
    cornflowerBlue: '#6495ED',
    cornsilk: '#FFF8DC',
    crimson: '#DC143C',
    cyan: '#00FFFF',
    darkBlue: '#00008B',
    darkCyan: '#008B8B',
    darkGoldenrod: '#B8860B',
    darkGray: '#A9A9A9',
    darkGreen: '#006400',
    darkGrey: '#A9A9A9',
    darkKhaki: '#BDB76B',
    darkMagenta: '#8B008B',
    darkOliveGreen: '#556B2F',
    darkOrange: '#FF8C00',
    darkOrchid: '#9932CC',
    darkRed: '#8B0000',
    darkSalmon: '#E9967A',
    darkSeaGreen: '#8FBC8F',
    darkSlateBlue: '#483D8B',
    darkSlateGray: '#2F4F4F',
    darkSlateGrey: '#2F4F4F',
    darkTurquoise: '#00CED1',
    darkViolet: '#9400D3',
    deepPink: '#FF1493',
    deepSkyBlue: '#00BFFF',
    dimGray: '#696969',
    dimGrey: '#696969',
    dodgerBlue: '#1E90FF',
    firebrick: '#B22222',
    floralWhite: '#FFFAF0',
    forestGreen: '#228B22',
    fuchsia: '#FF00FF',
    gainsboro: '#DCDCDC',
    ghostWhite: '#F8F8FF',
    gold: '#FFD700',
    goldenrod: '#DAA520',
    gray: '#808080',
    green: '#008000',
    greenYellow: '#ADFF2F',
    grey: '#808080',
    honeydew: '#F0FFF0',
    hotPink: '#FF69B4',
    indianRed: '#CD5C5C',
    indigo: '#4B0082',
    ivory: '#FFFFF0',
    khaki: '#F0E68C',
    lavender: '#E6E6FA',
    lavenderBlush: '#FFF0F5',
    lawnGreen: '#7CFC00',
    lemonChiffon: '#FFFACD',
    lightBlue: '#ADD8E6',
    lightCoral: '#F08080',
    lightCyan: '#E0FFFF',
    lightGoldenrodYellow: '#FAFAD2',
    lightGray: '#D3D3D3',
    lightGreen: '#90EE90',
    lightGrey: '#D3D3D3',
    lightPink: '#FFB6C1',
    lightSalmon: '#FFA07A',
    lightSeaGreen: '#20B2AA',
    lightSkyBlue: '#87CEFA',
    lightSlateGray: '#778899',
    lightSlateGrey: '#778899',
    lightSteelBlue: '#B0C4DE',
    lightYellow: '#FFFFE0',
    lime: '#00FF00',
    limeGreen: '#32CD32',
    linen: '#FAF0E6',
    magenta: '#FF00FF',
    maroon: '#800000',
    mediumAquamarine: '#66CDAA',
    mediumBlue: '#0000CD',
    mediumOrchid: '#BA55D3',
    mediumPurple: '#9370DB',
    mediumSeaGreen: '#3CB371',
    mediumSlateBlue: '#7B68EE',
    mediumSpringGreen: '#00FA9A',
    mediumTurquoise: '#48D1CC',
    mediumVioletRed: '#C71585',
    midnightBlue: '#191970',
    mintCream: '#F5FFFA',
    mistyRose: '#FFE4E1',
    moccasin: '#FFE4B5',
    navajoWhite: '#FFDEAD',
    navy: '#000080',
    oldLace: '#FDF5E6',
    olive: '#808000',
    oliveDrab: '#6B8E23',
    orange: '#FFA500',
    orangeRed: '#FF4500',
    orchid: '#DA70D6',
    paleGoldenrod: '#EEE8AA',
    paleGreen: '#98FB98',
    paleTurquoise: '#AFEEEE',
    paleVioletRed: '#DB7093',
    papayaWhip: '#FFEFD5',
    peachPuff: '#FFDAB9',
    peru: '#CD853F',
    pink: '#FFC0CB',
    plum: '#DDA0DD',
    powderBlue: '#B0E0E6',
    purple: '#800080',
    red: '#FF0000',
    rosyBrown: '#BC8F8F',
    royalBlue: '#4169E1',
    saddleBrown: '#8B4513',
    salmon: '#FA8072',
    sandyBrown: '#F4A460',
    seaGreen: '#2E8B57',
    seaShell: '#FFF5EE',
    sienna: '#A0522D',
    silver: '#C0C0C0',
    skyBlue: '#87CEEB',
    slateBlue: '#6A5ACD',
    slateGray: '#708090',
    slateGrey: '#708090',
    snow: '#FFFAFA',
    springGreen: '#00FF7F',
    steelBlue: '#4682B4',
    tan: '#D2B48C',
    teal: '#008080',
    thistle: '#D8BFD8',
    tomato: '#FF6347',
    turquoise: '#40E0D0',
    violet: '#EE82EE',
    wheat: '#F5DEB3',
    white: '#FFFFFF',
    whiteSmoke: '#F5F5F5',
    yellow: '#FFFF00',
    yellowGreen: '#9ACD32',
  };

  return presetColors[presetName] || '#000000';
}

/**
 * Clamp value between min and max
 *
 * @param value - Value to clamp
 * @param min - Minimum value
 * @param max - Maximum value
 * @returns Clamped value
 */
function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Convert color to RGBA with opacity
 *
 * @param color - Color string (hex, rgb, or named)
 * @param opacity - Opacity (0-1)
 * @returns RGBA color string
 */
export function colorWithOpacity(color: string, opacity: number): string {
  const rgb = hexToRgb(color.replace('#', ''));
  if (!rgb) return `rgba(0, 0, 0, ${opacity})`;

  return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${clamp(opacity, 0, 1)})`;
}

/**
 * Parse gradient colors from Office XML
 *
 * @param gradFill - Gradient fill object
 * @param themeColors - Custom theme colors (optional)
 * @returns Array of CSS color strings
 */
export function parseGradientColors(
  gradFill: any,
  themeColors?: Record<string, string>
): string[] {
  const colors: string[] = [];
  const gsLst = gradFill?.['a:gsLst']?.['a:gs'];

  if (!gsLst) return colors;

  const stops = Array.isArray(gsLst) ? gsLst : [gsLst];

  for (const stop of stops) {
    const color = parseOfficeColor(stop, themeColors);
    if (color) {
      colors.push(color);
    }
  }

  return colors;
}

/**
 * Determine if color is light or dark
 *
 * @param color - Color string (hex)
 * @returns true if color is light
 */
export function isLightColor(color: string): boolean {
  const rgb = hexToRgb(color.replace('#', ''));
  if (!rgb) return false;

  // Calculate luminance
  const luminance = (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b) / 255;
  return luminance > 0.5;
}

/**
 * Normalize color string to ensure # prefix
 *
 * @param color - Color string (hex without or with #)
 * @returns Normalized color string with # prefix
 */
export function normalizeColor(color: string | undefined | null): string {
  if (!color) return '#000000';
  if (color.startsWith('#')) return color;
  return `#${color}`;
}
