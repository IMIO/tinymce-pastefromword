/**
 * @license Copyright (c) 2003-2023, CKSource Holding sp. z o.o. All rights reserved.
 * Ported from CKEditor 4 (core/tools.js, core/filter.js) for the TinyMCE Paste from Word plugin.
 * Licensed under the terms of the GPL, LGPL or MPL licenses. See LICENSE.md.
 */

// CSS text helpers. They work on plain objects: `{ 'margin-left': '36pt' }`.

const NAMED_COLORS = new Set(
  (
    'aliceblue antiquewhite aqua aquamarine azure beige bisque black blanchedalmond blue ' +
    'blueviolet brown burlywood cadetblue chartreuse chocolate coral cornflowerblue cornsilk ' +
    'crimson cyan darkblue darkcyan darkgoldenrod darkgray darkgreen darkgrey darkkhaki ' +
    'darkmagenta darkolivegreen darkorange darkorchid darkred darksalmon darkseagreen ' +
    'darkslateblue darkslategray darkslategrey darkturquoise darkviolet deeppink deepskyblue ' +
    'dimgray dimgrey dodgerblue firebrick floralwhite forestgreen fuchsia gainsboro ghostwhite ' +
    'gold goldenrod gray green greenyellow grey honeydew hotpink indianred indigo ivory khaki ' +
    'lavender lavenderblush lawngreen lemonchiffon lightblue lightcoral lightcyan ' +
    'lightgoldenrodyellow lightgray lightgreen lightgrey lightpink lightsalmon lightseagreen ' +
    'lightskyblue lightslategray lightslategrey lightsteelblue lightyellow lime limegreen linen ' +
    'magenta maroon mediumaquamarine mediumblue mediumorchid mediumpurple mediumseagreen ' +
    'mediumslateblue mediumspringgreen mediumturquoise mediumvioletred midnightblue mintcream ' +
    'mistyrose moccasin navajowhite navy oldlace olive olivedrab orange orangered orchid ' +
    'palegoldenrod palegreen paleturquoise palevioletred papayawhip peachpuff peru pink plum ' +
    'powderblue purple rebeccapurple red rosybrown royalblue saddlebrown salmon sandybrown ' +
    'seagreen seashell sienna silver skyblue slateblue slategray slategrey snow springgreen ' +
    'steelblue tan teal thistle tomato turquoise violet wheat white whitesmoke yellow ' +
    'yellowgreen windowtext transparent currentcolor'
  ).split(' '),
);

const BORDER_STYLES = ['none', 'hidden', 'dotted', 'dashed', 'solid', 'double', 'groove', 'ridge', 'inset', 'outset'];

const WIDTH_REGEXP = /^(thin|medium|thick|[+-]?\d+(\.\d+)?[a-z%]+|[+-]?0+(\.0+)?|\.\d+[a-z%]+)$/;
const RGBA_REGEXP = /rgba?\(\s*\d+%?\s*,\s*\d+%?\s*,\s*\d+%?\s*(?:,\s*[0-9.]+\s*)?\)/gi;
const HSLA_REGEXP = /hsla?\(\s*[0-9.]+\s*,\s*\d+%\s*,\s*\d+%\s*(?:,\s*[0-9.]+\s*)?\)/gi;

/** Replace `rgb(r, g, b)` colors with hex colors in a CSS text. */
export function convertRgbToHex(styleText) {
  return styleText.replace(/(?:rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\))/gi, (match, red, green, blue) => {
    const color = [red, green, blue].map((part) => ('0' + parseInt(part, 10).toString(16)).slice(-2));
    return '#' + color.join('');
  });
}

/** Lower-case hex colors and expand 3 digit hex colors to 6 digits. */
export function normalizeHex(styleText) {
  return styleText.replace(/#(([0-9a-f]{3}){1,2})($|;|\s+)/gi, (match, hexColor, hexColorPart, separator) => {
    let normalized = hexColor.toLowerCase();
    if (normalized.length === 3) {
      const parts = normalized.split('');
      normalized = [parts[0], parts[0], parts[1], parts[1], parts[2], parts[2]].join('');
    }
    return '#' + normalized + separator;
  });
}

/**
 * Parse a CSS text into an object.
 *
 * @param {string} styleText
 * @param {boolean} [normalize] Lower-case the property names and trim the values.
 * @returns {Object<string, string>}
 */
export function parseCssText(styleText, normalize) {
  const retval = {};

  if (styleText) {
    styleText = normalizeHex(convertRgbToHex(styleText));
  }

  if (!styleText || styleText === ';') {
    return retval;
  }

  styleText.replace(/&quot;/g, '"').replace(/\s*([^:;\s]+)\s*:\s*([^;]+)\s*(?=;|$)/g, (match, name, value) => {
    if (normalize) {
      name = name.toLowerCase();
      if (name === 'font-family') {
        value = value.replace(/\s*,\s*/g, ',');
      }
      value = value.trim();
    }

    retval[name] = value;
    return match;
  });

  return retval;
}

/**
 * Write a styles object as a CSS text.
 *
 * @param {Object<string, string>} styles
 * @param {boolean} [sort]
 * @returns {string}
 */
export function writeCssText(styles, sort) {
  const parts = Object.keys(styles).map((name) => name + ':' + styles[name]);

  if (sort) {
    parts.sort();
  }

  return parts.join('; ');
}

const UNITS_TO_PX = {
  px: 1,
  pt: 96 / 72,
  pc: 16,
  in: 96,
  cm: 96 / 2.54,
  mm: 96 / 25.4,
  q: 96 / 25.4 / 4,
  em: 16,
  rem: 16,
  ex: 8,
  ch: 8,
};

/**
 * Convert a CSS length to a whole number of pixels.
 *
 * Percentages are returned unchanged. Values that cannot be parsed return 0.
 *
 * @param {string|number} cssLength
 * @returns {number|string}
 */
export function convertToPx(cssLength) {
  if (typeof cssLength === 'number') {
    return Math.round(cssLength);
  }

  const value = String(cssLength ?? '').trim();

  if (/%$/.test(value)) {
    return value;
  }

  const match = value.match(/^([+-]?(?:\d+\.?\d*|\.\d+))\s*([a-z]*)$/i);

  if (!match) {
    return 0;
  }

  const number = parseFloat(match[1]);
  const unit = match[2].toLowerCase();
  const factor = unit === '' ? 1 : UNITS_TO_PX[unit];

  if (factor === undefined || Number.isNaN(number)) {
    return 0;
  }

  const result = Math.round(Math.abs(number) * factor);
  return number < 0 ? -result : result;
}

function findColors(value) {
  let ret = [];

  ret = ret.concat(value.match(RGBA_REGEXP) || []);
  ret = ret.concat(value.match(HSLA_REGEXP) || []);

  ret = ret.concat(
    value.split(/\s+/).filter((entry) => {
      if (/^#[a-f0-9]{3}(?:[a-f0-9]{3})?$/i.test(entry)) {
        return true;
      }
      return NAMED_COLORS.has(entry.toLowerCase());
    }),
  );

  return ret;
}

function sideShorthand(value, split) {
  const ret = {};
  const parts = split ? split(value) : value.split(/\s+/);

  const mapStyles = (map) => {
    ret.top = parts[map[0]];
    ret.right = parts[map[1]];
    ret.bottom = parts[map[2]];
    ret.left = parts[map[3]];
  };

  switch (parts.length) {
    case 1:
      mapStyles([0, 0, 0, 0]);
      break;
    case 2:
      mapStyles([0, 1, 0, 1]);
      break;
    case 3:
      mapStyles([0, 1, 2, 1]);
      break;
    case 4:
      mapStyles([0, 1, 2, 3]);
      break;
    default:
      break;
  }

  return ret;
}

/**
 * Split a `margin` shorthand value into its four sides.
 *
 * @param {string} value
 * @returns {{top: string, right: string, bottom: string, left: string}}
 */
export function parseMargin(value) {
  return sideShorthand(value, (width) => {
    return width.match(/(?:-?[.\d]+(?:%|\w*)|auto|inherit|initial|unset|revert)/g) || ['0px'];
  });
}

/**
 * Parse a `background` shorthand value.
 *
 * @param {string} value
 * @returns {{color?: string, unprocessed?: string}}
 */
export function parseBackground(value) {
  const ret = {};
  const colors = findColors(value);

  if (colors.length) {
    ret.color = colors[0];

    for (const colorToken of colors) {
      value = value.replace(colorToken, '');
    }
  }

  value = value.trim();

  if (value) {
    ret.unprocessed = value;
  }

  return ret;
}

/** A parsed border value: width, style and color. */
export class Border {
  constructor(props = {}) {
    this.width = props.width;
    this.style = props.style;
    this.color = props.color;

    if (this.color) {
      this.color = this.color.replace(/windowtext/g, 'black');
    }
  }

  toString() {
    return [this.width, this.style, this.color].filter(Boolean).join(' ');
  }

  /**
   * Parse a `border` shorthand value.
   *
   * @param {string} value
   * @returns {Border}
   */
  static fromCssRule(value) {
    const props = {};
    const input = value.split(/\s+/g);
    const colors = findColors(value);

    if (colors.length) {
      props.color = colors[0];
    }

    for (const val of input) {
      if (!props.style && BORDER_STYLES.includes(val)) {
        props.style = val;
        continue;
      }

      if (!props.width && WIDTH_REGEXP.test(val)) {
        props.width = val;
        continue;
      }
    }

    return new Border(props);
  }

  /**
   * Build a border for each side from `border-width`, `border-style`, `border-color`
   * and the side specific properties found in `styles`.
   *
   * @param {Object<string, string>} styles
   * @param {Border|Object} [fallback] Values used when a property is missing.
   * @returns {Object<string, Border>} Keys: `border-top`, `border-right`, `border-bottom`, `border-left`.
   */
  static splitCssValues(styles, fallback = {}) {
    const types = ['width', 'style', 'color'];
    const sides = ['top', 'right', 'bottom', 'left'];

    const stylesMap = {};
    for (const type of types) {
      const style = styles['border-' + type] || fallback[type];
      stylesMap[type] = style ? sideShorthand(style) : null;
    }

    const ret = {};
    for (const side of sides) {
      const map = {};

      for (const type of types) {
        // A side property has a greater specificity than the shorthand.
        const sideProperty = styles['border-' + side + '-' + type];
        if (sideProperty) {
          map[type] = sideProperty;
        } else {
          map[type] = stylesMap[type] && stylesMap[type][side];
        }
      }

      ret['border-' + side] = new Border(map);
    }

    return ret;
  }
}

/**
 * Replace a `margin` shorthand in `styles` by the four `margin-*` properties.
 * The CKEditor filter transformation used `0` as a fallback width.
 *
 * @param {Object<string, string>} styles
 */
export function splitMarginShorthand(styles) {
  if (!styles.margin) {
    return;
  }

  const widths = styles.margin.match(/(auto|0|(?:-?[.\d]+(?:\w+|%)))/g) || ['0px'];
  const maps = { 1: [0, 0, 0, 0], 2: [0, 1, 0, 1], 3: [0, 1, 2, 1], 4: [0, 1, 2, 3] };
  const map = maps[widths.length];

  if (map) {
    styles['margin-top'] = widths[map[0]];
    styles['margin-right'] = widths[map[1]];
    styles['margin-bottom'] = widths[map[2]];
    styles['margin-left'] = widths[map[3]];
  }

  delete styles.margin;
}

/**
 * Replace a `margin` (or `MARGIN`) shorthand in `style` by the four side properties.
 *
 * @param {Object<string, string>} style
 */
export function parseShorthandMargins(style) {
  const marginCase = style.margin ? 'margin' : style.MARGIN ? 'MARGIN' : false;

  if (marginCase) {
    const margin = parseMargin(style[marginCase]);
    for (const key of Object.keys(margin)) {
      style['margin-' + key] = margin[key];
    }
    delete style[marginCase];
  }
}
