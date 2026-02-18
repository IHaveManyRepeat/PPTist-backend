/**
 * PPTX XML Parser Service
 *
 * Parses PPTX XML structure using fast-xml-parser and extracts
 * slide elements, layouts, styles, and relationships.
 *
 * @module services/pptx/parser
 */

import { XMLParser } from 'fast-xml-parser';
import { logger } from '../../utils/logger';
import { PPTXValidationError } from '../../utils/errors';
import { generateElementId } from '../../utils/id-generator';
import type { ExtractedPPTX } from './unzip.js';

/**
 * PPTX XML namespace
 */
export const PPTX_NAMESPACES = {
  A: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  P: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  R: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const;

/**
 * Parsed slide element
 */
export interface ParsedSlide {
  id: string;
  index: number;
  elements: ParsedElement[];
  background?: SlideBackground;
  layoutId?: string;
  transition?: SlideTransition;
  notes?: string;
}

/**
 * Parsed element (base type for all slide elements)
 */
export interface ParsedElement {
  type: ElementType;
  id: string;
  name?: string;
  hidden?: boolean;
  locked?: boolean;
  zIndex?: number;
  [key: string]: any; // Additional type-specific properties
}

/**
 * Element type enum
 */
export enum ElementType {
  TEXT = 'text',
  IMAGE = 'image',
  SHAPE = 'shape',
  LINE = 'line',
  CHART = 'chart',
  TABLE = 'table',
  GROUP = 'group',
  VIDEO = 'video',
  AUDIO = 'audio',
  UNKNOWN = 'unknown',
}

/**
 * Slide background
 */
export interface SlideBackground {
  type?: 'solid' | 'gradient' | 'image' | 'pattern';
  color?: string;
  gradientColors?: string[];
  imageRef?: string;
}

/**
 * Slide transition
 */
export interface SlideTransition {
  type?: string;
  duration?: number;
  direction?: string;
}

/**
 * Relationship entry
 */
export interface Relationship {
  id: string;
  type: string;
  target: string;
}

/**
 * Parser options
 */
export interface ParserOptions {
  /**
   * Whether to preserve whitespace in text (default: false)
   */
  preserveWhitespace?: boolean;

  /**
   * Whether to parse relationships (default: true)
   */
  parseRelationships?: boolean;

  /**
   * Whether to parse styles (default: true)
   */
  parseStyles?: boolean;

  /**
   * Maximum element depth (default: 10)
   */
  maxDepth?: number;
}

/**
 * Default parser options
 */
const DEFAULT_OPTIONS: Required<ParserOptions> = {
  preserveWhitespace: false,
  parseRelationships: true,
  parseStyles: true,
  maxDepth: 10,
};

/**
 * Create fast-xml-parser instance with PPTX-optimized configuration
 */
function createXMLParser(options: ParserOptions = {}): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '', // Keep attribute names clean
    textNodeName: '#text',
    ignoreDeclaration: true,
    trimValues: !options.preserveWhitespace,
    parseTagValue: true,
    parseAttributeValue: true,
    allowBooleanAttributes: true,
    numberParseOptions: {
      leadingZeros: true,
      hex: true,
    },
    // Always return arrays for repeatable elements
    isArray: (name, jpath) => {
      // Elements that should always be arrays
      const arrayElements = [
        'a:p',
        'a:r',
        'pic',
        'sp',
        'cnv',
        'graphic',
        'tbl',
        'tr',
        'tc',
      ];
      return arrayElements.includes(name);
    },
  });
}

/**
 * Parse slide XML and extract elements
 *
 * @param xml - Slide XML content
 * @param slideIndex - Slide index (1-based)
 * @param options - Parser options
 * @returns ParsedSlide
 */
export function parseSlideXML(
  xml: string,
  slideIndex: number,
  options: ParserOptions = {}
): ParsedSlide {
  const parser = createXMLParser(options);

  try {
    const parsed = parser.parse(xml);
    // slideRoot is the p:sld object if it exists, otherwise it's the parsed root
    const slideRoot = parsed['p:sld'] || parsed;

    logger.debug(`Parsing slide ${slideIndex}`, {
      hasRoot: !!slideRoot,
    });

    const slide: ParsedSlide = {
      id: slideRoot['id'] || `slide-${slideIndex}`,
      index: slideIndex,
      elements: [],
    };

    // Extract background
    if (slideRoot['p:cSld']) {
      slide.background = parseBackground(
        slideRoot['p:cSld']['p:spTree']
      );
    }

    // Extract elements from spTree (shape tree)
    const spTree = slideRoot['p:cSld']?.['p:spTree'];
    if (spTree) {
      // Extract raw spTree XML to preserve element order
      const spTreeXml = extractSpTreeXml(xml);
      slide.elements = parseShapeTree(spTree, spTreeXml, options);
    }

    // Extract transition
    if (slideRoot['p:transition']) {
      slide.transition = parseTransition(slideRoot['p:transition']);
    }

    logger.debug(`Parsed slide ${slideIndex}`, {
      elementCount: slide.elements.length,
      hasBackground: !!slide.background,
    });

    return slide;
  } catch (error) {
    logger.error(`Failed to parse slide ${slideIndex} XML`, {
      error: error instanceof Error ? error.message : String(error),
    });
    throw new PPTXValidationError(
      `Failed to parse slide ${slideIndex}: ${error}`,
      'PARSE_ERROR'
    );
  }
}

/**
 * Extract spTree XML fragment from slide XML
 * Used to preserve element order before fast-xml-parser reorders elements by type
 */
function extractSpTreeXml(slideXml: string): string {
  const match = slideXml.match(/<p:spTree[^>]*>[\s\S]*?<\/p:spTree>/);
  return match ? match[0] : '';
}

/**
 * Element tag to type mapping
 */
const TAG_TO_TYPE: Record<string, ElementType> = {
  'p:sp': ElementType.SHAPE,
  'p:pic': ElementType.IMAGE,
  'p:cxnSp': ElementType.LINE,
  'p:graphicFrame': ElementType.CHART,
  'p:grpSp': ElementType.GROUP,
};

/**
 * Extract element order from raw XML string
 * Returns an array of { tag, index } representing the order elements appear in XML
 *
 * Example return: [{ tag: 'p:sp', index: 0 }, { tag: 'p:pic', index: 0 }, { tag: 'p:sp', index: 1 }]
 */
function extractElementOrder(spTreeXml: string): Array<{ tag: string; index: number }> {
  const elementTags = Object.keys(TAG_TO_TYPE);
  const order: Array<{ tag: string; index: number }> = [];
  const counters: Record<string, number> = {};

  // Regex matches opening tags (handles attributes)
  const tagRegex = /<(p:sp|p:pic|p:cxnSp|p:graphicFrame|p:grpSp)(?:\s[^>]*)?>/g;
  let match;

  while ((match = tagRegex.exec(spTreeXml)) !== null) {
    const tag = match[1];
    counters[tag] = (counters[tag] || 0);
    order.push({ tag, index: counters[tag] });
    counters[tag]++;
  }

  return order;
}

/**
 * Parse shape tree to extract elements
 *
 * @param spTree - Parsed spTree object from fast-xml-parser
 * @param spTreeXml - Raw XML string of spTree (used to preserve element order)
 * @param options - Parser options
 */
function parseShapeTree(
  spTree: any,
  spTreeXml: string,
  options: ParserOptions
): ParsedElement[] {
  const elements: ParsedElement[] = [];

  // Extract element order from raw XML
  const elementOrder = extractElementOrder(spTreeXml);

  // Pre-process: convert each tag's elements to array for indexed access
  const elementsByTag: Record<string, any[]> = {};
  for (const tag of Object.keys(TAG_TO_TYPE)) {
    const items = spTree[tag];
    elementsByTag[tag] = Array.isArray(items) ? items : items ? [items] : [];
  }

  // Rebuild elements array in original XML order
  for (const { tag, index } of elementOrder) {
    const item = elementsByTag[tag]?.[index];
    if (!item) continue;

    const type = TAG_TO_TYPE[tag];
    const element = parseElement(item, type, options);
    if (element) {
      element.zIndex = elements.length; // Set zIndex based on array position
      elements.push(element);
    }
  }

  return elements;
}

/**
 * Parse individual element
 */
function parseElement(
  item: any,
  type: ElementType,
  options: ParserOptions
): ParsedElement | null {
  try {
    const baseProps = extractBaseProperties(item);

    switch (type) {
      case ElementType.SHAPE:
        return parseShape(item, baseProps);
      case ElementType.IMAGE:
        return parseImage(item, baseProps);
      case ElementType.CHART:
        return parseChartOrTable(item, baseProps);
      case ElementType.LINE:
        return parseLine(item, baseProps);
      case ElementType.GROUP:
        return parseGroup(item, baseProps, options);
      default:
        logger.warn(`Unknown element type: ${type}`);
        return null;
    }
  } catch (error) {
    logger.error(`Failed to parse element`, {
      error: error instanceof Error ? error.message : String(error),
      type,
    });
    return null;
  }
}

/**
 * Extract base properties common to all elements
 */
function extractBaseProperties(item: any): Partial<ParsedElement> {
  const nvSpPr = item['p:nvSpPr'] || item['p:nvPicPr'] || item['p:nvGraphicFramePr'];
  const cNvPr = nvSpPr?.['p:cNvPr'] || {};

  return {
    id: ((cNvPr['id'] as string) || generateElementId()),
    name: cNvPr['name'],
    hidden: cNvPr['hidden'] === '1' || cNvPr['hidden'] === true,
    locked: cNvPr['locked'] === '1' || cNvPr['locked'] === true,
  };
}

/**
 * Parse shape element
 */
function parseShape(item: any, baseProps: Partial<ParsedElement>): ParsedElement {
  const spPr = item['p:spPr'] || {};
  const xfrm = spPr['a:xfrm'];  // 正确提取 xfrm (transform) 对象

  return {
    ...baseProps,
    id: baseProps.id || generateElementId(),
    type: ElementType.SHAPE,
    shapeType: determineShapeType(item),
    position: parsePosition(xfrm),   // 传入 xfrm 而非 spPr
    size: parseSize(xfrm),           // 传入 xfrm 而非 spPr
    fill: parseFill(spPr),
    stroke: parseStroke(spPr),
    effects: parseEffects(spPr),
    textBox: parseTextBox(item['p:txBody']),
  };
}

/**
 * Parse image element
 */
function parseImage(item: any, baseProps: Partial<ParsedElement>): ParsedElement {
  const spPr = item['p:spPr'] || {};
  const xfrm = spPr['a:xfrm'];  // 正确提取 xfrm (transform) 对象
  // 图片的 blipFill 在 p:pic 元素下，不在 spPr 中
  const blipFill = item['p:blipFill'] || spPr['a:blipFill'];

  return {
    ...baseProps,
    id: baseProps.id || generateElementId(),
    type: ElementType.IMAGE,
    position: parsePosition(xfrm),   // 传入 xfrm 而非 spPr
    size: parseSize(xfrm),           // 传入 xfrm 而非 spPr
    imageRef: parseImageRef(blipFill),
    crop: parseCrop(blipFill),
    effects: parseEffects(spPr),
  };
}

/**
 * Parse chart or table element
 */
function parseChartOrTable(
  item: any,
  baseProps: Partial<ParsedElement>
): ParsedElement | null {
  const graphic = item['a:graphic'];
  const graphicData = graphic?.['a:graphicData'];
  const spPr = item['p:xfrm'] || item['p:spPr'] || {};
  const xfrm = spPr['a:xfrm'];

  if (!graphicData) {
    return null;
  }

  // Check if it's a chart
  if (graphicData['c:chart']) {
    return {
      ...baseProps,
      id: baseProps.id || generateElementId(),
      type: ElementType.CHART,
      chartRef: extractChartRef(item),
      position: parsePosition(xfrm),
      size: parseSize(xfrm),
    };
  }

  // Check if it's a table
  if (graphicData['a:tbl']) {
    return parseTable(graphicData['a:tbl'], baseProps);
  }

  return null;
}

/**
 * Parse table element
 */
function parseTable(tbl: any, baseProps: Partial<ParsedElement>): ParsedElement {
  const rows = Array.isArray(tbl['a:tr']) ? tbl['a:tr'] : [tbl['a:tr']].filter(Boolean);

  return {
    ...baseProps,
    id: baseProps.id || generateElementId(),
    type: ElementType.TABLE,
    rowCount: rows.length,
    colCount: tbl['a:tblPr']?.['#numCols'] || 0,
    rows: rows.map((tr: any) => parseTableRow(tr)),
    style: parseTableStyle(tbl['a:tblPr']),
  };
}

/**
 * Parse table row
 */
function parseTableRow(tr: any): any {
  const cells = Array.isArray(tr['a:tc']) ? tr['a:tc'] : [tr['a:tc']].filter(Boolean);
  return {
    cells: cells.map((tc: any) => parseTableCell(tc)),
    height: tr['a:trPr']?.['h'],
  };
}

/**
 * Parse table cell
 */
function parseTableCell(tc: any): any {
  return {
    text: extractTextFromTextBody(tc['a:txBody']),
    rowSpan: tc['a:tcPr']?.['rowSpan'] || 1,
    colSpan: tc['a:tcPr']?.['gridSpan'] || 1,
    fill: parseFill(tc['a:tcPr']),
  };
}

/**
 * Parse table style
 */
function parseTableStyle(tblPr: any): any {
  if (!tblPr) return null;

  return {
    bandCol: tblPr['bandCol'],
    bandRow: tblPr['bandRow'],
    firstCol: tblPr['firstCol'],
    firstRow: tblPr['firstRow'],
    lastCol: tblPr['lastCol'],
    lastRow: tblPr['lastRow'],
  };
}

/**
 * Parse line element
 */
function parseLine(item: any, baseProps: Partial<ParsedElement>): ParsedElement {
  const spPr = item['p:spPr'] || {};
  const cxnSpPr = item['p:cxnSpPr'] || {};
  const xfrm = spPr['a:xfrm'];  // 正确提取 xfrm (transform) 对象

  return {
    ...baseProps,
    id: baseProps.id || generateElementId(),
    type: ElementType.LINE,
    position: parsePosition(xfrm),   // 传入 xfrm 而非 spPr
    size: parseSize(xfrm),           // 传入 xfrm 而非 spPr
    startX: parsePoint(xfrm?.['a:off'])?.x || 0,
    startY: parsePoint(xfrm?.['a:off'])?.y || 0,
    endX: parsePoint(cxnSpPr['a:endPos'])?.x || 0,
    endY: parsePoint(cxnSpPr['a:endPos'])?.y || 0,
    stroke: parseStroke(spPr),
    style: parseLineStyle(item),
  };
}

/**
 * Parse group element
 */
function parseGroup(
  item: any,
  baseProps: Partial<ParsedElement>,
  options: ParserOptions
): ParsedElement {
  const elements: ParsedElement[] = [];
  const spPr = item['p:spPr'] || {};
  const xfrm = spPr['a:xfrm'];  // 正确提取 xfrm

  // Parse child elements
  for (const [tag, type] of [
    ['p:sp', ElementType.SHAPE],
    ['p:pic', ElementType.IMAGE],
    ['p:graphicFrame', ElementType.CHART],
    ['p:cxnSp', ElementType.LINE],
  ] as const) {
    const items = Array.isArray(item[tag]) ? item[tag] : [item[tag]].filter(Boolean);

    for (const child of items) {
      const element = parseElement(child, type as ElementType, options);
      if (element) {
        elements.push(element);
      }
    }
  }

  return {
    ...baseProps,
    id: baseProps.id || generateElementId(),
    type: ElementType.GROUP,
    elements,
    position: parsePosition(xfrm),
    size: parseSize(xfrm),  // 添加 size 属性
  };
}

/**
 * Parse background
 */
function parseBackground(spTree: any): SlideBackground | undefined {
  const bg = spTree?.['p:bg'] || spTree?.['background'];

  if (!bg) return undefined;

  const bgPr = bg['p:bgPr'];
  const bgRef = bg['p:bgRef'];

  if (bgPr?.['a:solidFill']) {
    const colorResult = parseColor(bgPr['a:solidFill']);
    let colorStr: string | undefined;
    if (typeof colorResult === 'object' && colorResult !== null) {
      colorStr = colorResult.color;
    } else if (typeof colorResult === 'string') {
      colorStr = colorResult;
    } else {
      colorStr = undefined;
    }
    const result: SlideBackground = { type: 'solid', color: colorStr };
    return result;
  }

  if (bgPr?.['a:gradFill']) {
    return {
      type: 'gradient',
      gradientColors: parseGradientColors(bgPr['a:gradFill']),
    };
  }

  if (bgRef) {
    return {
      type: 'image',
      imageRef: bgRef['r:embed'],
    };
  }

  return undefined;
}

/**
 * Parse transition
 */
function parseTransition(transition: any): SlideTransition {
  return {
    type: Object.keys(transition).find(key => key.startsWith('p:'))?.replace('p:', ''),
    duration: transition['duration'] || transition['advTm'],
    direction: transition['dir'],
  };
}

/**
 * Parse position from transform
 */
function parsePosition(xfrm: any): { x: number; y: number } | undefined {
  if (!xfrm?.['a:off']) return undefined;

  return {
    x: parseInt(xfrm['a:off']['x'], 10) || 0,
    y: parseInt(xfrm['a:off']['y'], 10) || 0,
  };
}

/**
 * Parse size from transform
 */
function parseSize(xfrm: any): { width: number; height: number } | undefined {
  if (!xfrm?.['a:ext']) return undefined;

  return {
    width: parseInt(xfrm['a:ext']['cx'], 10) || 0,
    height: parseInt(xfrm['a:ext']['cy'], 10) || 0,
  };
}

/**
 * Parse point from EMU coordinates
 */
function parsePoint(point: any): { x: number; y: number } | undefined {
  if (!point) return undefined;

  return {
    x: parseInt(point['x'], 10) || 0,
    y: parseInt(point['y'], 10) || 0,
  };
}

/**
 * Parse fill
 * Handles solid fill, gradient fill, image fill, and noFill (transparent)
 */
function parseFill(fill: any): any {
  if (!fill) return undefined;

  // Handle noFill (transparent background)
  if (fill['a:noFill']) {
    return undefined;
  }

  if (fill['a:solidFill']) {
    const colorResult = parseColor(fill['a:solidFill']);
    const result: any = { type: 'solid' };

    if (typeof colorResult === 'object' && colorResult !== null) {
      result.color = colorResult.color;
      if (colorResult.opacity !== undefined) {
        result.opacity = colorResult.opacity;
      }
    } else if (colorResult) {
      result.color = colorResult;
    } else {
      return undefined; // No valid color
    }

    // Fully transparent should return undefined
    if (result.opacity === 0) return undefined;

    return result;
  }

  if (fill['a:gradFill']) {
    return {
      type: 'gradient',
      colors: parseGradientColors(fill['a:gradFill']),
    };
  }

  if (fill['a:blipFill']) {
    return {
      type: 'image',
      imageRef: fill['a:blipFill']['a:blip']?.['r:embed'],
    };
  }

  return undefined;
}

/**
 * Parsed color result with optional opacity
 */
interface ParsedColorResult {
  color: string;
  opacity?: number;
}

/**
 * Parse color with optional alpha transparency
 *
 * PPTX alpha values range from 0-100000, representing 0%-100% opacity.
 * We convert this to a 0-1 range for web use.
 */
function parseColor(colorFill: any): string | ParsedColorResult {
  if (!colorFill) return '';

  const srgbClr = colorFill['a:srgbClr'];
  const scrgbClr = colorFill['a:scrgbClr'];
  const sysClr = colorFill['a:sysClr'];
  const schemeClr = colorFill['a:schemeClr'];

  let color = '';
  let opacity: number | undefined;

  if (srgbClr?.['val']) {
    // sRGB color
    const val = srgbClr['val'];
    color = val.startsWith('#') ? val : `#${val}`;
    // Extract alpha (0-100000 → 0-1 opacity)
    const alpha = srgbClr['a:alpha']?.['val'];
    if (alpha !== undefined) {
      opacity = 1 - (parseInt(alpha, 10) / 100000);
    }
  } else if (scrgbClr) {
    // ScRGB color (RGB in percentage)
    const r = Math.round((parseInt(scrgbClr['r'], 10) / 100000) * 255);
    const g = Math.round((parseInt(scrgbClr['g'], 10) / 100000) * 255);
    const b = Math.round((parseInt(scrgbClr['b'], 10) / 100000) * 255);
    color = `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
    const alpha = scrgbClr['a:alpha']?.['val'];
    if (alpha !== undefined) {
      opacity = 1 - (parseInt(alpha, 10) / 100000);
    }
  } else if (sysClr?.['lastClr']) {
    // System color with last known value
    const val = sysClr['lastClr'];
    color = val.startsWith('#') ? val : `#${val}`;
    const alpha = sysClr['a:alpha']?.['val'];
    if (alpha !== undefined) {
      opacity = 1 - (parseInt(alpha, 10) / 100000);
    }
  } else if (schemeClr?.['val']) {
    // Scheme/theme color - use val as-is for now (should be resolved via theme)
    // Common values: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
    // For now, map common scheme colors to approximations
    const schemeMap: Record<string, string> = {
      dk1: '#000000',      // Dark 1 (usually black)
      lt1: '#ffffff',      // Light 1 (usually white)
      dk2: '#1f497d',      // Dark 2
      lt2: '#eeece1',      // Light 2
      accent1: '#4472c4',  // Accent 1
      accent2: '#ed7d31',  // Accent 2
      accent3: '#a5a5a5',  // Accent 3
      accent4: '#ffc000',  // Accent 4
      accent5: '#5b9bd5',  // Accent 5
      accent6: '#70ad47',  // Accent 6
      hlink: '#0563c1',    // Hyperlink
      folHlink: '#954f72', // Followed hyperlink
    };
    const schemeVal = schemeClr['val'];
    color = schemeMap[schemeVal] || '#000000';
    const alpha = schemeClr['a:alpha']?.['val'];
    if (alpha !== undefined) {
      opacity = 1 - (parseInt(alpha, 10) / 100000);
    }
  }

  // Return object with opacity if transparency is present
  if (opacity !== undefined && opacity < 1) {
    return { color, opacity };
  }

  return color;
}

/**
 * Parse gradient colors
 */
function parseGradientColors(gradFill: any): string[] {
  const colors: string[] = [];
  const gsLst = gradFill['a:gsLst']?.['a:gs'];

  if (!gsLst) return colors;

  const stops = Array.isArray(gsLst) ? gsLst : [gsLst];

  for (const stop of stops) {
    const colorResult = parseColor(stop);
    // Extract color string from result (handle both string and object return)
    let color: string | undefined;
    if (typeof colorResult === 'object' && colorResult !== null) {
      color = colorResult.color;
    } else if (typeof colorResult === 'string') {
      color = colorResult;
    }
    if (color) colors.push(color);
  }

  return colors;
}

/**
 * Parse stroke
 */
function parseStroke(spPr: any): any {
  const ln = spPr?.['a:ln'];

  if (!ln) return undefined;

  const colorResult = parseColor(ln['a:solidFill']);
  const color = typeof colorResult === 'object' && colorResult !== null
    ? colorResult.color
    : colorResult;

  return {
    width: ln['w'] ? parseInt(ln['w'], 10) / 12700 : 1, // EMU to points
    color,
    dashType: ln['a:prstDash']?.['val'],
    lineCap: ln['a:cap']?.['val'],
    lineJoin: ln['a:join']?.['a:round']?.['lim'],
  };
}

/**
 * Parse effects
 */
function parseEffects(spPr: any): any {
  const effectLst = spPr?.['a:effectLst'];

  if (!effectLst) return undefined;

  return {
    shadow: parseShadow(effectLst['a:outerShdw'] || effectLst['a:innerShdw']),
    reflection: effectLst['a:reflection'],
    glow: parseGlow(effectLst['a:glow']),
    blur: effectLst['a:blur'],
  };
}

/**
 * Parse shadow
 */
function parseShadow(shadow: any): any {
  if (!shadow) return undefined;

  const colorResult = parseColor(shadow);
  const color = typeof colorResult === 'object' && colorResult !== null
    ? colorResult.color
    : colorResult;

  return {
    type: shadow === 'a:outerShdw' ? 'outer' : 'inner',
    color,
    blur: shadow['blurRad'] ? parseInt(shadow['blurRad'], 10) / 12700 : undefined,
    offset: shadow['dist'] ? parseInt(shadow['dist'], 10) / 12700 : undefined,
    angle: shadow['dir'] ? parseInt(shadow['dir'], 10) / 60000 : undefined,
    opacity: shadow['opacity'] ? parseInt(shadow['opacity'], 10) / 100000 : undefined,
  };
}

/**
 * Parse glow
 */
function parseGlow(glow: any): any {
  if (!glow) return undefined;

  const colorResult = parseColor(glow);
  const color = typeof colorResult === 'object' && colorResult !== null
    ? colorResult.color
    : colorResult;

  return {
    color,
    radius: glow['rad'] ? parseInt(glow['rad'], 10) / 12700 : undefined,
  };
}

/**
 * Parse text box
 */
function parseTextBox(txBody: any): any {
  if (!txBody) return undefined;

  const paragraphs = Array.isArray(txBody['a:p'])
    ? txBody['a:p']
    : [txBody['a:p']].filter(Boolean);

  return {
    paragraphs: paragraphs.map((p: any) => parseParagraph(p)),
    verticalAlign: txBody['a:bodyPr']?.['anchor'],
  };
}

/**
 * Parse paragraph
 */
function parseParagraph(p: any): any {
  const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']].filter(Boolean);
  const parsedRuns = runs.map((r: any) => parseTextRun(r));

  return {
    text: parsedRuns.map((r: any) => r.text || '').join(''),
    runs: parsedRuns,
    alignment: p['a:pPr']?.['algn'],
    indent: p['a:pPr']?.['indent'],
    spacing: parseSpacing(p['a:pPr']),
  };
}

/**
 * Parse text run
 */
function parseTextRun(r: any): any {
  const rPr = r['a:rPr'];

  // Handle both string and object forms of a:t
  const textNode = r['a:t'];
  const text = typeof textNode === 'string' ? textNode : (textNode?.['#text'] || '');

  // Parse color with opacity support
  const colorResult = parseColor(rPr?.['a:solidFill']);
  const color = typeof colorResult === 'object' && colorResult !== null
    ? colorResult.color
    : colorResult;

  return {
    text,
    font: rPr?.['a:latin']?.['typeface'] || rPr?.['a:cs']?.['typeface'],
    size: rPr?.['sz'] ? parseInt(rPr['sz'], 10) / 100 : undefined, // in points
    bold: rPr?.['b'] === '1' || rPr?.['b'] === true,
    italic: rPr?.['i'] === '1' || rPr?.['i'] === true,
    underline: rPr?.['u'] === '1' || rPr?.['u'] === true,
    color,
  };
}

/**
 * Parse spacing
 */
function parseSpacing(pPr: any): any {
  if (!pPr?.['a:lnSpc'] && !pPr?.['a:spcBef'] && !pPr?.['a:spcAft']) {
    return undefined;
  }

  return {
    line: pPr?.['a:lnSpc']?.['a:spcPct']?.['val']
      ? parseInt(pPr['a:lnSpc']['a:spcPct']['val'], 10) / 1000
      : undefined,
    before: pPr?.['a:spcBef']?.['a:spcPct']?.['val']
      ? parseInt(pPr['a:spcBef']['a:spcPct']['val'], 10) / 1000
      : undefined,
    after: pPr?.['a:spcAft']?.['a:spcPct']?.['val']
      ? parseInt(pPr['a:spcAft']['a:spcPct']['val'], 10) / 1000
      : undefined,
  };
}

/**
 * Extract text from text body
 */
function extractTextFromTextBody(txBody: any): string {
  if (!txBody) return '';

  const paragraphs = Array.isArray(txBody['a:p'])
    ? txBody['a:p']
    : [txBody['a:p']].filter(Boolean);

  return paragraphs
    .map((p: any) => {
      const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']].filter(Boolean);
      return runs.map((r: any) => r['a:t']?.['#text'] || '').join('');
    })
    .join('\n');
}

/**
 * Parse image reference from blip fill
 */
function parseImageRef(blipFill: any): string | undefined {
  return blipFill?.['a:blip']?.['r:embed'];
}

/**
 * Parse crop from blip fill
 */
function parseCrop(blipFill: any): any {
  if (!blipFill?.['a:srcRect']) return undefined;

  const srcRect = blipFill['a:srcRect'];

  return {
    top: srcRect['t'] ? parseInt(srcRect['t'], 10) / 100000 : undefined,
    right: srcRect['r'] ? parseInt(srcRect['r'], 10) / 100000 : undefined,
    bottom: srcRect['b'] ? parseInt(srcRect['b'], 10) / 100000 : undefined,
    left: srcRect['l'] ? parseInt(srcRect['l'], 10) / 100000 : undefined,
  };
}

/**
 * Extract chart reference
 */
function extractChartRef(item: any): string | undefined {
  return item['p:nvGraphicFramePr']?.['p:cNvPr']?.['a:hlinkClick']?.['r:id'];
}

/**
 * Determine shape type
 */
function determineShapeType(item: any): string {
  const nvSpPr = item['p:nvSpPr'];
  const cNvPr = nvSpPr?.['p:cNvPr'];

  // Check for placeholder types
  if (cNvPr?.['type']) {
    return cNvPr['type'];
  }

  // Default
  return 'rectangle';
}

/**
 * Parse line style
 */
function parseLineStyle(item: any): string | undefined {
  const spPr = item['p:spPr'];
  const ln = spPr?.['a:ln'];

  if (!ln) return undefined;

  if (ln['a:prstDash']) return 'dashed';
  if (ln['a:round']) return 'round';
  if (ln['a:bevel']) return 'bevel';

  return 'solid';
}

/**
 * Generate unique element ID
 */
function generateLocalElementId(): string {
  return `el-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Parse all slides from extracted PPTX
 *
 * @param extracted - Extracted PPTX structure
 * @param options - Parser options
 * @returns Array of parsed slides
 */
export function parseAllSlides(
  extracted: ExtractedPPTX,
  options: ParserOptions = {}
): ParsedSlide[] {
  const slides: ParsedSlide[] = [];

  // Build relationship map for media resolution
  const relationshipMap = buildRelationshipMap(extracted);

  for (const [index, xml] of extracted.slides.entries()) {
    try {
      const slide = parseSlideXML(xml, index + 1, options);
      // Attach relationship map to each slide for media resolution
      (slide as any)._relationshipMap = relationshipMap.get(index);
      slides.push(slide);
    } catch (error) {
      logger.error(`Failed to parse slide ${index + 1}`, {
        error: error instanceof Error ? error.message : String(error),
      });
      // Continue parsing other slides
    }
  }

  logger.info('Parsed all slides', {
    totalSlides: slides.length,
    errors: extracted.slides.size - slides.length,
  });

  return slides;
}

/**
 * Build relationship map from extracted PPTX relationships
 * Maps rId (e.g., "rId1") to target file path (e.g., "../media/image1.png")
 */
function buildRelationshipMap(extracted: ExtractedPPTX): Map<number, Map<string, string>> {
  const parser = createXMLParser({});
  const slideRelationships = new Map<number, Map<string, string>>();

  for (const [slideIndex, relXml] of extracted.relationships.entries()) {
    const relMap = new Map<string, string>();

    try {
      const parsed = parser.parse(relXml);
      const relationships = parsed['Relationships']?.['Relationship'] || [];

      const rels = Array.isArray(relationships) ? relationships : [relationships];

      for (const rel of rels) {
        if (rel && rel['Id'] && rel['Target']) {
          relMap.set(rel['Id'], rel['Target']);
        }
      }

      slideRelationships.set(slideIndex, relMap);
    } catch (error) {
      logger.warn(`Failed to parse relationships for slide ${slideIndex}`, {
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  return slideRelationships;
}

/**
 * Resolve media reference from relationship ID to actual filename
 *
 * @param relId - Relationship ID (e.g., "rId1")
 * @param relationshipMap - Map of rId to target path
 * @returns Media filename or undefined
 */
export function resolveMediaFilename(relId: string, relationshipMap?: Map<string, string>): string | undefined {
  if (!relId || !relationshipMap) return undefined;

  const target = relationshipMap.get(relId);
  if (!target) return undefined;

  // Extract filename from target path (e.g., "../media/image1.png" -> "image1.png")
  const parts = target.split('/');
  return parts[parts.length - 1];
}
