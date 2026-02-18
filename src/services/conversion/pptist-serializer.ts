/**
 * PPTist Format Serializer
 *
 * Converts internal presentation data to PPTist-compatible JSON format.
 * This format matches the internal element types (PPTShapeElement, PPTTextElement, etc.)
 * used by the PPTist frontend.
 *
 * @module services/conversion/pptist-serializer
 */

import type { Presentation, Slide } from '../../types/presentations';
import type { ConversionMetadata, ConversionWarning } from '../../models';
import { normalizeColor } from '../../utils/color';

/**
 * Default theme values
 */
const DEFAULT_THEME = {
  width: 1280,
  height: 720,
  fontName: '微软雅黑',
  fontColor: '#000000',
};

/**
 * Validate and normalize coordinate value
 */
function validateCoordinate(value: number | undefined, defaultValue: number = 0): number {
  if (value === undefined || value === null || isNaN(value)) {
    return defaultValue;
  }
  // Ensure non-negative for width/height
  return Math.max(0, Math.round(value));
}

/**
 * Validate and normalize size value
 */
function validateSize(value: number | undefined, defaultValue: number = 100): number {
  if (value === undefined || value === null || isNaN(value) || value <= 0) {
    return defaultValue;
  }
  return Math.round(value);
}

/**
 * Convert shape element to PPTist format
 */
function convertShapeElement(element: any): any {
  const {
    id,
    x = 0,
    y = 0,
    width = 100,
    height = 100,
    rotate = 0,
    locked = false,
    visible = true,
    fill,
    outline,
    text,
    shape,
  } = element;

  // Validate coordinates
  const left = validateCoordinate(x);
  const top = validateCoordinate(y);
  const w = validateSize(width, 100);
  const h = validateSize(height, 100);

  // Build basic shape element
  const pptistShape: any = {
    type: 'shape',
    id: id || `shape_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
    left,
    top,
    width: w,
    height: h,
    rotate,
    lock: locked || undefined,
    viewBox: [200, 200],
    path: 'M 0 0 L 200 0 L 200 200 L 0 200 Z', // Default rectangle path
    fixedRatio: false,
    fill: extractFillColor(fill) || '#ffffff',
    outline: outline ? {
      width: outline.width || 1,
      style: outline.style || 'solid',
      color: normalizeColor(outline.color) || '#000000',
    } : undefined,
  };

  // Add text if present
  if (text && typeof text === 'string') {
    pptistShape.text = {
      content: escapeHtml(text),
      defaultFontName: DEFAULT_THEME.fontName,
      defaultColor: DEFAULT_THEME.fontColor,
      align: 'middle',
    };
  }

  return pptistShape;
}

/**
 * Convert text element to PPTist format
 */
function convertTextElement(element: any): any {
  const {
    id,
    x = 0,
    y = 0,
    width = 100,
    height = 50,
    rotate = 0,
    locked = false,
    visible = true,
    text,
    content,
    fill,
    outline,
    defaultValue,
  } = element;

  // Validate coordinates
  const left = validateCoordinate(x);
  const top = validateCoordinate(y);
  const w = validateSize(width, 100);
  const h = validateSize(height, 50);

  return {
    type: 'text',
    id: id || `text_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
    left,
    top,
    width: w,
    height: h,
    rotate,
    lock: locked || undefined,
    content: content || (text ? escapeHtml(text) : '') || defaultValue || '',
    defaultFontName: DEFAULT_THEME.fontName,
    defaultColor: DEFAULT_THEME.fontColor,
    outline: outline ? {
      width: outline.width || 1,
      style: outline.style || 'solid',
      color: normalizeColor(outline.color) || '#000000',
    } : undefined,
    fill: extractFillColor(fill) || undefined,
    lineHeight: 1.5,
  };
}

/**
 * Convert image element to PPTist format
 */
function convertImageElement(element: any): any {
  const {
    id,
    x = 0,
    y = 0,
    width = 100,
    height = 100,
    rotate = 0,
    locked = false,
    visible = true,
    src = '',
  } = element;

  // Validate coordinates
  const left = validateCoordinate(x);
  const top = validateCoordinate(y);
  const w = validateSize(width, 100);
  const h = validateSize(height, 100);

  return {
    type: 'image',
    id: id || `image_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
    left,
    top,
    width: w,
    height: h,
    rotate,
    lock: locked || undefined,
    fixedRatio: true,
    src,
  };
}

/**
 * Convert line element to PPTist format
 */
function convertLineElement(element: any): any {
  const {
    id,
    startX = 0,
    startY = 0,
    endX = 100,
    endY = 100,
    width = 2,
    color = '#000000',
    style = 'solid',
    x,
    y,
  } = element;

  // Use bounding box coordinates if available, otherwise calculate from start/end
  let left: number;
  let top: number;
  let w: number;
  let h: number;

  if (x !== undefined && y !== undefined && element.width !== undefined && element.height !== undefined) {
    // Use provided bounding box
    left = validateCoordinate(x);
    top = validateCoordinate(y);
    w = validateSize(element.width, 1);
    h = validateSize(element.height, 1);
  } else {
    // Calculate bounding box from start/end points
    left = Math.min(startX, endX);
    top = Math.min(startY, endY);
    w = Math.abs(endX - startX) || 1;
    h = Math.abs(endY - startY) || 1;
  }

  // Calculate relative start/end positions for PPTist
  // PPTist expects start/end as arrays relative to the bounding box
  const relativeStartX = startX - left;
  const relativeStartY = startY - top;
  const relativeEndX = endX - left;
  const relativeEndY = endY - top;

  return {
    type: 'line',
    id: id || `line_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
    left,
    top,
    width: w,
    height: h,
    // PPTist uses start/end arrays for relative coordinates
    start: [relativeStartX, relativeStartY],
    end: [relativeEndX, relativeEndY],
    style,
    color: normalizeColor(color),
    points: ['', ''], // Arrow head types
  };
}

/**
 * Extract fill color from fill object
 */
function extractFillColor(fill: any): string | undefined {
  if (!fill) return undefined;

  if (typeof fill === 'string') {
    return normalizeColor(fill);
  }

  if (fill.type === 'solid' && fill.color) {
    return normalizeColor(fill.color);
  }

  return undefined;
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text: string): string {
  const map: Record<string, string> = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;',
  };
  return text.replace(/[&<>"']/g, (m) => map[m]);
}

/**
 * Convert slide element to PPTist format
 */
function convertSlideElement(element: any): any | null {
  if (!element || !element.type) {
    return null;
  }

  switch (element.type) {
    case 'shape':
      return convertShapeElement(element);
    case 'text':
      return convertTextElement(element);
    case 'image':
      return convertImageElement(element);
    case 'line':
      return convertLineElement(element);
    default:
      console.warn(`Unknown element type: ${element.type}`);
      return null;
  }
}

/**
 * Convert slide background to PPTist format
 */
function convertBackground(background: any): any {
  if (!background) {
    return {
      type: 'solid',
      color: '#ffffff',
    };
  }

  if (background.type === 'solid' && background.color) {
    return {
      type: 'solid',
      color: normalizeColor(background.color),
    };
  }

  if (background.type === 'gradient') {
    return {
      type: 'gradient',
      gradient: {
        type: background.value?.path === 'line' ? 'linear' : 'radial',
        colors: (background.value?.colors || []).map((c: any) => ({
          ...c,
          pos: parseInt(c.pos) || 0,
          color: normalizeColor(c.color),
        })),
        rotate: (background.value?.rot || 0) + 90,
      },
    };
  }

  if (background.type === 'image') {
    return {
      type: 'image',
      image: {
        src: background.value?.picBase64 || '',
        size: 'cover',
      },
    };
  }

  return {
    type: 'solid',
    color: '#ffffff',
  };
}

/**
 * Convert slide to PPTist format
 */
function convertSlide(slide: Slide): any {
  const elements = (slide.elements || [])
    .map(convertSlideElement)
    .filter((el): el is any => el !== null);

  return {
    id: slide.id,
    elements,
    background: convertBackground(slide.background),
    remark: slide.notes || '',
  };
}

/**
 * Serialize presentation data to PPTist JSON format
 *
 * @param presentation - Presentation data
 * @param metadata - Conversion metadata
 * @param warnings - Conversion warnings
 * @returns PPTist-compatible JSON object
 */
export function serializeToPPTistFormat(
  presentation: Presentation,
  metadata: ConversionMetadata,
  warnings?: ConversionWarning[]
): any {
  // Convert slides
  const slides = (presentation.slides || []).map(convertSlide);

  // Build PPTist format
  const pptistData = {
    slides,
    theme: {
      width: presentation.width || DEFAULT_THEME.width,
      height: presentation.height || DEFAULT_THEME.height,
      fontName: DEFAULT_THEME.fontName,
      fontColor: DEFAULT_THEME.fontColor,
    },
  };

  return pptistData;
}
