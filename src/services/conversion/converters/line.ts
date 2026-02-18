/**
 * Line Element Converter
 *
 * Converts PPTX line elements to PPTist line elements.
 * Handles straight lines, arrows, and connector lines.
 *
 * @module services/conversion/converters/line
 */

import { BaseElementConverter } from '../base-converter';
import { ElementType, type ConversionContext } from '../../../types/converters';
import type { LineElement } from '../../../types/pptist';
import type { ParsedElement } from '../../../services/pptx/parser';
import { emuToPixels } from '../../../utils/coordinates';
import { normalizeColor } from '../../../utils/color';
import { generateTrackedId, ID_PREFIXES } from '../../../utils/id-generator';
import { logger } from '../../../utils/logger';

/**
 * Line converter implementation
 */
export class LineConverter extends BaseElementConverter<ParsedElement, LineElement> {
  readonly type = ElementType.LINE;
  readonly priority = 100;
  readonly supportedVersions = ['1.0.0', 'latest'];

  /**
   * Check if this converter can handle the element
   */
  async canConvert(element: ParsedElement): Promise<boolean> {
    return element.type === 'line';
  }

  /**
   * Convert PPTX line element to PPTist line element
   */
  async convert(element: ParsedElement, context: ConversionContext): Promise<LineElement | null> {
    logger.debug('Converting line element', {
      id: element.id,
      startX: element.startX,
      startY: element.startY,
      endX: element.endX,
      endY: element.endY,
    });

    // Convert EMU coordinates to pixels
    const startXPx = emuToPixels(element.startX || 0);
    const startYPx = emuToPixels(element.startY || 0);
    const endXPx = emuToPixels(element.endX || 914400); // Default 1 inch
    const endYPx = emuToPixels(element.endY || 0);

    // Calculate bounding box
    const left = Math.min(startXPx, endXPx);
    const top = Math.min(startYPx, endYPx);
    const width = Math.abs(endXPx - startXPx);
    const height = Math.abs(endYPx - startYPx);

    // Calculate relative start/end positions (relative to bounding box)
    const relativeStartX = startXPx - left;
    const relativeStartY = startYPx - top;
    const relativeEndX = endXPx - left;
    const relativeEndY = endYPx - top;

    const pptistElement: LineElement = {
      id: generateTrackedId(ID_PREFIXES.LINE, undefined, element.id),
      type: 'line',
      // Position and size (bounding box)
      x: left,
      y: top,
      width: width || 1, // Minimum 1px for zero-length lines
      height: height || 1,
      // Layer order
      zIndex: element.zIndex,
      // Line-specific properties
      startX: startXPx,
      startY: startYPx,
      endX: endXPx,
      endY: endYPx,
      // Line styling
      style: this.convertLineStyle(element.style),
      color: normalizeColor(element.stroke?.color) || '#000000',
    };

    // Convert stroke properties
    if (element.stroke) {
      this.convertStroke(element.stroke, pptistElement);
    }

    // Convert effects
    if (element.effects) {
      this.convertEffects(element.effects, pptistElement);
    }

    return pptistElement;
  }

  /**
   * Convert line style
   */
  private convertLineStyle(style?: string): 'solid' | 'dashed' {
    if (!style) return 'solid';

    const styleMap: Record<string, 'solid' | 'dashed'> = {
      solid: 'solid',
      dash: 'dashed',
      dashed: 'dashed',
      dot: 'dashed',
      dotted: 'dashed',
      dashDot: 'dashed',
      lgDash: 'dashed',
      sysDash: 'dashed',
    };

    return styleMap[style] || 'solid';
  }

  /**
   * Convert stroke properties
   */
  private convertStroke(stroke: any, pptistElement: LineElement): void {
    if (stroke.width !== undefined) {
      // Line width is typically in EMU, convert to points then pixels
      // But stroke.width from parser is already in points (see parseStroke in parser.ts)
      pptistElement.width = stroke.width;
    }

    if (stroke.color) {
      pptistElement.color = normalizeColor(stroke.color);
    }

    if (stroke.lineCap) {
      pptistElement.cap = this.convertLineCap(stroke.lineCap);
    }
  }

  /**
   * Convert line cap
   */
  private convertLineCap(cap: string): 'round' | 'butt' | 'square' {
    const capMap: Record<string, 'round' | 'butt' | 'square'> = {
      round: 'round',
      flat: 'butt',
      square: 'square',
    };

    return capMap[cap] || 'butt';
  }

  /**
   * Convert effects
   */
  private convertEffects(effects: any, pptistElement: LineElement): void {
    if (effects.shadow) {
      pptistElement.shadow = this.convertShadow(effects.shadow);
    }

    if (effects.glow) {
      pptistElement.glow = this.convertGlow(effects.glow);
    }
  }

  /**
   * Convert shadow effect
   */
  private convertShadow(shadow: any): any {
    const pptistShadow: any = {};

    if (shadow.color) pptistShadow.color = normalizeColor(shadow.color);
    if (shadow.offset !== undefined) pptistShadow.offset = shadow.offset;
    if (shadow.blur !== undefined) pptistShadow.blur = shadow.blur;
    if (shadow.angle !== undefined) pptistShadow.angle = shadow.angle;
    if (shadow.opacity !== undefined) pptistShadow.opacity = shadow.opacity;

    return pptistShadow;
  }

  /**
   * Convert glow effect
   */
  private convertGlow(glow: any): any {
    const pptistGlow: any = {};

    if (glow.color) pptistGlow.color = normalizeColor(glow.color);
    if (glow.radius !== undefined) pptistGlow.radius = glow.radius;

    return pptistGlow;
  }
}

/**
 * Create line converter instance
 */
export function createLineConverter(): LineConverter {
  return new LineConverter();
}
