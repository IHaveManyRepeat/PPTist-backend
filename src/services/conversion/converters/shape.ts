/**
 * Shape Element Converter
 *
 * Converts PPTX shape elements to PPTist shape elements.
 * Handles rectangles, circles, polygons, and custom shapes.
 *
 * @module services/conversion/converters/shape
 */

import { BaseElementConverter } from '../base-converter';
import { ElementType, type ConversionContext } from '../../../types/converters';
import type { ShapeElement } from '../../../types/pptist';
import type { ParsedElement } from '../../../services/pptx/parser';
import { emuToPixels } from '../../../utils/coordinates';
import { normalizeColor } from '../../../utils/color';
import { generateTrackedId, ID_PREFIXES } from '../../../utils/id-generator';
import { logger } from '../../../utils/logger';

/**
 * Shape converter implementation
 */
export class ShapeConverter extends BaseElementConverter<ParsedElement, ShapeElement> {
  readonly type = ElementType.SHAPE;
  readonly priority = 100;
  readonly supportedVersions = ['1.0.0', 'latest'];

  /**
   * Check if this converter can handle the element
   */
  async canConvert(element: ParsedElement): Promise<boolean> {
    return element.type === 'shape';
  }

  /**
   * Convert PPTX shape element to PPTist shape element
   */
  async convert(element: ParsedElement, context: ConversionContext): Promise<ShapeElement | null> {
    logger.debug('Converting shape element', {
      id: element.id,
      shapeType: element.shapeType,
    });

    // Convert EMU to pixels (apply 96/72 scaling)
    const x = emuToPixels(element.position?.x || 0);
    const y = emuToPixels(element.position?.y || 0);
    const width = emuToPixels(element.size?.width || 914400); // Default 1 inch
    const height = emuToPixels(element.size?.height || 914400); // Default 1 inch

    const pptistElement: ShapeElement = {
      id: generateTrackedId(ID_PREFIXES.SHAPE, undefined, element.id),
      type: 'shape',
      x,
      y,
      width,
      height,
      rotate: element.rotation || 0,
      locked: element.locked || false,
      visible: !element.hidden,
      zIndex: element.zIndex,
      fill: this.convertFill(element.fill),
      outline: this.convertOutline(element.stroke),
      text: this.convertTextContent(element),
      shape: this.convertShapeType(element.shapeType),
    };

    // Convert effects
    if (element.effects) {
      this.convertEffects(element.effects, pptistElement);
    }

    // Convert text box (if shape contains text)
    if (element.textBox) {
      this.convertTextBox(element.textBox, pptistElement, context);
    }

    return pptistElement;
  }

  /**
   * Convert fill
   */
  private convertFill(fill: any): any {
    if (!fill) return undefined;

    if (fill.type === 'solid' && fill.color) {
      const result: any = {
        type: 'solid',
        color: normalizeColor(fill.color),
      };
      // 传递透明度
      if (fill.opacity !== undefined) {
        result.opacity = fill.opacity;
      }
      return result;
    }

    if (fill.type === 'gradient' && fill.colors) {
      return {
        type: 'gradient',
        gradientColors: fill.colors.map((c: string) => normalizeColor(c)),
        gradientAngle: 90,
        gradientPosition: 0,
      };
    }

    if (fill.type === 'image' && fill.imageRef) {
      return {
        type: 'image',
        image: fill.imageRef,
      };
    }

    return undefined;
  }

  /**
   * Convert outline (stroke)
   */
  private convertOutline(stroke: any): any {
    if (!stroke) return undefined;

    const pptistOutline: any = {};

    if (stroke.width !== undefined) pptistOutline.width = stroke.width;
    if (stroke.color) pptistOutline.color = normalizeColor(stroke.color);
    if (stroke.dashType) pptistOutline.style = this.convertDashStyle(stroke.dashType);
    if (stroke.lineCap) pptistOutline.lineCap = stroke.lineCap;
    if (stroke.lineJoin) pptistOutline.lineJoin = stroke.lineJoin;

    return pptistOutline;
  }

  /**
   * Convert dash style
   */
  private convertDashStyle(dashType: string): string {
    const dashMap: Record<string, string> = {
      solid: 'solid',
      dash: 'dashed',
      dashDot: 'dash-dot',
      dot: 'dotted',
    };

    return dashMap[dashType] || 'solid';
  }

  /**
   * Convert text content
   */
  private convertTextContent(element: ParsedElement): string {
    if (!element.textBox) return '';

    const paragraphs = element.textBox.paragraphs || [];
    return paragraphs.map((p: any) => p.text || '').join('\n');
  }

  /**
   * Convert shape type
   */
  private convertShapeType(pptxShapeType?: string): any {
    // Default rectangle
    const defaultShape = {
      type: 'rect',
      radius: 0,
    };

    if (!pptxShapeType) return defaultShape;

    // Map PPTX shape types to PPTist shape types
    const shapeTypeMap: Record<string, any> = {
      // Rectangles
      rect: { type: 'rect', radius: 0 },
      roundRect: { type: 'rect', radius: 0.2 },
      snipRoundRect: { type: 'rect', radius: 0.1 },
      // Circles and ellipses
      ellipse: { type: 'ellipse' },
      oval: { type: 'ellipse' },
      circle: { type: 'ellipse' },
      // Triangles
      triangle: { type: 'path', path: 'M0,1 L0.5,0 L1,1 Z' },
      rtTriangle: { type: 'path', path: 'M0,0 L1,0 L0,1 Z' },
      // Polygons
      diamond: { type: 'path', path: 'M0.5,0 L1,0.5 L0.5,1 L0,0.5 Z' },
      pentagon: { type: 'path', path: 'M0.5,0 L1,0.38 L0.82,1 L0.18,1 L0,0.38 Z' },
      hexagon: { type: 'path', path: 'M0.25,0 L0.75,0 L1,0.5 L0.75,1 L0.25,1 L0,0.5 Z' },
      heptagon: { type: 'path', path: 'M0.5,0 L0.9,0.25 L1,0.65 L0.8,1 L0.2,1 L0,0.65 L0.1,0.25 Z' },
      octagon: { type: 'path', path: 'M0.3,0 L0.7,0 L1,0.3 L1,0.7 L0.7,1 L0.3,1 L0,0.7 L0,0.3 Z' },
      // Stars
      star5: { type: 'path', path: 'M0.5,0 L0.61,0.35 L1,0.38 L0.71,0.62 L0.79,1 L0.5,0.77 L0.21,1 L0.29,0.62 L0,0.38 L0.39,0.35 Z' },
      star6: { type: 'path', path: 'M0.5,0 L0.6,0.33 L0.93,0.33 L0.67,0.55 L0.77,0.87 L0.5,0.67 L0.23,0.87 L0.33,0.55 L0.07,0.33 L0.4,0.33 Z' },
      star10: { type: 'path', path: 'M0.5,0 L0.55,0.2 L0.75,0.2 L0.6,0.35 L0.65,0.55 L0.5,0.45 L0.35,0.55 L0.4,0.35 L0.25,0.2 L0.45,0.2 Z' },
      // Arrows
      rightArrow: { type: 'path', path: 'M0,0.33 L0.67,0.33 L0.67,0 L1,0.5 L0.67,1 L0.67,0.67 L0,0.67 Z' },
      leftArrow: { type: 'path', path: 'M0,0.5 L0.33,0 L0.33,0.33 L1,0.33 L1,0.67 L0.33,0.67 L0.33,1 Z' },
      upArrow: { type: 'path', path: 'M0.33,0 L0.33,0.67 L0,0.67 L0.5,1 L1,0.67 L0.67,0.67 L0.67,0 Z' },
      downArrow: { type: 'path', path: 'M0,0.33 L0.5,0 L1,0.33 L0.67,0.33 L0.67,1 L0.33,1 L0.33,0.33 Z' },
      // Callouts
      roundedRectangularCallout: { type: 'rect', radius: 0.2 },
      ovalCallout: { type: 'ellipse' },
      // Basic shapes - note: 'rect' already defined above, removing duplicate
      square: { type: 'rect', radius: 0 },
    };

    return shapeTypeMap[pptxShapeType] || defaultShape;
  }

  /**
   * Convert effects
   */
  private convertEffects(effects: any, pptistElement: ShapeElement): void {
    if (effects.shadow) {
      pptistElement.shadow = this.convertShadow(effects.shadow);
    }

    if (effects.reflection) {
      pptistElement.reflection = this.convertReflection(effects.reflection);
    }

    if (effects.glow) {
      pptistElement.glow = this.convertGlow(effects.glow);
    }

    if (effects.blur) {
      pptistElement.blur = this.convertBlur(effects.blur);
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
   * Convert reflection effect
   */
  private convertReflection(reflection: any): any {
    const pptistReflection: any = {};

    if (reflection.opacity !== undefined) pptistReflection.opacity = reflection.opacity;
    if (reflection.blur !== undefined) pptistReflection.blur = reflection.blur;
    if (reflection.distance !== undefined) pptistReflection.distance = reflection.distance;
    if (reflection.direction !== undefined) pptistReflection.direction = reflection.direction;
    if (reflection.fade !== undefined) pptistReflection.fade = reflection.fade;

    return pptistReflection;
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

  /**
   * Convert blur effect
   */
  private convertBlur(blur: any): any {
    const pptistBlur: any = {};

    if (blur.radius !== undefined) pptistBlur.radius = blur.radius;

    return pptistBlur;
  }

  /**
   * Convert text box
   */
  private convertTextBox(textBox: any, pptistElement: ShapeElement, context: ConversionContext): void {
    if (!textBox.paragraphs) return;

    pptistElement.text = textBox.paragraphs
      .map((p: any) => p.text || '')
      .join('\n');
  }
}

/**
 * Create shape converter instance
 */
export function createShapeConverter(): ShapeConverter {
  return new ShapeConverter();
}
