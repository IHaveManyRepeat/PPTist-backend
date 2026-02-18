/**
 * Image Element Converter
 *
 * Converts PPTX image elements to PPTist image elements.
 * Handles embedded base64 images and external image files.
 *
 * @module services/conversion/converters/image
 */

import { BaseElementConverter } from '../base-converter';
import { ElementType, type ConversionContext } from '../../../types/converters';
import type { ImageElement } from '../../../types/pptist';
import type { ParsedElement } from '../../../services/pptx/parser';
import { emuToPixels } from '../../../utils/coordinates';
import { normalizeColor } from '../../../utils/color';
import { generateTrackedId, ID_PREFIXES } from '../../../utils/id-generator';
import { logger } from '../../../utils/logger';

/**
 * Image converter implementation
 */
export class ImageConverter extends BaseElementConverter<ParsedElement, ImageElement> {
  readonly type = ElementType.IMAGE;
  readonly priority = 100;
  readonly supportedVersions = ['1.0.0', 'latest'];

  /**
   * Check if this converter can handle the element
   */
  async canConvert(element: ParsedElement): Promise<boolean> {
    return element.type === 'image';
  }

  /**
   * Convert PPTX image element to PPTist image element
   */
  async convert(element: ParsedElement, context: ConversionContext): Promise<ImageElement | null> {
    logger.debug('Converting image element', {
      id: element.id,
      imageRef: element.imageRef,
    });

    // Convert EMU to pixels (apply 96/72 scaling)
    const x = emuToPixels(element.position?.x || 0);
    const y = emuToPixels(element.position?.y || 0);
    const width = emuToPixels(element.size?.width || 914400); // Default 1 inch
    const height = emuToPixels(element.size?.height || 914400); // Default 1 inch

    const pptistElement: ImageElement = {
      id: generateTrackedId(ID_PREFIXES.IMAGE, undefined, element.id),
      type: 'image',
      x,
      y,
      width,
      height,
      rotate: element.rotation || 0,
      locked: element.locked || false,
      visible: !element.hidden,
      zIndex: element.zIndex,
      src: this.resolveImageSource(element, context),
      flip: undefined,
    };

    // Convert crop
    if (element.crop) {
      pptistElement.crop = this.convertCrop(element.crop);
    }

    // Convert effects
    if (element.effects) {
      this.convertEffects(element.effects, pptistElement);
    }

    // Convert fill (for image border)
    if (element.fill) {
      pptistElement.fill = this.convertFill(element.fill);
    }

    // Convert stroke (for image border)
    if (element.stroke) {
      pptistElement.stroke = this.convertStroke(element.stroke);
    }

    return pptistElement;
  }

  /**
   * Resolve image source (base64 or file path)
   */
  private resolveImageSource(element: ParsedElement, context: ConversionContext): string {
    if (!element.imageRef) {
      logger.warn('Image element has no imageRef', { id: element.id });
      return '';
    }

    // Resolve media reference from context
    const mediaFile = context.resolveMediaReference(element.imageRef);

    if (mediaFile) {
      // Check if media is embedded as base64
      if (mediaFile.isEmbedded && mediaFile.base64) {
        return `data:${mediaFile.contentType};base64,${mediaFile.base64}`;
      }

      // External file path
      if (mediaFile.path) {
        return mediaFile.path;
      }

      // Fallback to data URL
      if (mediaFile.data) {
        return `data:${mediaFile.contentType};base64,${mediaFile.data.toString('base64')}`;
      }
    }

    logger.warn('Failed to resolve image reference', {
      imageRef: element.imageRef,
      id: element.id,
    });

    return '';
  }

  /**
   * Convert crop values
   */
  private convertCrop(crop: any): any {
    const pptistCrop: any = {};

    if (crop.top !== undefined) pptistCrop.top = crop.top;
    if (crop.right !== undefined) pptistCrop.right = crop.right;
    if (crop.bottom !== undefined) pptistCrop.bottom = crop.bottom;
    if (crop.left !== undefined) pptistCrop.left = crop.left;

    return pptistCrop;
  }

  /**
   * Convert effects
   */
  private convertEffects(effects: any, pptistElement: ImageElement): void {
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
   * Convert fill
   */
  private convertFill(fill: any): any {
    if (!fill) return undefined;

    if (fill.type === 'solid' && fill.color) {
      return {
        type: 'solid',
        color: normalizeColor(fill.color),
      };
    }

    return undefined;
  }

  /**
   * Convert stroke
   */
  private convertStroke(stroke: any): any {
    if (!stroke) return undefined;

    const pptistStroke: any = {};

    if (stroke.width !== undefined) pptistStroke.width = stroke.width;
    if (stroke.color) pptistStroke.color = normalizeColor(stroke.color);
    if (stroke.dashType) pptistStroke.style = stroke.dashType;

    return pptistStroke;
  }
}

/**
 * Create image converter instance
 */
export function createImageConverter(): ImageConverter {
  return new ImageConverter();
}
