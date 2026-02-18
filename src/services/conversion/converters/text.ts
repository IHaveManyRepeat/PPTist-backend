/**
 * Text Element Converter
 *
 * Converts PPTX text elements to PPTist text elements.
 * Handles text content, formatting, fonts, colors, and alignment.
 *
 * @module services/conversion/converters/text
 */

import { BaseElementConverter } from '../base-converter';
import { ElementType, type ConversionContext } from '../../../types/converters';
import type { TextElement } from '../../../types/pptist';
import type { ParsedElement } from '../../../services/pptx/parser';
import { parseOfficeColor, normalizeColor } from '../../../utils/color';
import { emuToPixels, fontSizePointsToPixels } from '../../../utils/coordinates';
import { generateTrackedId, ID_PREFIXES } from '../../../utils/id-generator';
import { logger } from '../../../utils/logger';

/**
 * PPTX字体到Web字体的映射
 * 未在映射表中的字体将使用默认字体
 */
const FONT_FAMILY_MAP: Record<string, string> = {
  // 中文字体
  '微软雅黑': 'Microsoft YaHei',
  '宋体': 'SimSun',
  '黑体': 'SimHei',
  '楷体': 'KaiTi',
  '仿宋': 'FangSong',
  '等线': 'DengXian',
  '新宋体': 'NSimSun',
  '华文黑体': 'STHeiti',
  '华文宋体': 'STSong',
  '华文楷体': 'STKaiti',

  // 英文字体
  'Arial': 'Arial',
  'Times New Roman': 'Times New Roman',
  'Calibri': 'Calibri',
  'Cambria': 'Cambria',
  'Courier New': 'Courier New',
  'Georgia': 'Georgia',
  'Verdana': 'Verdana',
  'Tahoma': 'Tahoma',
  'Trebuchet MS': 'Trebuchet MS',
  'Impact': 'Impact',
  'Comic Sans MS': 'Comic Sans MS',
  'Century Gothic': 'Century Gothic',
  'Palatino Linotype': 'Palatino Linotype',
};

/** 默认字体 */
const DEFAULT_FONT = 'Arial';

/**
 * 映射字体名称
 * - 在映射表中：返回映射的字体
 * - 不在映射表：返回默认字体
 */
function mapFontFamily(fontName: string | undefined): string {
  if (!fontName) return DEFAULT_FONT;

  // 检查是否有直接映射
  if (FONT_FAMILY_MAP[fontName]) {
    return FONT_FAMILY_MAP[fontName];
  }

  // 检查大小写不敏感的映射
  const lowerFontName = fontName.toLowerCase();
  for (const [key, value] of Object.entries(FONT_FAMILY_MAP)) {
    if (key.toLowerCase() === lowerFontName) {
      return value;
    }
  }

  // 未映射的字体使用默认字体
  logger.debug(`Font "${fontName}" not in mapping, using default: ${DEFAULT_FONT}`);
  return DEFAULT_FONT;
}

/**
 * Text converter implementation
 */
export class TextConverter extends BaseElementConverter<ParsedElement, TextElement> {
  readonly type = ElementType.TEXT;
  readonly priority = 100;
  readonly supportedVersions = ['1.0.0', 'latest'];

  /**
   * Check if this converter can handle the element
   */
  async canConvert(element: ParsedElement): Promise<boolean> {
    return element.type === 'text' || element.textBox !== undefined;
  }

  /**
   * Convert PPTX text element to PPTist text element
   */
  async convert(element: ParsedElement, context: ConversionContext): Promise<TextElement | null> {
    logger.debug('Converting text element', {
      id: element.id,
      hasTextBox: !!element.textBox,
    });

    // Convert EMU to pixels (apply 96/72 scaling)
    const x = emuToPixels(element.position?.x || 0);
    const y = emuToPixels(element.position?.y || 0);
    const width = emuToPixels(element.size?.width || 914400); // Default 1 inch
    const height = emuToPixels(element.size?.height || 457200); // Default 0.5 inch

    const pptistElement: TextElement = {
      id: generateTrackedId(ID_PREFIXES.TEXT, undefined, element.id),
      type: 'text',
      x,
      y,
      width,
      height,
      rotate: element.rotation || 0,
      locked: element.locked || false,
      visible: !element.hidden,
      zIndex: element.zIndex,
      defaultValue: this.extractTextContent(element),
    };

    // Convert text formatting
    if (element.textBox) {
      this.convertTextBoxFormatting(element.textBox, pptistElement, context);
    }

    // Convert fill (background color)
    if (element.fill) {
      pptistElement.fill = this.convertFill(element.fill, context);
    }

    // Convert stroke (border)
    if (element.stroke) {
      pptistElement.stroke = this.convertStroke(element.stroke);
    }

    // Convert effects (shadow, glow, etc.)
    if (element.effects) {
      this.convertEffects(element.effects, pptistElement);
    }

    return pptistElement;
  }

  /**
   * Extract text content from text box
   */
  private extractTextContent(element: ParsedElement): string {
    if (!element.textBox) return '';

    const paragraphs = element.textBox.paragraphs || [];
    return paragraphs.map((p: any) => p.text || p.runs?.map((r: any) => r.text || '').join('') || '').join('\n');
  }

  /**
   * Convert text box formatting
   */
  private convertTextBoxFormatting(
    textBox: any,
    pptistElement: TextElement,
    context: ConversionContext
  ): void {
    if (!textBox.paragraphs) return;

    // Convert paragraphs to PPTist format
    pptistElement.content = textBox.paragraphs.map((p: any) => this.convertParagraph(p, context));

    // Set default vertical alignment
    if (textBox.verticalAlign) {
      pptistElement.verticalAlign = this.convertVerticalAlign(textBox.verticalAlign);
    }
  }

  /**
   * Convert paragraph
   */
  private convertParagraph(p: any, context: ConversionContext): any {
    const paragraph: any = {
      text: p.text || '',
    };

    // Convert text runs (formatting)
    if (p.runs && p.runs.length > 0) {
      paragraph.formatting = p.runs.map((r: any) => this.convertTextRun(r, context));
    }

    // Paragraph alignment
    if (p.alignment) {
      paragraph.align = this.convertTextAlign(p.alignment);
    }

    // Paragraph spacing
    if (p.spacing) {
      paragraph.lineSpacing = p.spacing.line;
      paragraph.spaceBefore = p.spacing.before;
      paragraph.spaceAfter = p.spacing.after;
    }

    // Paragraph indent
    if (p.indent !== undefined) {
      paragraph.indent = p.indent;
    }

    return paragraph;
  }

  /**
   * Convert text run (inline formatting)
   */
  private convertTextRun(r: any, context: ConversionContext): any {
    const run: any = {
      text: r.text || '',
    };

    // Font family - 使用字体映射
    if (r.font) {
      run.fontFamily = mapFontFamily(r.font);
    }

    // Font size (convert from points to pixels)
    if (r.size) {
      run.fontSize = fontSizePointsToPixels(r.size);
    }

    // Font weight
    if (r.bold) {
      run.fontWeight = 'bold';
    }

    // Font style
    if (r.italic) {
      run.fontStyle = 'italic';
    }

    // Text decoration
    if (r.underline) {
      run.textDecoration = 'underline';
    }

    // Text color
    if (r.color) {
      run.color = normalizeColor(r.color);
    }

    return run;
  }

  /**
   * Convert text alignment
   */
  private convertTextAlign(align: string): string {
    const alignmentMap: Record<string, string> = {
      left: 'left',
      r: 'right',  // 'r' typically means 'right' in alignment
      center: 'center',
      ctr: 'center',
      right: 'right',
      justify: 'justify',
      just: 'justify',
      dist: 'justify',
      thaiDist: 'justify',
    };

    return alignmentMap[align] || 'left';
  }

  /**
   * Convert vertical alignment
   */
  private convertVerticalAlign(align: string): 'top' | 'middle' | 'bottom' {
    const alignmentMap: Record<string, 'top' | 'middle' | 'bottom'> = {
      top: 'top',
      t: 'top',
      middle: 'middle',
      ctr: 'middle',
      bottom: 'bottom',
      b: 'bottom',
      just: 'top',
      dist: 'bottom',
    };

    return alignmentMap[align] || 'top';
  }

  /**
   * Convert fill (background color)
   */
  private convertFill(fill: any, context: ConversionContext): any {
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

    if (fill.type === 'gradient') {
      return {
        type: 'gradient',
        colors: (fill.colors || []).map((c: string) => normalizeColor(c)),
      };
    }

    if (fill.type === 'image' && fill.imageRef) {
      const resolvedImage = context.resolveMediaReference(fill.imageRef);
      if (resolvedImage) {
        return {
          type: 'image',
          image: resolvedImage,
        };
      }
    }

    return undefined;
  }

  /**
   * Convert stroke (border)
   */
  private convertStroke(stroke: any): any {
    if (!stroke) return undefined;

    const pptistStroke: any = {};

    if (stroke.width !== undefined) {
      pptistStroke.width = stroke.width;
    }

    if (stroke.color) {
      pptistStroke.color = normalizeColor(stroke.color);
    }

    if (stroke.dashType) {
      pptistStroke.style = this.convertStrokeDashType(stroke.dashType);
    }

    return pptistStroke;
  }

  /**
   * Convert stroke dash type
   */
  private convertStrokeDashType(dashType: string): string {
    const dashMap: Record<string, string> = {
      solid: 'solid',
      dash: 'dashed',
      dashDot: 'dash-dot',
      dot: 'dotted',
      lgDash: 'dashed',
      lgDashDot: 'dash-dot',
      lgDashDotDot: 'dash-dot-dot',
      sysDash: 'dashed',
      sysDashDot: 'dash-dot',
      sysDashDotDot: 'dash-dot-dot',
      sysDot: 'dotted',
    };

    return dashMap[dashType] || 'solid';
  }

  /**
   * Convert effects (shadow, glow, etc.)
   */
  private convertEffects(effects: any, pptistElement: TextElement): void {
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

    if (shadow.type) {
      pptistShadow.type = shadow.type;
    }

    if (shadow.color) {
      pptistShadow.color = normalizeColor(shadow.color);
    }

    if (shadow.offset !== undefined) {
      pptistShadow.offset = shadow.offset;
    }

    if (shadow.blur !== undefined) {
      pptistShadow.blur = shadow.blur;
    }

    if (shadow.angle !== undefined) {
      pptistShadow.angle = shadow.angle;
    }

    if (shadow.opacity !== undefined) {
      pptistShadow.opacity = shadow.opacity;
    }

    return pptistShadow;
  }

  /**
   * Convert glow effect
   */
  private convertGlow(glow: any): any {
    const pptistGlow: any = {};

    if (glow.color) {
      pptistGlow.color = normalizeColor(glow.color);
    }

    if (glow.radius !== undefined) {
      pptistGlow.radius = glow.radius;
    }

    return pptistGlow;
  }
}

/**
 * Create text converter instance
 */
export function createTextConverter(): TextConverter {
  return new TextConverter();
}
