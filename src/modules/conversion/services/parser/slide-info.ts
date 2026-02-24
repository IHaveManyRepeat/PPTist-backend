/**
 * 幻灯片信息解析器
 *
 * @module modules/conversion/services/parser/slide-info
 * @description 解析 presentation.xml 获取幻灯片基本信息。
 */

import type JSZip from 'jszip'
import type { XmlObject } from '../../context/parsing-context.js'
import { readXmlFile } from './utils.js'

/**
 * 幻灯片信息解析结果
 */
export interface SlideInfoResult {
  /** 幻灯片宽度（EMU） */
  width: number
  /** 幻灯片高度（EMU） */
  height: number
  /** 默认文本样式 */
  defaultTextStyle: XmlObject
}

/**
 * 获取演示文稿的基本信息
 *
 * @description
 * 从 presentation.xml 中提取幻灯片尺寸和默认文本样式。
 * 幻灯片尺寸以 EMU (English Metric Units) 为单位。
 *
 * @param zip - JSZip 实例
 * @returns 包含宽度、高度和默认文本样式的对象
 *
 * @example
 * ```typescript
 * const { width, height } = await getSlideInfo(zip);
 * // width 和 height 以 EMU 为单位
 * // 转换为像素: pixels = EMU / 914400 * 96
 * ```
 */
export async function getSlideInfo(zip: JSZip): Promise<SlideInfoResult> {
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const sldSzAttrs = content?.['p:presentation']?.['p:sldSz']?.['attrs'] || {}
  const defaultTextStyle = content?.['p:presentation']?.['p:defaultTextStyle'] || {}

  return {
    width: parseInt(sldSzAttrs['cx'] || '9144000', 10),
    height: parseInt(sldSzAttrs['cy'] || '6858000', 10),
    defaultTextStyle,
  }
}

export default { getSlideInfo }
