/**
 * 主题解析器
 *
 * @module modules/conversion/services/parser/theme
 * @description 解析 PPTX 的主题文件，提取颜色方案等信息。
 */

import type JSZip from 'jszip'
import type { XmlObject } from '../../context/parsing-context.js'
import { getTextByPathList } from '../../resolvers/color-resolver.js'
import { readXmlFile } from './utils.js'

/**
 * 主题解析结果
 */
export interface ThemeResult {
  /** 主题 XML 内容 */
  themeContent: XmlObject
  /** 主题颜色数组（accent1-accent6） */
  themeColors: string[]
  /** 主题文件路径 */
  themePath?: string
}

/**
 * 获取演示文稿的主题信息
 *
 * @description
 * 从主题文件中提取颜色方案和其他样式信息。
 * 主题颜色用于解析 schemeClr 类型的颜色引用。
 *
 * @param zip - JSZip 实例
 * @returns 包含主题内容、主题颜色数组和主题路径的对象
 *
 * @example
 * ```typescript
 * const { themeContent, themeColors } = await getTheme(zip);
 * // themeColors 包含 accent1-accent6 的颜色值
 * ```
 */
export async function getTheme(zip: JSZip): Promise<ThemeResult> {
  // 从 presentation.xml.rels 获取主题路径
  const preResContent = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
  const relationshipArray = preResContent?.['Relationships']?.['Relationship'] || []
  const relationships = Array.isArray(relationshipArray) ? relationshipArray : [relationshipArray]

  let themeURI: string | undefined
  for (const rel of relationships) {
    if (rel?.['attrs']?.['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
      themeURI = rel['attrs']['Target']
      break
    }
  }

  if (!themeURI) {
    return { themeContent: {}, themeColors: [] }
  }

  const themePath = `ppt/${themeURI}`
  const themeContent = await readXmlFile(zip, themePath)

  // 提取主题颜色
  const themeColors: string[] = []
  const clrScheme = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:clrScheme'])

  if (clrScheme) {
    for (let i = 1; i <= 6; i++) {
      const colorResult = getTextByPathList(clrScheme as XmlObject, [`a:accent${i}`, 'a:srgbClr', 'attrs', 'val'])
      const color = typeof colorResult === 'string' ? colorResult : undefined
      if (color) {
        themeColors.push(color.startsWith('#') ? color : '#' + color)
      }
    }
  }

  return { themeContent, themeColors, themePath }
}

export default { getTheme }
