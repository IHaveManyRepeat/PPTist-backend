/**
 * 内容类型解析器
 *
 * @module modules/conversion/services/parser/content-types
 * @description 解析 PPTX 的 [Content_Types].xml 文件，提取幻灯片和布局文件列表。
 */

import type { XmlObject } from '../../context/parsing-context.js'
import { readXmlFile } from './utils.js'

/**
 * 内容类型解析结果
 */
export interface ContentTypesResult {
  /** 幻灯片文件路径数组 */
  slides: string[]
  /** 布局文件路径数组 */
  slideLayouts: string[]
}

/**
 * 按数字排序文件路径
 *
 * @param arr - 文件路径数组
 * @returns 排序后的数组
 */
function sortByNumber(arr: string[]): string[] {
  return arr.sort((a, b) => {
    const n1 = parseInt(/(\d+)\.xml/.exec(a)?.[1] || '0', 10)
    const n2 = parseInt(/(\d+)\.xml/.exec(b)?.[1] || '0', 10)
    return n1 - n2
  })
}

/**
 * 解析 [Content_Types].xml 获取幻灯片和布局文件列表
 *
 * @description
 * 从 PPTX 的 [Content_Types].xml 中提取幻灯片和布局文件的位置信息。
 * 这些信息用于确定演示文稿的结构和幻灯片顺序。
 *
 * @param zip - JSZip 实例
 * @returns 包含幻灯片和布局文件路径数组的对象
 *
 * @example
 * ```typescript
 * const { slides, slideLayouts } = await getContentTypes(zip);
 * console.log(`找到 ${slides.length} 张幻灯片`);
 * ```
 */
export async function getContentTypes(
  zip: import('jszip').default
): Promise<ContentTypesResult> {
  const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
  const subObj = ContentTypesJson?.['Types']?.['Override'] || []

  const slidesLocArray: string[] = []
  const slideLayoutsLocArray: string[] = []

  const overrides = Array.isArray(subObj) ? subObj : [subObj]

  for (const item of overrides) {
    const contentType = item?.['attrs']?.['ContentType']
    const partName = item?.['attrs']?.['PartName']

    if (!contentType || !partName) continue

    if (contentType === 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml') {
      slidesLocArray.push(partName.substr(1))
    } else if (contentType === 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml') {
      slideLayoutsLocArray.push(partName.substr(1))
    }
  }

  return {
    slides: sortByNumber(slidesLocArray),
    slideLayouts: sortByNumber(slideLayoutsLocArray),
  }
}

export default { getContentTypes }
