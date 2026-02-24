/**
 * 图片元素解析器
 *
 * @module modules/conversion/services/parser/elements/picture
 * @description 解析 PPTX 中的图片/视频/音频元素（pic）。
 */

import { v4 as uuidv4 } from 'uuid'
import type { PPTXElement } from '../../../types/pptx.js'
import type { XmlObject, ParsingContext } from '../../../context/parsing-context.js'
import { parseTransform } from './shape.js'

/**
 * 解析图片/视频/音频元素
 *
 * @description
 * 从 PPTX 的 pic 节点中提取媒体元素信息。
 * 根据嵌入的关系类型，返回 image、video 或 audio 元素。
 *
 * @param pic - 图片 XML 节点
 * @param context - 解析上下文（未使用，但保持接口一致）
 * @returns 解析后的 PPTX 元素，如果解析失败返回 null
 */
export function parsePicture(
  pic: XmlObject,
  _context: ParsingContext
): PPTXElement | null {
  const nvPicPr = pic['p:nvPicPr']
  const spPr = pic['p:spPr']
  const blipFill = pic['p:blipFill']

  const transform = parseTransform(spPr)
  const cNvPr = nvPicPr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  // 获取 rId
  const rId = blipFill?.['a:blip']?.['attrs']?.['r:embed'] as string
  if (!rId) return null

  // 检查是否是视频/音频
  const nvPr = nvPicPr?.['p:nvPr']
  const videoFile = nvPr?.['a:videoFile']?.['attrs']?.['r:link'] as string
  const audioFile = nvPr?.['a:audioFile']?.['attrs']?.['r:link'] as string

  if (videoFile) {
    return {
      type: 'video',
      id: String(id),
      transform,
      name,
      rId: videoFile,
    }
  }

  if (audioFile) {
    return {
      type: 'audio',
      id: String(id),
      transform,
      name,
      rId: audioFile,
    }
  }

  return {
    type: 'image',
    id: String(id),
    transform,
    name,
    rId,
  }
}

export default { parsePicture }
