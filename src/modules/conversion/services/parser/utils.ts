/**
 * 解析器工具函数
 *
 * @module modules/conversion/services/parser/utils
 * @description 提供 PPTX 解析器共用的工具函数。
 */

import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'
import type { XmlObject } from '../../context/parsing-context.js'

/**
 * 读取 ZIP 中的 XML 文件并解析为对象
 *
 * @param zip - JSZip 实例
 * @param path - ZIP 内的文件路径
 * @returns 解析后的 XML 对象，如果文件不存在则返回空对象
 *
 * @example
 * ```typescript
 * const content = await readXmlFile(zip, 'ppt/presentation.xml');
 * const slideSize = content?.['p:presentation']?.['p:sldSz']?.['attrs'];
 * ```
 */
export async function readXmlFile(zip: JSZip, path: string): Promise<XmlObject> {
  const content = await zip.file(path)?.async('string')
  if (!content) return {}

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '',
    attributesGroupName: 'attrs',
    textNodeName: '#text',
  })

  return parser.parse(content) as XmlObject
}

/**
 * 获取 MIME 类型
 *
 * @param path - 文件路径
 * @returns MIME 类型字符串
 */
export function getMimeType(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase()
  const mimeTypes: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    mp4: 'video/mp4',
    avi: 'video/x-msvideo',
    mov: 'video/quicktime',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
  }
  return mimeTypes[ext || ''] || 'application/octet-stream'
}

export default { readXmlFile, getMimeType }
