/**
 * 关系文件解析器
 *
 * @module modules/conversion/services/parser/relationships
 * @description 解析 PPTX 的 .rels 关系文件，提取资源映射和引用关系。
 */

import type JSZip from 'jszip'
import type { ResourceMap } from '../../context/parsing-context.js'
import { readXmlFile } from './utils.js'

/**
 * 关系解析结果
 */
export interface RelationshipsResult {
  /** 资源映射（rId -> 资源信息） */
  resources: ResourceMap
  /** 布局文件名 */
  layoutFilename?: string
  /** 母版文件名 */
  masterFilename?: string
  /** 主题文件名 */
  themeFilename?: string
  /** 备注文件名 */
  noteFilename?: string
}

/**
 * 解析关系文件
 *
 * @description
 * 解析 .rels 文件，提取资源 ID 到目标路径的映射，
 * 以及特殊关系类型（布局、母版、主题、备注）的文件名。
 *
 * @param zip - JSZip 实例
 * @param relsPath - 关系文件路径
 * @returns 关系解析结果
 *
 * @example
 * ```typescript
 * const { resources, layoutFilename } = await parseRelationships(
 *   zip,
 *   'ppt/slides/_rels/slide1.xml.rels'
 * );
 * ```
 */
export async function parseRelationships(
  zip: JSZip,
  relsPath: string
): Promise<RelationshipsResult> {
  const resources: ResourceMap = {}
  let layoutFilename: string | undefined
  let masterFilename: string | undefined
  let themeFilename: string | undefined
  let noteFilename: string | undefined

  try {
    const relsContent = await readXmlFile(zip, relsPath)
    let relationships = relsContent?.['Relationships']?.['Relationship'] || []
    relationships = Array.isArray(relationships) ? relationships : [relationships]

    for (const rel of relationships) {
      const id = rel?.['attrs']?.['Id']
      const type = rel?.['attrs']?.['Type']
      const target = rel?.['attrs']?.['Target']

      if (!id || !type || !target) continue

      const typeName = type.replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
      const normalizedTarget = target.replace('../', 'ppt/')

      resources[id] = { type: typeName, target: normalizedTarget }

      // 识别特殊关系
      if (typeName === 'slideLayout') {
        layoutFilename = normalizedTarget
      } else if (typeName === 'slideMaster') {
        masterFilename = normalizedTarget
      } else if (typeName === 'theme') {
        themeFilename = normalizedTarget
      } else if (typeName === 'notesSlide') {
        noteFilename = normalizedTarget
      }
    }
  } catch {
    // 关系文件可能不存在
  }

  return { resources, layoutFilename, masterFilename, themeFilename, noteFilename }
}

/**
 * 解析幻灯片关系文件获取 rId -> media target 映射
 *
 * @param zip - JSZip 实例
 * @param slideNum - 幻灯片编号
 * @returns rId 到媒体路径的映射
 */
export async function parseSlideRels(
  zip: JSZip,
  slideNum: number
): Promise<Map<string, string>> {
  const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`
  const relsXml = await zip.file(relsPath)?.async('string')

  const rIdToTarget = new Map<string, string>()

  if (relsXml) {
    const { XMLParser } = await import('fast-xml-parser')
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '',
      textNodeName: '#text',
    })
    const rels = parser.parse(relsXml) as any
    const relationships = rels?.['Relationships']?.['Relationship'] || []
    const relArray = Array.isArray(relationships) ? relationships : [relationships]

    for (const rel of relArray) {
      const id = rel['Id'] as string
      const target = rel['Target'] as string
      if (id && target && target.includes('media/')) {
        const normalizedTarget = target.replace(/^\.\.\//, '')
        rIdToTarget.set(id, normalizedTarget)
      }
    }
  }

  return rIdToTarget
}

export default { parseRelationships, parseSlideRels }
