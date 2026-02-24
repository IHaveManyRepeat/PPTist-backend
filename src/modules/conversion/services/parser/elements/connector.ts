/**
 * 连接线解析器
 *
 * @module modules/conversion/services/parser/elements/connector
 * @description 解析 PPTX 中的连接线元素（cxnSp）。
 */

import { v4 as uuidv4 } from 'uuid'
import type { PPTXElement } from '../../../types/pptx.js'
import type { XmlObject, ParsingContext } from '../../../context/parsing-context.js'
import { parseTransform } from './shape.js'
import { resolveSolidFill } from '../../../resolvers/color-resolver.js'

/** EMU 到点的转换比例 */
const RATIO_EMUs_Points = 1 / 12700

/**
 * 解析连接线元素
 *
 * @description
 * 从 PPTX 的 cxnSp 节点中提取连接线信息。
 *
 * @param conn - 连接线 XML 节点
 * @param context - 解析上下文
 * @returns 解析后的 PPTX 元素
 */
export function parseConnector(
  conn: XmlObject,
  context: ParsingContext
): PPTXElement | null {
  const nvCxnSpPr = conn['p:nvCxnSpPr']
  const spPr = conn['p:spPr']

  const transform = parseTransform(spPr)
  const cNvPr = nvCxnSpPr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  const ln = spPr?.['a:ln']
  const solidFill = ln?.['a:solidFill']
  const color = solidFill ? resolveSolidFill(solidFill, context) : undefined
  const width = ln?.['attrs']?.['w'] ? parseInt(ln['attrs']['w'], 10) * RATIO_EMUs_Points : undefined

  return {
    type: 'line',
    id: String(id),
    transform,
    name,
    startX: transform.x,
    startY: transform.y,
    endX: transform.x + transform.width,
    endY: transform.y + transform.height,
    color,
    width,
  }
}

export default { parseConnector }
