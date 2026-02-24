/**
 * 图形框架解析器
 *
 * @module modules/conversion/services/parser/elements/graphic-frame
 * @description 解析 PPTX 中的图形框架（graphicFrame），包括图表和表格。
 */

import { v4 as uuidv4 } from 'uuid'
import type { PPTXElement, PPTXTransform } from '../../../types/pptx.js'
import type { XmlObject, ParsingContext } from '../../../context/parsing-context.js'
import { parseTable } from '../../../parsers/table-parser.js'
import { parseChart } from '../../../parsers/chart-parser.js'

/**
 * 解析图形框架（图表/表格）
 *
 * @description
 * 从 PPTX 的 graphicFrame 节点中提取图表或表格信息。
 * 根据 graphicData 的 URI 判断是图表还是表格。
 *
 * @param frame - 图形框架 XML 节点
 * @param context - 解析上下文
 * @returns 解析后的 PPTX 元素，如果不支持返回 null
 */
export async function parseGraphicFrame(
  frame: XmlObject,
  context: ParsingContext
): Promise<PPTXElement | null> {
  const nvGraphicFramePr = frame['p:nvGraphicFramePr']
  const xfrm = frame['p:xfrm']
  const graphic = frame['a:graphic']

  const off = xfrm?.['a:off']?.['attrs'] || {}
  const ext = xfrm?.['a:ext']?.['attrs'] || {}

  const transform: PPTXTransform = {
    x: parseInt(off['x'] || '0', 10),
    y: parseInt(off['y'] || '0', 10),
    width: parseInt(ext['cx'] || '0', 10),
    height: parseInt(ext['cy'] || '0', 10),
  }

  const cNvPr = nvGraphicFramePr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  // 检查图形类型
  const graphicData = graphic?.['a:graphicData']
  const uri = graphicData?.['attrs']?.['uri'] as string

  // 图表
  if (uri?.includes('chart')) {
    const chartRId = graphicData?.['c:chart']?.['attrs']?.['r:id'] as string
    if (chartRId) {
      return parseChart(frame, chartRId, transform, context)
    }
    return {
      type: 'chart',
      id: String(id),
      transform,
      name,
      chartType: 'column',
      rId: '',
    }
  }

  // 表格
  if (uri?.includes('table')) {
    return parseTable(frame, transform, context)
  }

  // 不支持的图形类型
  return null
}

export default { parseGraphicFrame }
