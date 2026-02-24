/**
 * 形状元素解析器
 *
 * @module modules/conversion/services/parser/elements/shape
 * @description 解析 PPTX 中的形状元素（sp）。
 */

import { v4 as uuidv4 } from 'uuid'
import type {
  PPTXElement,
  PPTXShapeElement,
  PPTXTransform,
  PPTXParagraph,
} from '../../../types/pptx.js'
import type { XmlObject, ParsingContext } from '../../../context/parsing-context.js'
import { getTextByPathList, resolveSolidFill, resolveSolidFillWithAlpha } from '../../../resolvers/color-resolver.js'
import { generateShapePath } from '../../../generators/svg-path-generator.js'

/** EMU 到点的转换比例 */
const RATIO_EMUs_Points = 1 / 12700

/**
 * 解析变换信息
 *
 * @param spPr - 形状属性节点
 * @returns 变换信息对象
 */
export function parseTransform(spPr: XmlObject | undefined): PPTXTransform {
  const xfrm = spPr?.['a:xfrm']
  const off = xfrm?.['a:off']?.['attrs'] || {}
  const ext = xfrm?.['a:ext']?.['attrs'] || {}
  const rot = xfrm?.['attrs']?.['rot']

  return {
    x: parseInt(off['x'] || '0', 10),
    y: parseInt(off['y'] || '0', 10),
    width: parseInt(ext['cx'] || '0', 10),
    height: parseInt(ext['cy'] || '0', 10),
    rotation: rot ? parseInt(rot, 10) / 60000 : undefined,
  }
}

/**
 * 解析文本运行的颜色（支持多层级继承）
 *
 * @description
 * 继承优先级（从高到低）：
 * 1. 运行级别: a:r/a:rPr/a:solidFill
 * 2. 段落默认: a:p/a:pPr/a:defRPr/a:solidFill
 * 3. 列表级别: a:p/a:pPr/a:lvlXpPr/a:defRPr/a:solidFill (X = 0-8)
 * 4. 文本体列表样式: a:txBody/a:lstStyle/a:defRPr/a:solidFill
 */
function resolveTextColor(
  run: XmlObject,
  paragraph: XmlObject,
  txBody: XmlObject,
  context: ParsingContext
): string | undefined {
  // 1. 运行级别
  const runSolidFill = run?.['a:rPr']?.['a:solidFill']
  if (runSolidFill) {
    return resolveSolidFill(runSolidFill, context)
  }

  // 2. 段落默认运行属性
  const pPr = paragraph?.['a:pPr']
  const defRPrSolidFill = pPr?.['a:defRPr']?.['a:solidFill']
  if (defRPrSolidFill) {
    return resolveSolidFill(defRPrSolidFill, context)
  }

  // 3. 列表级别样式
  if (pPr) {
    const levelAttr = pPr?.['attrs']?.['lvl']
    if (levelAttr !== undefined) {
      const level = parseInt(String(levelAttr), 10)
      const lvlPPr = pPr?.[`a:lvl${level}pPr`]
      const lvlDefRPrSolidFill = lvlPPr?.['a:defRPr']?.['a:solidFill']
      if (lvlDefRPrSolidFill) {
        return resolveSolidFill(lvlDefRPrSolidFill, context)
      }
    }
  }

  // 4. 文本体列表样式默认
  const lstStyleDefRPrSolidFill = txBody?.['a:lstStyle']?.['a:defRPr']?.['a:solidFill']
  if (lstStyleDefRPrSolidFill) {
    return resolveSolidFill(lstStyleDefRPrSolidFill, context)
  }

  return undefined
}

/**
 * 解析文本体
 *
 * @param txBody - 文本体节点
 * @param context - 解析上下文
 * @returns 段落数组
 */
export function parseTextBodyToParagraphs(
  txBody: XmlObject | undefined,
  context: ParsingContext
): PPTXParagraph[] {
  if (!txBody) return []

  const pArray = txBody['a:p'] || []
  const paragraphs = Array.isArray(pArray) ? pArray : [pArray]

  return paragraphs.map((p: XmlObject) => {
    const runs: any[] = []

    // 解析文本运行
    const rArray = p?.['a:r'] || []
    const runArray = Array.isArray(rArray) ? rArray : [rArray]

    for (const r of runArray) {
      const run = r as XmlObject
      const rPr = run?.['a:rPr']?.['attrs'] || {}
      const text = run?.['a:t'] || ''

      // 解析颜色（支持多层级继承）
      const color = resolveTextColor(run, p, txBody, context)

      runs.push({
        text: String(text),
        bold: rPr['b'] === '1',
        italic: rPr['i'] === '1',
        underline: rPr['u'] === 'sng' || rPr['u'] === '1',
        strike: rPr['strike'] === 'sngStrike',
        fontSize: rPr['sz'] ? parseInt(rPr['sz'], 10) / 100 : undefined,
        fontName: rPr['latin'] as string | undefined,
        color,
      })
    }

    // 解析段落属性
    const pPr = p?.['a:pPr']?.['attrs']
    const algn = pPr?.['algn']

    let align: 'left' | 'center' | 'right' | 'justify' | undefined
    switch (algn) {
      case 'l':
        align = 'left'
        break
      case 'r':
        align = 'right'
        break
      case 'ctr':
        align = 'center'
        break
      case 'just':
      case 'dist':
        align = 'justify'
        break
    }

    const bullet = !!(p?.['a:pPr']?.['a:buFont'] || p?.['a:pPr']?.['a:buChar'] || p?.['a:pPr']?.['a:buAutoNum'])
    const level = pPr?.['lvl'] ? parseInt(pPr['lvl'], 10) : undefined

    return { runs, align, bullet, level }
  })
}

/**
 * 解析形状元素
 *
 * @description
 * 从 PPTX 的 spTree 中提取形状信息，包括：
 * - 位置和尺寸 (transform)
 * - 填充样式 (fill)
 * - 边框样式 (outline)
 * - 文本内容 (textBody)
 *
 * @param shape - 形状 XML 节点
 * @param context - 解析上下文
 * @returns 解析后的 PPTX 元素，如果解析失败返回 null
 */
export function parseShape(
  shape: XmlObject,
  context: ParsingContext
): PPTXElement | null {
  const nvSpPr = shape['p:nvSpPr']
  const spPr = shape['p:spPr']
  const txBody = shape['p:txBody']

  const transform = parseTransform(spPr)
  const cNvPr = nvSpPr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  // 检查是否有实际文本内容
  let paragraphs: PPTXParagraph[] = []
  let hasActualText = false

  if (txBody) {
    paragraphs = parseTextBodyToParagraphs(txBody, context)
    hasActualText = paragraphs.some(p => p.runs.some(run => run.text && run.text.trim().length > 0))
  }

  // 如果有实际文本内容，返回文本元素
  if (hasActualText) {
    return {
      type: 'text',
      id: String(id),
      transform,
      name,
      paragraphs,
    }
  }

  // 形状
  const prstGeom = spPr?.['a:prstGeom']?.['attrs']?.['prst'] as string
  const shapeType = prstGeom || 'rect'

  // 解析 avLst 中的调整值
  const avLst = spPr?.['a:prstGeom']?.['a:avLst']
  const gdList = avLst?.['a:gd']
  const gdArray = Array.isArray(gdList) ? gdList : gdList ? [gdList] : []

  let adj: number | undefined
  for (const gd of gdArray) {
    const gdAttrs = gd?.['attrs']
    if (gdAttrs?.['name'] === 'adj') {
      const fmla = gdAttrs['fmla'] as string
      if (fmla?.startsWith('val ')) {
        adj = parseInt(fmla.substring(4), 10)
      }
      break
    }
  }

  // 解析填充（带透明度）
  let fill: string | undefined
  let fillOpacity: number | undefined
  const solidFill = spPr?.['a:solidFill']
  if (solidFill) {
    const fillResult = resolveSolidFillWithAlpha(solidFill, context)
    fill = fillResult.color
    fillOpacity = fillResult.alpha
  }

  // 解析轮廓
  let outline: { color?: string; width?: number; style?: 'solid' | 'dashed' | 'dotted' } | undefined
  const ln = spPr?.['a:ln']
  if (ln) {
    const lnAttrs = ln?.['attrs'] || {}
    const lnSolidFill = ln['a:solidFill']
    outline = {
      color: lnSolidFill ? resolveSolidFill(lnSolidFill, context) : undefined,
      width: lnAttrs['w'] ? parseInt(lnAttrs['w'], 10) * RATIO_EMUs_Points : undefined,
      style: ln?.['a:prstDash']?.['attrs']?.['val'] === 'solid' ? 'solid' :
             ln?.['a:prstDash']?.['attrs']?.['val'] ? 'dashed' : 'solid',
    }
  }

  return {
    type: 'shape',
    id: String(id),
    transform,
    name,
    shapeType,
    adj,
    fill,
    fillOpacity,
    outline,
    path: generateShapePath(shapeType, transform.width * RATIO_EMUs_Points, transform.height * RATIO_EMUs_Points, adj),
  }
}

export default { parseShape, parseTransform, parseTextBodyToParagraphs }
