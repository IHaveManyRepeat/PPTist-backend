import { registerConverter } from './index.js'
import type { PPTXShapeElement, PPTXElement } from '../types/pptx.js'
import type { PPTShapeElement, ShapeText } from '../types/pptist.js'
import type { ConversionContext } from '../../../types/index.js'
import { v4 as uuidv4 } from 'uuid'
import { createEmuConverters } from '../utils/geometry.js'
import { generateShapePath } from '../generators/svg-path-generator.js'

/**
 * Convert PPTX paragraphs to shape text
 */
function convertShapeText(element: PPTXShapeElement): ShapeText | undefined {
  if (!element.paragraphs || element.paragraphs.length === 0) {
    return undefined
  }

  const firstRun = element.paragraphs[0]?.runs[0]
  const runsHtml = element.paragraphs
    .map((p) => {
      const runHtml = p.runs
        .map((run) => {
          let html = run.text
          if (run.bold) html = `<b>${html}</b>`
          if (run.italic) html = `<i>${html}</i>`
          if (run.color) html = `<span style="color: ${run.color}">${html}</span>`
          return html
        })
        .join('')
      return `<p>${runHtml}</p>`
    })
    .join('')

  return {
    content: runsHtml,
    defaultFontName: firstRun?.fontName || 'Arial',
    defaultColor: firstRun?.color || '#000000',
    align: 'middle',
    lineHeight: 1.5,
  }
}

/**
 * Detect if element is a shape element
 */
function isShapeElement(element: PPTXElement): element is PPTXShapeElement {
  return element.type === 'shape'
}

/**
 * Convert PPTX shape element to PPTist shape element
 */
function convertShape(element: PPTXShapeElement, context: ConversionContext): PPTShapeElement {
  const { transform, shapeType, adj, fill, fillOpacity, outline, paragraphs } = element
  const { toPixelX, toPixelY } = createEmuConverters(context.slideSize)

  // 计算像素尺寸
  const pixelWidth = toPixelX(transform.width)
  const pixelHeight = toPixelY(transform.height)

  // 使用像素尺寸重新生成路径（确保路径坐标与元素尺寸匹配）
  // 注意：不使用 parser 传递的 path，因为它使用的是点坐标而非像素坐标
  const shapePath = shapeType
    ? generateShapePath(shapeType, pixelWidth, pixelHeight, adj)
    : `M0,0 L${pixelWidth},0 L${pixelWidth},${pixelHeight} L0,${pixelHeight} Z`

  // viewBox 使用像素尺寸（与元素尺寸一致）
  const viewBox: [number, number] = [pixelWidth, pixelHeight]

  // 合并透明度到填充色中（8位hex格式 #RRGGBBAA）
  let finalFill = fill || 'FFFFFF'
  if (fillOpacity !== undefined && finalFill) {
    // 移除 # 前缀（如果有）
    const colorHex = finalFill.replace('#', '')
    // 将 0-1 的透明度转换为 00-FF
    const alphaHex = Math.round(fillOpacity * 255).toString(16).padStart(2, '0').toUpperCase()
    finalFill = `#${colorHex}${alphaHex}`
  } else if (finalFill && !finalFill.startsWith('#')) {
    // 如果没有透明度且没有 # 前缀，添加 #
    finalFill = `#${finalFill}`
  }

  const pptistShape: PPTShapeElement = {
    id: uuidv4(),
    type: 'shape',
    left: toPixelX(transform.x),
    top: toPixelY(transform.y),
    width: toPixelX(transform.width),
    height: toPixelY(transform.height),
    rotate: transform.rotation || 0,
    viewBox: viewBox,
    path: shapePath,
    fixedRatio: false,
    fill: finalFill,
    // fillOpacity 保留为可选字段（向后兼容），但不再使用
    outline: outline
      ? {
          style: outline.style || 'solid',
          width: outline.width || 1,
          color: outline.color || '#000000',
        }
      : { style: 'solid', width: 1, color: '#000000' },
    opacity: 1,
    text: paragraphs ? convertShapeText(element) : undefined,
  }

  return pptistShape
}

/**
 * Register shape converter
 */
export function registerShapeConverter(): void {
  registerConverter(
    (element, context) => convertShape(element as PPTXShapeElement, context),
    isShapeElement,
    5 // Lower priority than text
  )
}

export default { registerShapeConverter, convertShape, isShapeElement }
