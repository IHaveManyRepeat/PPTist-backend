import { registerConverter } from './index.js'
import type { PPTXLineElement, PPTXElement } from '../types/pptx.js'
import type { PPTShapeElement } from '../types/pptist.js'
import type { ConversionContext } from '../../../types/index.js'
import { v4 as uuidv4 } from 'uuid'
import { createEmuConverters } from '../utils/geometry.js'

/**
 * Detect if element is a line element
 */
function isLineElement(element: PPTXElement): element is PPTXLineElement {
  return element.type === 'line'
}

/**
 * Convert PPTX line element to PPTist shape element
 * Lines are represented as shapes with a path and transparent fill
 */
function convertLineToShape(element: PPTXLineElement, context: ConversionContext): PPTShapeElement {
  const { transform, startX, startY, endX, endY, color, style, width } = element
  const { toPixelX, toPixelY } = createEmuConverters(context.slideSize)

  // 计算元素边界（像素坐标）
  const left = toPixelX(transform.x)
  const top = toPixelY(transform.y)
  const w = toPixelX(transform.width)
  const h = toPixelY(transform.height)

  // 计算起点终点相对于元素边界的偏移（在 1000x1000 viewBox 中）
  // 防止除零错误
  const safeW = w > 0 ? w : 1
  const safeH = h > 0 ? h : 1

  const relStartX = ((toPixelX(startX) - left) / safeW) * 1000
  const relStartY = ((toPixelY(startY) - top) / safeH) * 1000
  const relEndX = ((toPixelX(endX) - left) / safeW) * 1000
  const relEndY = ((toPixelY(endY) - top) / safeH) * 1000

  // 生成线条的 SVG path
  const path = `M ${relStartX.toFixed(2)} ${relStartY.toFixed(2)} L ${relEndX.toFixed(2)} ${relEndY.toFixed(2)}`

  const pptistShape: PPTShapeElement = {
    id: uuidv4(),
    type: 'shape',
    left,
    top,
    width: w,
    height: h,
    rotate: transform.rotation || 0,
    viewBox: [1000, 1000],
    path,
    fixedRatio: false,
    fill: 'transparent', // 线条无填充
    outline: {
      style: style || 'solid',
      width: width || 1,
      color: color || '#000000',
    },
    opacity: 1,
  }

  return pptistShape
}

/**
 * Register line converter
 */
export function registerLineConverter(): void {
  registerConverter(
    (element, context) => convertLineToShape(element as PPTXLineElement, context),
    isLineElement,
    5
  )
}

export default { registerLineConverter, convertLineToShape, isLineElement }
