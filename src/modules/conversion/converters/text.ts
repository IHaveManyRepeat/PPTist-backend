import { registerConverter } from './index.js'
import type { PPTXTextElement, PPTXElement } from '../types/pptx.js'
import type { PPTTextElement } from '../types/pptist.js'
import type { ConversionContext } from '../../../types/index.js'
import { v4 as uuidv4 } from 'uuid'
import { createEmuConverters } from '../utils/geometry.js'

/**
 * Convert PPTX text runs to HTML content
 */
function runsToHtml(paragraphs: PPTXTextElement['paragraphs']): string {
  return paragraphs
    .map((p) => {
      const runsHtml = p.runs
        .map((run) => {
          let html = run.text
          if (run.bold) html = `<b>${html}</b>`
          if (run.italic) html = `<i>${html}</i>`
          if (run.underline) html = `<u>${html}</u>`
          if (run.strike) html = `<s>${html}</s>`
          if (run.fontSize) {
            html = `<span style="font-size: ${run.fontSize}px">${html}</span>`
          }
          if (run.color) {
            html = `<span style="color: ${run.color}">${html}</span>`
          }
          if (run.fontName) {
            html = `<span style="font-family: ${run.fontName}">${html}</span>`
          }
          return html
        })
        .join('')

      // Wrap in paragraph with alignment
      const align = p.align || 'left'
      const bulletStyle = p.bullet ? 'list-style-type: disc; margin-left: 20px;' : ''
      return `<p style="text-align: ${align}; ${bulletStyle}">${runsHtml}</p>`
    })
    .join('')
}

/**
 * Detect if element is a text element
 */
function isTextElement(element: PPTXElement): element is PPTXTextElement {
  return element.type === 'text'
}

/**
 * Convert PPTX text element to PPTist text element
 */
function convertText(element: PPTXTextElement, context: ConversionContext): PPTTextElement {
  const { transform, paragraphs } = element
  const { toPixelX, toPixelY } = createEmuConverters(context.slideSize)

  // Get default formatting from first paragraph's first run
  const firstRun = paragraphs[0]?.runs[0]
  const defaultFontName = firstRun?.fontName || 'Arial'
  const defaultColor = firstRun?.color || '#000000'

  const pptistText: PPTTextElement = {
    id: uuidv4(),
    type: 'text',
    left: toPixelX(transform.x),
    top: toPixelY(transform.y),
    width: toPixelX(transform.width),
    height: toPixelY(transform.height),
    rotate: transform.rotation || 0,
    content: runsToHtml(paragraphs),
    defaultFontName,
    defaultColor,
    lineHeight: 1.5,
    wordSpace: 0,
    paragraphSpace: 5,
  }

  return pptistText
}

/**
 * Register text converter
 */
export function registerTextConverter(): void {
  registerConverter(
    (element, context) => convertText(element as PPTXTextElement, context),
    isTextElement,
    10 // High priority for text elements
  )
}

export default { registerTextConverter, convertText, isTextElement }
