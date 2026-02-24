import { registerConverter } from './index.js'
import type { PPTXTableElement, PPTXElement, PPTXTableCell } from '../types/pptx.js'
import type { PPTTableElement, TableCell, TableCellStyle, TextAlign } from '../types/pptist.js'
import type { ConversionContext } from '../../../types/index.js'
import { v4 as uuidv4 } from 'uuid'
import { createEmuConverters } from '../utils/geometry.js'

/**
 * Map alignment string to PPTist format
 */
function mapAlign(align?: string): TextAlign {
  switch (align) {
    case 'left':
    case 'center':
    case 'right':
    case 'justify':
      return align
    default:
      return 'left'
  }
}

/**
 * Convert PPTX table cell to PPTist table cell
 */
function convertCell(cell: PPTXTableCell): TableCell {
  const formatting = cell.formatting
  const style: TableCellStyle | undefined = formatting
    ? {
        bold: formatting.bold,
        em: formatting.italic,
        color: formatting.color,
        backcolor: formatting.backgroundColor,
        fontsize: formatting.fontSize ? `${formatting.fontSize}px` : undefined,
        align: mapAlign(formatting.align),
      }
    : undefined

  return {
    id: uuidv4(),
    colspan: cell.rowSpan || 1,
    rowspan: cell.colSpan || 1,
    text: cell.text,
    style,
  }
}

/**
 * Detect if element is a table element
 */
function isTableElement(element: PPTXElement): element is PPTXTableElement {
  return element.type === 'table'
}

/**
 * Convert PPTX table element to PPTist table element
 */
function convertTable(element: PPTXTableElement, context: ConversionContext): PPTTableElement {
  const { transform, rows } = element
  const { toPixelX, toPixelY } = createEmuConverters(context.slideSize)

  // Convert all cells
  const pptistRows = rows.map((row) => row.map(convertCell))

  // Calculate column widths (equal distribution)
  const colCount = rows[0]?.length || 1
  const colWidths = Array(colCount).fill(1 / colCount)

  const pptistTable: PPTTableElement = {
    id: uuidv4(),
    type: 'table',
    left: toPixelX(transform.x),
    top: toPixelY(transform.y),
    width: toPixelX(transform.width),
    height: toPixelY(transform.height),
    rotate: transform.rotation || 0,
    outline: { style: 'solid', width: 1, color: '#000000' },
    colWidths,
    cellMinHeight: 30,
    data: pptistRows,
  }

  return pptistTable
}

/**
 * Register table converter
 */
export function registerTableConverter(): void {
  registerConverter(
    (element, context) => convertTable(element as PPTXTableElement, context),
    isTableElement,
    5
  )
}

export default { registerTableConverter, convertTable, isTableElement }
