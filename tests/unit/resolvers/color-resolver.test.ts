/**
 * 颜色解析器单元测试
 */

import { describe, it, expect } from 'vitest'
import {
  resolveSolidFill,
  resolveSolidFillWithAlpha,
  resolveColor,
  getTextByPathList,
  getSchemeColorFromTheme,
} from '../../../src/modules/conversion/resolvers/color-resolver.js'
import type { ParsingContext, XmlObject } from '../../../src/modules/conversion/context/parsing-context.js'

// 创建测试用的解析上下文
const createTestParsingContext = (overrides: Partial<ParsingContext> = {}): ParsingContext => ({
  zip: {} as any,
  slideContent: {},
  slideLayoutContent: {},
  slideMasterContent: {},
  themeContent: {
    'a:theme': {
      'a:themeElements': {
        'a:clrScheme': {
          'a:dk1': { 'a:sysClr': { attrs: { lastClr: '000000' } } },
          'a:lt1': { 'a:sysClr': { attrs: { lastClr: 'FFFFFF' } } },
          'a:accent1': { 'a:srgbClr': { attrs: { val: '4472C4' } } },
          'a:accent2': { 'a:srgbClr': { attrs: { val: 'ED7D31' } } },
        },
      },
    },
  },
  themeColors: ['#4472C4', '#ED7D31'],
  slideResObj: {},
  layoutResObj: {},
  masterResObj: {},
  slideLayoutTables: { idTable: {}, idxTable: {}, typeTable: {} },
  slideMasterTables: { idTable: {}, idxTable: {}, typeTable: {} },
  defaultTextStyle: {},
  slideIndex: 0,
  ...overrides,
})

describe('color-resolver', () => {
  describe('getTextByPathList', () => {
    it('should return undefined for undefined object', () => {
      expect(getTextByPathList(undefined, ['a', 'b'])).toBeUndefined()
    })

    it('should return value at path', () => {
      const obj = { a: { b: { c: 'value' } } }
      expect(getTextByPathList(obj as XmlObject, ['a', 'b', 'c'])).toBe('value')
    })

    it('should return undefined for non-existent path', () => {
      const obj = { a: { b: 'value' } }
      expect(getTextByPathList(obj as XmlObject, ['a', 'c'])).toBeUndefined()
    })

    it('should handle nested objects', () => {
      const obj = {
        'p:sld': {
          'p:cSld': {
            'p:spTree': {
              'p:sp': {
                attrs: { id: '123' },
              },
            },
          },
        },
      }
      expect(getTextByPathList(obj as XmlObject, ['p:sld', 'p:cSld', 'p:spTree', 'p:sp', 'attrs', 'id'])).toBe('123')
    })
  })

  describe('resolveSolidFill', () => {
    it('should return empty string for undefined solidFill', () => {
      expect(resolveSolidFill(undefined, createTestParsingContext())).toBe('')
    })

    it('should resolve srgbClr (RGB color)', () => {
      const solidFill = {
        'a:srgbClr': {
          attrs: { val: 'FF0000' },
        },
      }
      const result = resolveSolidFill(solidFill as XmlObject, createTestParsingContext())
      expect(result).toBe('#FF0000')
    })

    it('should resolve prstClr (preset color)', () => {
      const solidFill = {
        'a:prstClr': {
          attrs: { val: 'red' },
        },
      }
      const result = resolveSolidFill(solidFill as XmlObject, createTestParsingContext())
      expect(result).toBe('#FF0000')
    })

    it('should resolve sysClr (system color)', () => {
      const solidFill = {
        'a:sysClr': {
          attrs: { lastClr: '00FF00' },
        },
      }
      const result = resolveSolidFill(solidFill as XmlObject, createTestParsingContext())
      expect(result).toBe('#00FF00')
    })

    it('should resolve schemeClr (theme color)', () => {
      const solidFill = {
        'a:schemeClr': {
          attrs: { val: 'accent1' },
        },
      }
      const result = resolveSolidFill(solidFill as XmlObject, createTestParsingContext())
      expect(result).toBe('#4472C4')
    })

    it('should apply alpha modifier', () => {
      const solidFill = {
        'a:srgbClr': {
          attrs: { val: 'FF0000' },
          'a:alpha': {
            attrs: { val: '50000' }, // 50%
          },
        },
      }
      const result = resolveSolidFill(solidFill as XmlObject, createTestParsingContext())
      // Alpha 会使颜色变成 8 位 hex
      expect(result).toMatch(/#[0-9A-Fa-f]{8}/)
    })

    it('should apply tint modifier', () => {
      const solidFill = {
        'a:srgbClr': {
          attrs: { val: 'FF0000' },
          'a:tint': {
            attrs: { val: '50000' }, // 50% tint
          },
        },
      }
      const result = resolveSolidFill(solidFill as XmlObject, createTestParsingContext())
      expect(result).toMatch(/#[0-9A-Fa-f]{6}/)
    })

    it('should apply shade modifier', () => {
      const solidFill = {
        'a:srgbClr': {
          attrs: { val: 'FF0000' },
          'a:shade': {
            attrs: { val: '50000' }, // 50% shade
          },
        },
      }
      const result = resolveSolidFill(solidFill as XmlObject, createTestParsingContext())
      expect(result).toMatch(/#[0-9A-Fa-f]{6}/)
    })
  })

  describe('resolveSolidFillWithAlpha', () => {
    it('should return empty color for undefined solidFill', () => {
      const result = resolveSolidFillWithAlpha(undefined, createTestParsingContext())
      expect(result.color).toBe('')
      expect(result.alpha).toBeUndefined()
    })

    it('should return color without alpha when no alpha modifier', () => {
      const solidFill = {
        'a:srgbClr': {
          attrs: { val: 'FF0000' },
        },
      }
      const result = resolveSolidFillWithAlpha(solidFill as XmlObject, createTestParsingContext())
      expect(result.color).toBe('#FF0000')
      expect(result.alpha).toBeUndefined()
    })

    it('should extract alpha value separately', () => {
      const solidFill = {
        'a:srgbClr': {
          attrs: { val: 'FF0000' },
          'a:alpha': {
            attrs: { val: '50000' }, // 50%
          },
        },
      }
      const result = resolveSolidFillWithAlpha(solidFill as XmlObject, createTestParsingContext())
      expect(result.alpha).toBe(0.5)
    })
  })

  describe('resolveColor', () => {
    it('should return empty string for undefined node', () => {
      expect(resolveColor(undefined, createTestParsingContext())).toBe('')
    })

    it('should delegate to resolveSolidFill', () => {
      const node = {
        'a:srgbClr': {
          attrs: { val: '00FF00' },
        },
      }
      const result = resolveColor(node as XmlObject, createTestParsingContext())
      expect(result).toBe('#00FF00')
    })
  })

  describe('getSchemeColorFromTheme', () => {
    it('should return empty string for unknown scheme color', () => {
      const context = createTestParsingContext()
      const result = getSchemeColorFromTheme('a:unknown', context)
      expect(result).toBe('')
    })

    it('should resolve accent color from theme', () => {
      const context = createTestParsingContext()
      const result = getSchemeColorFromTheme('a:accent1', context)
      expect(result).toBe('4472C4')
    })
  })
})
