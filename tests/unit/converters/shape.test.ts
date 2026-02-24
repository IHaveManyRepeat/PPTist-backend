/**
 * 形状转换器单元测试
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest'
import shapeModule from '../../../src/modules/conversion/converters/shape.js'
import { clearConverters, getConverter, getRegisteredConverters } from '../../../src/modules/conversion/converters/index.js'
import type { PPTXShapeElement } from '../../../src/modules/conversion/types/pptx.js'
import type { ConversionContext } from '../../../src/types/index.js'

// 从默认导出中获取函数
const { isShapeElement, convertShape, registerShapeConverter } = shapeModule as any

// 创建测试用的转换上下文
const createTestContext = (): ConversionContext => ({
  requestId: 'test-request-id',
  startTime: Date.now(),
  warnings: [],
  mediaMap: new Map(),
  slideSize: { width: 9144000, height: 6858000 },
  currentSlideIndex: 0,
})

// 创建测试用的形状元素
const createTestShapeElement = (overrides: Partial<PPTXShapeElement> = {}): PPTXShapeElement => ({
  type: 'shape',
  id: 'shape-1',
  transform: {
    x: 914400,   // 1 inch in EMUs
    y: 685800,   // 0.75 inch in EMUs
    width: 1828800, // 2 inches
    height: 1371600, // 1.5 inches
    rotation: 0,
  },
  shapeType: 'rect',
  fill: '#4472C4',
  outline: {
    color: '#000000',
    width: 1,
    style: 'solid',
  },
  ...overrides,
})

describe('Shape Converter', () => {
  beforeEach(() => {
    clearConverters()
  })

  afterEach(() => {
    clearConverters()
  })

  describe('isShapeElement', () => {
    it('should return true for shape element', () => {
      const element = createTestShapeElement()
      expect(isShapeElement(element)).toBe(true)
    })

    it('should return false for non-shape element', () => {
      const element = { type: 'image', id: 'img-1' }
      expect(isShapeElement(element)).toBe(false)
    })

    it('should return false for text element', () => {
      const element = { type: 'text', id: 'text-1' }
      expect(isShapeElement(element)).toBe(false)
    })
  })

  describe('convertShape', () => {
    it('should convert basic shape element', () => {
      const element = createTestShapeElement()
      const context = createTestContext()

      const result = convertShape(element, context)

      expect(result.type).toBe('shape')
      expect(result.left).toBeGreaterThan(0)
      expect(result.top).toBeGreaterThan(0)
      expect(result.width).toBeGreaterThan(0)
      expect(result.height).toBeGreaterThan(0)
    })

    it('should convert transform values to pixels', () => {
      const element = createTestShapeElement({
        transform: {
          x: 914400,   // 1 inch = 96px at 96 DPI
          y: 685800,
          width: 1828800,
          height: 1371600,
          rotation: 0,
        },
      })
      const context = createTestContext()

      const result = convertShape(element, context)

      // EMU 到像素的转换
      // 1 inch = 914400 EMUs, 1 inch = 96 pixels (at 96 DPI)
      // 所以 x = 914400 / 914400 * 96 = 96 pixels
      expect(result.left).toBeCloseTo(96, 0)
      expect(result.width).toBeCloseTo(192, 0)
    })

    it('should handle rotation', () => {
      const element = createTestShapeElement({
        transform: {
          ...createTestShapeElement().transform,
          rotation: 45,
        },
      })
      const context = createTestContext()

      const result = convertShape(element, context)

      expect(result.rotate).toBe(45)
    })

    it('should generate shape path', () => {
      const element = createTestShapeElement({ shapeType: 'rect' })
      const context = createTestContext()

      const result = convertShape(element, context)

      expect(result.path).toBeDefined()
      expect(result.path).toContain('M')
      expect(result.path).toContain('Z')
    })

    it('should handle outline', () => {
      const element = createTestShapeElement({
        outline: {
          color: '#FF0000',
          width: 2,
          style: 'dashed',
        },
      })
      const context = createTestContext()

      const result = convertShape(element, context)

      expect(result.outline.color).toBe('#FF0000')
      expect(result.outline.width).toBe(2)
      expect(result.outline.style).toBe('dashed')
    })

    it('should merge fillOpacity into fill color', () => {
      const element = createTestShapeElement({
        fill: '4472C4',
        fillOpacity: 0.5,
      })
      const context = createTestContext()

      const result = convertShape(element, context)

      // 50% opacity = 80 in hex
      expect(result.fill).toMatch(/4472C480$/i)
    })

    it('should handle paragraphs (shape with text)', () => {
      const element = createTestShapeElement({
        paragraphs: [
          {
            runs: [{ text: 'Hello', bold: true }],
            align: 'center',
            bullet: false,
          },
        ],
      })
      const context = createTestContext()

      const result = convertShape(element, context)

      expect(result.text).toBeDefined()
      expect(result.text?.content).toContain('Hello')
      expect(result.text?.content).toContain('<b>')
    })
  })

  describe('registerShapeConverter', () => {
    it('should register shape converter', () => {
      registerShapeConverter()

      const element = createTestShapeElement()
      const converter = getConverter(element)

      expect(converter).toBeDefined()
    })

    it('should have correct priority (5)', () => {
      registerShapeConverter()

      const element = createTestShapeElement()
      const converters = getRegisteredConverters()
      const shapeConverter = converters.find(c => c.detector(element))

      expect(shapeConverter?.priority).toBe(5)
    })
  })
})
