/**
 * 转换器注册系统单元测试
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import {
  registerConverter,
  getConverter,
  convertElement,
  getRegisteredConverters,
  clearConverters,
} from '../../../src/modules/conversion/converters/index.js'
import type { PPTXElement } from '../../../src/modules/conversion/types/pptx.js'
import type { PPTElement } from '../../../src/modules/conversion/types/pptist.js'
import type { ConversionContext } from '../../../src/types/index.js'

// 创建测试用的转换上下文
const createTestContext = (): ConversionContext => ({
  requestId: 'test-request-id',
  startTime: Date.now(),
  warnings: [],
  mediaMap: new Map(),
  slideSize: { width: 9144000, height: 6858000 },
  currentSlideIndex: 0,
})

describe('Converter Registry', () => {
  beforeEach(() => {
    clearConverters()
  })

  afterEach(() => {
    clearConverters()
  })

  describe('registerConverter', () => {
    it('should register a converter', () => {
      const converter = vi.fn()
      const detector = () => true

      registerConverter(converter, detector, 0)

      const converters = getRegisteredConverters()
      expect(converters).toHaveLength(1)
    })

    it('should sort converters by priority (descending)', () => {
      const converter1 = vi.fn()
      const converter2 = vi.fn()
      const detector = () => true

      registerConverter(converter1, detector, 1)
      registerConverter(converter2, detector, 5)

      const converters = getRegisteredConverters()
      expect(converters[0].priority).toBe(5)
      expect(converters[1].priority).toBe(1)
    })
  })

  describe('getConverter', () => {
    it('should return null when no converter matches', () => {
      const element = { type: 'unknown' } as PPTXElement
      expect(getConverter(element)).toBeNull()
    })

    it('should return matching converter', () => {
      const converter = vi.fn().mockReturnValue({ type: 'shape' })
      const detector = (el: PPTXElement) => el.type === 'shape'

      registerConverter(converter as any, detector, 0)

      const element = { type: 'shape' } as PPTXElement
      const result = getConverter(element)

      expect(result).toBe(converter)
    })

    it('should return highest priority matching converter', () => {
      const converter1 = vi.fn().mockReturnValue({ type: 'shape', version: 1 })
      const converter2 = vi.fn().mockReturnValue({ type: 'shape', version: 2 })
      const detector = (el: PPTXElement) => el.type === 'shape'

      registerConverter(converter1 as any, detector, 1)
      registerConverter(converter2 as any, detector, 5)

      const element = { type: 'shape' } as PPTXElement
      const result = getConverter(element)

      expect(result).toBe(converter2)
    })
  })

  describe('convertElement', () => {
    it('should return null when no converter matches', () => {
      const element = { type: 'unknown' } as PPTXElement
      const context = createTestContext()

      const result = convertElement(element, context)

      expect(result).toBeNull()
    })

    it('should call matching converter', () => {
      const mockResult: PPTElement = {
        id: 'test-id',
        type: 'shape',
        left: 0,
        top: 0,
        width: 100,
        height: 100,
        rotate: 0,
        viewBox: [100, 100],
        path: 'M0,0 L100,0 L100,100 L0,100 Z',
        fixedRatio: false,
        fill: '#FFFFFF',
        outline: { style: 'solid', width: 1, color: '#000000' },
        opacity: 1,
      }

      const converter = vi.fn().mockReturnValue(mockResult)
      const detector = (el: PPTXElement) => el.type === 'shape'

      registerConverter(converter as any, detector, 0)

      const element = { type: 'shape' } as PPTXElement
      const context = createTestContext()

      const result = convertElement(element, context)

      expect(result).toEqual(mockResult)
      expect(converter).toHaveBeenCalledWith(element, context)
    })

    it('should return array of elements when converter returns array', () => {
      const mockResults: PPTElement[] = [
        { id: 'id-1', type: 'shape', left: 0, top: 0, width: 50, height: 50, rotate: 0 } as PPTElement,
        { id: 'id-2', type: 'shape', left: 50, top: 50, width: 50, height: 50, rotate: 0 } as PPTElement,
      ]

      const converter = vi.fn().mockReturnValue(mockResults)
      const detector = (el: PPTXElement) => el.type === 'group'

      registerConverter(converter as any, detector, 0)

      const element = { type: 'group' } as PPTXElement
      const context = createTestContext()

      const result = convertElement(element, context)

      expect(Array.isArray(result)).toBe(true)
      expect(result).toHaveLength(2)
    })
  })

  describe('clearConverters', () => {
    it('should clear all registered converters', () => {
      const converter = vi.fn()
      const detector = () => true

      registerConverter(converter, detector, 0)
      expect(getRegisteredConverters()).toHaveLength(1)

      clearConverters()
      expect(getRegisteredConverters()).toHaveLength(0)
    })
  })
})
