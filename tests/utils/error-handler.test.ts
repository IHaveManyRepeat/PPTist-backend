/**
 * 错误处理器单元测试
 */

import { describe, it, expect, beforeEach, vi } from 'vitest'
import { ConversionErrorHandler, createErrorHandler } from '../../src/utils/error-handler.js'
import type { Logger } from 'pino'

// 创建模拟 logger
const mockLogger = {
  warn: vi.fn(),
  debug: vi.fn(),
  error: vi.fn(),
  info: vi.fn(),
  child: vi.fn().mockReturnThis(),
} as unknown as Logger

describe('ConversionErrorHandler', () => {
  let errorHandler: ConversionErrorHandler

  beforeEach(() => {
    vi.clearAllMocks()
    errorHandler = new ConversionErrorHandler('test-request-id', mockLogger)
  })

  describe('constructor', () => {
    it('should create instance with requestId', () => {
      expect(errorHandler).toBeInstanceOf(ConversionErrorHandler)
    })

    it('should create instance without logger (use default)', () => {
      const handler = new ConversionErrorHandler('test-id')
      expect(handler).toBeInstanceOf(ConversionErrorHandler)
    })
  })

  describe('handleElementError', () => {
    it('should log warning when element conversion fails', () => {
      const element = { id: 'shape-1', type: 'shape' }
      const error = new Error('Conversion failed')

      errorHandler.handleElementError(element, error, {
        operation: 'conversion',
        slideIndex: 0,
      })

      expect(mockLogger.warn).toHaveBeenCalled()
    })

    it('should return null by default', () => {
      const element = { id: 'shape-1', type: 'shape' }
      const error = new Error('Conversion failed')

      const result = errorHandler.handleElementError(element, error)

      expect(result).toBeNull()
    })

    it('should return default value when provided', () => {
      const element = { id: 'shape-1', type: 'shape' }
      const error = new Error('Conversion failed')
      const defaultValue = { fallback: true }

      const result = errorHandler.handleElementError(element, error, {
        defaultValue,
      })

      expect(result).toEqual(defaultValue)
    })

    it('should throw error when throw option is true', () => {
      const element = { id: 'shape-1', type: 'shape' }
      const error = new Error('Critical error')

      expect(() => {
        errorHandler.handleElementError(element, error, { throw: true })
      }).toThrow('Critical error')
    })

    it('should add warning to warnings list', () => {
      const element = { id: 'shape-1', type: 'shape' }
      const error = new Error('Conversion failed')

      errorHandler.handleElementError(element, error, {
        operation: 'conversion',
        slideIndex: 2,
      })

      const warnings = errorHandler.getWarnings()
      expect(warnings).toHaveLength(1)
      expect(warnings[0].code).toBe('WARN_ELEMENT_FAILED')
    })
  })

  describe('addWarning', () => {
    it('should add warning to list', () => {
      errorHandler.addWarning('WARN_SMARTART_SKIPPED', 'SmartArt was skipped', {
        elementId: 'smart-1',
        slideIndex: 0,
      })

      expect(errorHandler.hasWarnings()).toBe(true)
      expect(errorHandler.getWarningCount()).toBe(1)
    })

    it('should group warnings by code', () => {
      errorHandler.addWarning('WARN_SMARTART_SKIPPED', 'SmartArt 1 skipped')
      errorHandler.addWarning('WARN_SMARTART_SKIPPED', 'SmartArt 2 skipped')
      errorHandler.addWarning('WARN_MACRO_SKIPPED', 'Macro skipped')

      const warnings = errorHandler.getWarnings()
      expect(warnings).toHaveLength(2)

      const smartArtWarning = warnings.find(w => w.code === 'WARN_SMARTART_SKIPPED')
      expect(smartArtWarning?.count).toBe(2)

      const macroWarning = warnings.find(w => w.code === 'WARN_MACRO_SKIPPED')
      expect(macroWarning?.count).toBe(1)
    })
  })

  describe('getWarnings', () => {
    it('should return empty array when no warnings', () => {
      expect(errorHandler.getWarnings()).toEqual([])
    })

    it('should return grouped warnings with counts', () => {
      errorHandler.addWarning('WARN_CODE_1', 'Warning 1')
      errorHandler.addWarning('WARN_CODE_1', 'Warning 1 again')
      errorHandler.addWarning('WARN_CODE_2', 'Warning 2')

      const warnings = errorHandler.getWarnings()
      expect(warnings).toHaveLength(2)
    })
  })

  describe('getWarningDetails', () => {
    it('should return all warnings without grouping', () => {
      errorHandler.addWarning('WARN_CODE_1', 'Warning 1')
      errorHandler.addWarning('WARN_CODE_1', 'Warning 1 again')

      const details = errorHandler.getWarningDetails()
      expect(details).toHaveLength(2)
    })
  })

  describe('hasWarnings', () => {
    it('should return false when no warnings', () => {
      expect(errorHandler.hasWarnings()).toBe(false)
    })

    it('should return true when warnings exist', () => {
      errorHandler.addWarning('WARN_TEST', 'Test warning')
      expect(errorHandler.hasWarnings()).toBe(true)
    })
  })

  describe('clearWarnings', () => {
    it('should clear all warnings', () => {
      errorHandler.addWarning('WARN_TEST', 'Test warning')
      expect(errorHandler.hasWarnings()).toBe(true)

      errorHandler.clearWarnings()
      expect(errorHandler.hasWarnings()).toBe(false)
    })
  })

  describe('createChildProcessor', () => {
    it('should create child processor with context', () => {
      const childProcessor = errorHandler.createChildProcessor('slide-0')
      expect(childProcessor).toBeInstanceOf(ConversionErrorHandler)
    })

    it('child processor warnings should propagate to parent', () => {
      const childProcessor = errorHandler.createChildProcessor('slide-0')

      childProcessor.addWarning('WARN_CHILD', 'Child warning')

      // 父处理器也应该有警告
      expect(errorHandler.hasWarnings()).toBe(true)
    })
  })
})

describe('createErrorHandler', () => {
  it('should create error handler instance', () => {
    const handler = createErrorHandler('test-id')
    expect(handler).toBeInstanceOf(ConversionErrorHandler)
  })

  it('should create error handler with custom logger', () => {
    const handler = createErrorHandler('test-id', mockLogger)
    expect(handler).toBeInstanceOf(ConversionErrorHandler)
  })
})
