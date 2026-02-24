/**
 * 错误类单元测试
 */

import { describe, it, expect } from 'vitest'
import { ConversionError, ConversionWarning, Errors, Warnings } from '../../src/utils/errors.js'

describe('ConversionError', () => {
  describe('constructor', () => {
    it('should create error with code and message', () => {
      const error = new ConversionError('ERR_INVALID_FORMAT', 'Invalid file format')

      expect(error.code).toBe('ERR_INVALID_FORMAT')
      expect(error.message).toBe('Invalid file format')
      expect(error.name).toBe('ConversionError')
    })

    it('should create error with suggestion', () => {
      const error = new ConversionError(
        'ERR_FILE_TOO_LARGE',
        'File too large',
        'Please upload a smaller file'
      )

      expect(error.suggestion).toBe('Please upload a smaller file')
    })
  })

  describe('toJSON', () => {
    it('should return error response format', () => {
      const error = new ConversionError(
        'ERR_INVALID_FORMAT',
        'Invalid format',
        'Upload a PPTX file'
      )

      const json = error.toJSON()

      expect(json.success).toBe(false)
      expect(json.error.code).toBe('ERR_INVALID_FORMAT')
      expect(json.error.message).toBe('Invalid format')
      expect(json.error.suggestion).toBe('Upload a PPTX file')
    })
  })

  describe('getStatusCode', () => {
    it('should return 400 for ERR_INVALID_FORMAT', () => {
      const error = new ConversionError('ERR_INVALID_FORMAT', 'Test')
      expect(error.getStatusCode()).toBe(400)
    })

    it('should return 413 for ERR_FILE_TOO_LARGE', () => {
      const error = new ConversionError('ERR_FILE_TOO_LARGE', 'Test')
      expect(error.getStatusCode()).toBe(413)
    })

    it('should return 400 for ERR_PROTECTED_FILE', () => {
      const error = new ConversionError('ERR_PROTECTED_FILE', 'Test')
      expect(error.getStatusCode()).toBe(400)
    })

    it('should return 400 for ERR_CORRUPTED_FILE', () => {
      const error = new ConversionError('ERR_CORRUPTED_FILE', 'Test')
      expect(error.getStatusCode()).toBe(400)
    })

    it('should return 400 for ERR_EMPTY_FILE', () => {
      const error = new ConversionError('ERR_EMPTY_FILE', 'Test')
      expect(error.getStatusCode()).toBe(400)
    })

    it('should return 500 for ERR_CONVERSION_FAILED', () => {
      const error = new ConversionError('ERR_CONVERSION_FAILED', 'Test')
      expect(error.getStatusCode()).toBe(500)
    })
  })
})

describe('ConversionWarning', () => {
  describe('constructor', () => {
    it('should create warning with code and message', () => {
      const warning = new ConversionWarning('WARN_SMARTART_SKIPPED', 'SmartArt skipped')

      expect(warning.code).toBe('WARN_SMARTART_SKIPPED')
      expect(warning.message).toBe('SmartArt skipped')
    })

    it('should create warning with count', () => {
      const warning = new ConversionWarning('WARN_MACRO_SKIPPED', 'Macros skipped', 5)

      expect(warning.count).toBe(5)
    })
  })

  describe('toString', () => {
    it('should return formatted string without count', () => {
      const warning = new ConversionWarning('WARN_TEST', 'Test warning')

      expect(warning.toString()).toBe('WARN_TEST: Test warning')
    })

    it('should return formatted string with count', () => {
      const warning = new ConversionWarning('WARN_TEST', 'Test warning', 3)

      expect(warning.toString()).toBe('WARN_TEST: Test warning (3 instances)')
    })
  })

  describe('toInfo', () => {
    it('should return WarningInfo object', () => {
      const warning = new ConversionWarning('WARN_TEST', 'Test warning', 2)

      const info = warning.toInfo()

      expect(info.code).toBe('WARN_TEST')
      expect(info.message).toBe('Test warning')
      expect(info.count).toBe(2)
    })
  })
})

describe('Errors factory', () => {
  describe('invalidFormat', () => {
    it('should create ERR_INVALID_FORMAT error', () => {
      const error = Errors.invalidFormat()

      expect(error.code).toBe('ERR_INVALID_FORMAT')
      expect(error.getStatusCode()).toBe(400)
    })

    it('should accept custom detail message', () => {
      const error = Errors.invalidFormat('Custom error message')

      expect(error.message).toBe('Custom error message')
    })
  })

  describe('fileTooLarge', () => {
    it('should create ERR_FILE_TOO_LARGE error with default size', () => {
      const error = Errors.fileTooLarge()

      expect(error.code).toBe('ERR_FILE_TOO_LARGE')
      expect(error.message).toContain('50MB')
      expect(error.getStatusCode()).toBe(413)
    })

    it('should accept custom max size', () => {
      const error = Errors.fileTooLarge('100MB')

      expect(error.message).toContain('100MB')
    })
  })

  describe('protectedFile', () => {
    it('should create ERR_PROTECTED_FILE error', () => {
      const error = Errors.protectedFile()

      expect(error.code).toBe('ERR_PROTECTED_FILE')
      expect(error.getStatusCode()).toBe(400)
    })
  })

  describe('corruptedFile', () => {
    it('should create ERR_CORRUPTED_FILE error', () => {
      const error = Errors.corruptedFile()

      expect(error.code).toBe('ERR_CORRUPTED_FILE')
      expect(error.getStatusCode()).toBe(400)
    })
  })

  describe('emptyFile', () => {
    it('should create ERR_EMPTY_FILE error', () => {
      const error = Errors.emptyFile()

      expect(error.code).toBe('ERR_EMPTY_FILE')
      expect(error.getStatusCode()).toBe(400)
    })
  })

  describe('conversionFailed', () => {
    it('should create ERR_CONVERSION_FAILED error with default message', () => {
      const error = Errors.conversionFailed()

      expect(error.code).toBe('ERR_CONVERSION_FAILED')
      expect(error.getStatusCode()).toBe(500)
    })

    it('should accept custom detail message', () => {
      const error = Errors.conversionFailed('Custom failure reason')

      expect(error.message).toBe('Custom failure reason')
    })
  })
})

describe('Warnings factory', () => {
  describe('smartArtSkipped', () => {
    it('should create WARN_SMARTART_SKIPPED warning', () => {
      const warning = Warnings.smartArtSkipped()

      expect(warning.code).toBe('WARN_SMARTART_SKIPPED')
    })

    it('should accept count', () => {
      const warning = Warnings.smartArtSkipped(5)

      expect(warning.count).toBe(5)
    })
  })

  describe('macroSkipped', () => {
    it('should create WARN_MACRO_SKIPPED warning', () => {
      const warning = Warnings.macroSkipped()

      expect(warning.code).toBe('WARN_MACRO_SKIPPED')
    })
  })

  describe('activeXSkipped', () => {
    it('should create WARN_ACTIVEX_SKIPPED warning', () => {
      const warning = Warnings.activeXSkipped()

      expect(warning.code).toBe('WARN_ACTIVEX_SKIPPED')
    })
  })

  describe('fontFallback', () => {
    it('should create WARN_FONT_FALLBACK warning', () => {
      const warning = Warnings.fontFallback()

      expect(warning.code).toBe('WARN_FONT_FALLBACK')
    })
  })
})
