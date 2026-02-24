/**
 * 统一错误处理器
 *
 * @module utils/error-handler
 * @description 提供统一的错误处理和警告收集机制，用于转换过程中的错误管理。
 *
 * @example
 * ```typescript
 * const errorHandler = new ConversionErrorHandler(requestId, logger);
 *
 * // 处理元素转换错误
 * const result = errorHandler.handleElementError(element, error, {
 *   operation: 'convert',
 *   slideIndex: 0,
 * });
 *
 * // 添加警告
 * errorHandler.addWarning('WARN_ELEMENT_SKIPPED', 'Unsupported element type', {
 *   elementType: 'smartArt',
 * });
 *
 * // 获取所有警告
 * const warnings = errorHandler.getWarnings();
 * ```
 */

import type { Logger } from 'pino'
import type { WarningInfo } from '../types/index.js'
import { getLogger } from './logger.js'

/**
 * 元素错误处理选项
 */
export interface ElementErrorOptions {
  /** 操作类型：解析(parsing)或转换(conversion) */
  operation?: 'parsing' | 'conversion'
  /** 当前幻灯片索引 */
  slideIndex?: number
  /** 是否抛出错误（默认 false，仅记录警告） */
  throw?: boolean
  /** 默认返回值 */
  defaultValue?: unknown
}

/**
 * 警告详情
 */
export interface WarningDetail {
  /** 警告代码 */
  code: string
  /** 警告消息 */
  message: string
  /** 相关元素 ID */
  elementId?: string
  /** 当前幻灯片索引 */
  slideIndex?: number
  /** 额外详情 */
  details?: Record<string, unknown>
  /** 时间戳 */
  timestamp: number
}

/**
 * 转换错误处理器
 *
 * 提供统一的错误处理和警告收集机制，用于 PPTX 解析和转换过程。
 *
 * @class ConversionErrorHandler
 * @example
 * ```typescript
 * const errorHandler = new ConversionErrorHandler('req-123', logger);
 *
 * try {
 *   const element = parseShape(data, context);
 * } catch (error) {
 *   const result = errorHandler.handleElementError(
 *     { id: 'shape-1', type: 'shape' },
 *     error,
 *     { operation: 'parsing', slideIndex: 0 }
 *   );
 * }
 * ```
 */
export class ConversionErrorHandler {
  private readonly requestId: string
  private readonly logger: Logger
  private readonly warnings: WarningDetail[] = []

  /**
   * 创建错误处理器实例
   *
   * @param requestId - 请求唯一标识符，用于日志追踪
   * @param logger - 可选的日志记录器，如果不提供则使用默认 logger
   */
  constructor(requestId: string, logger?: Logger) {
    this.requestId = requestId
    this.logger = logger || getLogger()
  }

  /**
   * 处理元素级别的错误
   *
   * 当元素解析或转换失败时调用此方法。错误会被记录到日志，
   * 并添加到警告列表，但不会中断整个转换过程（除非指定 throw: true）。
   *
   * @param element - 出错的元素，包含 id 和其他信息
   * @param error - 捕获的错误对象
   * @param options - 错误处理选项
   * @returns 返回 defaultValue 或 null，允许调用者继续处理其他元素
   *
   * @example
   * ```typescript
   * const errorHandler = new ConversionErrorHandler(requestId);
   *
   * for (const element of elements) {
   *   try {
   *     const result = convertElement(element, context);
   *     if (result) output.push(result);
   *   } catch (error) {
   *     // 错误被记录但不会中断循环
   *     errorHandler.handleElementError(element, error, {
   *       operation: 'conversion',
   *       slideIndex: i,
   *     });
   *   }
   * }
   * ```
   */
  handleElementError<T = unknown>(
    element: { id?: string; type?: string; [key: string]: unknown },
    error: unknown,
    options: ElementErrorOptions = {}
  ): T | null {
    const {
      operation = 'conversion',
      slideIndex,
      throw: shouldThrow = false,
      defaultValue = null,
    } = options

    const elementId = element?.id || 'unknown'
    const elementType = element?.type || 'unknown'

    // 构建错误信息
    const errorMessage = error instanceof Error ? error.message : String(error)
    const errorStack = error instanceof Error ? error.stack : undefined

    // 记录错误日志
    this.logger.warn(
      {
        requestId: this.requestId,
        elementId,
        elementType,
        operation,
        slideIndex,
        error: errorMessage,
        stack: errorStack,
      },
      `Element ${operation} failed: ${elementId}`
    )

    // 添加到警告列表
    this.addWarning('WARN_ELEMENT_FAILED', `Element ${operation} failed: ${errorMessage}`, {
      elementId,
      elementType,
      slideIndex,
      operation,
    })

    // 如果需要抛出错误
    if (shouldThrow) {
      if (error instanceof Error) {
        throw error
      }
      throw new Error(errorMessage)
    }

    return defaultValue as T | null
  }

  /**
   * 添加警告信息
   *
   * 用于记录非致命的问题，如跳过的元素、降级处理等。
   *
   * @param code - 警告代码，遵循 WARN_XXX_XXX 格式
   * @param message - 人类可读的警告消息
   * @param details - 可选的额外详情
   *
   * @example
   * ```typescript
   * errorHandler.addWarning(
   *   'WARN_SMARTART_SKIPPED',
   *   'SmartArt elements are not supported and were skipped',
   *   { count: 3, slideIndexes: [0, 2, 5] }
   * );
   * ```
   */
  addWarning(
    code: string,
    message: string,
    details?: {
      elementId?: string
      elementType?: string
      slideIndex?: number
      [key: string]: unknown
    }
  ): void {
    const warning: WarningDetail = {
      code,
      message,
      elementId: details?.elementId,
      slideIndex: details?.slideIndex,
      details,
      timestamp: Date.now(),
    }

    this.warnings.push(warning)

    // 记录调试日志
    this.logger.debug(
      {
        requestId: this.requestId,
        warning: warning,
      },
      `Warning added: ${code}`
    )
  }

  /**
   * 批量添加警告
   *
   * 用于合并来自其他处理器的警告。
   *
   * @param warnings - 要添加的警告数组
   */
  addWarnings(warnings: WarningInfo[]): void {
    for (const warning of warnings) {
      this.addWarning(warning.code, warning.message, {
        count: warning.count,
      })
    }
  }

  /**
   * 获取所有收集的警告
   *
   * @returns 警告信息数组，包含代码、消息和计数
   *
   * @example
   * ```typescript
   * const warnings = errorHandler.getWarnings();
   * // [
   * //   { code: 'WARN_ELEMENT_FAILED', message: '...', count: 2 },
   * //   { code: 'WARN_SMARTART_SKIPPED', message: '...', count: 3 }
   * // ]
   * ```
   */
  getWarnings(): WarningInfo[] {
    // 按警告代码分组并计数
    const warningMap = new Map<string, { message: string; count: number }>()

    for (const warning of this.warnings) {
      const existing = warningMap.get(warning.code)
      if (existing) {
        existing.count++
      } else {
        warningMap.set(warning.code, {
          message: warning.message,
          count: 1,
        })
      }
    }

    // 转换为数组格式
    return Array.from(warningMap.entries()).map(([code, data]) => ({
      code,
      message: data.message,
      count: data.count,
    }))
  }

  /**
   * 获取所有原始警告详情
   *
   * 与 getWarnings() 不同，此方法返回所有警告的完整详情，不进行分组。
   *
   * @returns 原始警告详情数组
   */
  getWarningDetails(): WarningDetail[] {
    return [...this.warnings]
  }

  /**
   * 检查是否有任何警告
   *
   * @returns 如果有警告返回 true
   */
  hasWarnings(): boolean {
    return this.warnings.length > 0
  }

  /**
   * 获取警告数量
   *
   * @returns 警告总数
   */
  getWarningCount(): number {
    return this.warnings.length
  }

  /**
   * 清除所有警告
   *
   * 通常用于测试或重置处理器状态。
   */
  clearWarnings(): void {
    this.warnings.length = 0
  }

  /**
   * 创建子处理器
   *
   * 用于处理特定任务（如单个幻灯片）时创建独立的错误处理器，
   * 同时共享警告收集。
   *
   * @param context - 子上下文名称，用于日志区分
   * @returns 新的错误处理器实例
   */
  createChildProcessor(context: string): ConversionErrorHandler {
    const childLogger = this.logger.child({ context })
    const child = new ConversionErrorHandler(this.requestId, childLogger)

    // 返回一个代理，将警告转发到父处理器
    const originalAddWarning = child.addWarning.bind(child)
    child.addWarning = (code, message, details) => {
      originalAddWarning(code, message, details)
      this.addWarning(code, message, details)
    }

    return child
  }
}

/**
 * 创建错误处理器的工厂函数
 *
 * 便捷函数，用于快速创建错误处理器实例。
 *
 * @param requestId - 请求唯一标识符
 * @param logger - 可选的日志记录器
 * @returns 新的 ConversionErrorHandler 实例
 */
export function createErrorHandler(
  requestId: string,
  logger?: Logger
): ConversionErrorHandler {
  return new ConversionErrorHandler(requestId, logger)
}

export default ConversionErrorHandler
