import type { Slide } from '../types/pptist.js'
import type { ConversionContext } from '../../../types/index.js'
import { convertElement } from '../converters/index.js'
import { getLogger } from '../../../utils/logger.js'
import { createErrorHandler } from '../../../utils/error-handler.js'

/**
 * Convert a single PPTX slide to PPTist slide
 *
 * @param pptxSlide - PPTX slide data containing elements and metadata
 * @param slideIndex - Zero-based index of the slide
 * @param context - Conversion context with media map and warnings
 * @returns Converted PPTist slide
 */
export function convertSlide(
  pptxSlide: { id: string; elements: any[]; background?: any; notes?: string },
  slideIndex: number,
  context: ConversionContext
): Slide {
  const elements: any[] = []
  const logger = getLogger()
  const errorHandler = createErrorHandler(context.requestId, logger)

  for (const pptxElement of pptxSlide.elements) {
    try {
      const converted = convertElement(pptxElement, context)

      if (converted === null) {
        continue
      }

      if (Array.isArray(converted)) {
        elements.push(...converted)
      } else {
        elements.push(converted)
      }
    } catch (error) {
      // 使用统一错误处理器记录错误
      errorHandler.handleElementError(pptxElement, error, {
        operation: 'conversion',
        slideIndex,
      })
    }
  }

  // 将警告合并到上下文中
  if (errorHandler.hasWarnings()) {
    context.warnings.push(...errorHandler.getWarnings())
  }

  const slide: Slide = {
    id: `slide-${slideIndex + 1}`,
    elements,
    background: convertBackground(pptxSlide),
    remark: pptxSlide.notes,
  }

  return slide
}

/**
 * Convert slide background
 */
function convertBackground(pptxSlide: { background?: any }): Slide['background'] {
  if (!pptxSlide.background) return undefined

  const bg = pptxSlide.background

  if (bg.type === 'solid' && bg.color) {
    return {
      type: 'solid',
      color: bg.color,
    }
  }

  if (bg.type === 'image' && bg.imageRId) {
    return {
      type: 'image',
      image: {
        src: bg.imageRId,
        size: 'cover',
      },
    }
  }

  if (bg.type === 'gradient' && bg.gradient) {
    return {
      type: 'gradient',
      gradient: {
        type: bg.gradient.type,
        colors: bg.gradient.colors,
        rotate: bg.gradient.angle || 0,
      },
    }
  }

  return undefined
}

/**
 * Convert all slides from PPTX presentation
 */
export function convertSlides(
  presentation: { slides: any[] },
  context: ConversionContext
): Slide[] {
  const slides: Slide[] = []

  for (let i = 0; i < presentation.slides.length; i++) {
    // 设置当前幻灯片索引（用于媒体查找的组合键）
    context.currentSlideIndex = i
    const slide = convertSlide(presentation.slides[i], i, context)
    slides.push(slide)
  }

  return slides
}

/**
 * Create initial conversion context
 */
export function createConversionContext(
  requestId: string,
  slideSize: { width: number; height: number }
): ConversionContext {
  return {
    requestId,
    startTime: Date.now(),
    warnings: [],
    mediaMap: new Map(),
    slideSize,
    currentSlideIndex: 0,
  }
}

/**
 * Process media files from PPTX and add to context
 * Uses slideMediaMaps with slideIndex_rId composite key to avoid rId collisions
 */
export function processMedia(
  presentation: { slideMediaMaps: Map<string, { data: Buffer; contentType: string }>[] },
  context: ConversionContext
): void {
  // 为每个幻灯片的媒体使用 slideIndex_rId 作为组合键
  presentation.slideMediaMaps.forEach((slideMedia, slideIndex) => {
    for (const [rId, mediaInfo] of slideMedia) {
      const key = `${slideIndex}_${rId}`  // 组合键避免冲突
      const base64 = mediaInfo.data.toString('base64')
      context.mediaMap.set(key, {
        type: mediaInfo.contentType.startsWith('image')
          ? 'image'
          : mediaInfo.contentType.startsWith('video')
            ? 'video'
            : 'audio',
        data: base64,
        mimeType: mediaInfo.contentType,
      })
    }
  })
}

export default {
  convertSlide,
  convertSlides,
  createConversionContext,
  processMedia,
}
