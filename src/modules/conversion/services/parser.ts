/**
 * PPTX 解析器
 *
 * @module modules/conversion/services/parser
 * @description 解析 PPTX 文件并提取演示文稿数据。
 * 遵循 ECMA-376 Office Open XML 标准。
 *
 * 此文件作为向后兼容的入口点，从模块化的 parser/ 目录重新导出。
 *
 * @example
 * ```typescript
 * import { parsePPTX } from './services/parser.js';
 *
 * const buffer = await fs.readFile('presentation.pptx');
 * const presentation = await parsePPTX(buffer);
 *
 * console.log(`幻灯片数量: ${presentation.slides.length}`);
 * ```
 */

// 从模块化的 parser 目录重新导出所有内容
export { parsePPTX, readXmlFile, getMimeType } from './parser/index.js'
export { getContentTypes, type ContentTypesResult } from './parser/content-types.js'
export { getSlideInfo, type SlideInfoResult } from './parser/slide-info.js'
export { getTheme, type ThemeResult } from './parser/theme.js'
export { parseRelationships, parseSlideRels, type RelationshipsResult } from './parser/relationships.js'
export {
  parseShape,
  parseTransform,
  parseTextBodyToParagraphs,
} from './parser/elements/index.js'

// 默认导出
export { default } from './parser/index.js'

