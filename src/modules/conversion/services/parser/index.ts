/**
 * PPTX 解析器
 *
 * @module modules/conversion/services/parser
 * @description 解析 PPTX 文件并提取演示文稿数据。
 * 遵循 ECMA-376 Office Open XML 标准。
 *
 * @example
 * ```typescript
 * import { parsePPTX } from './services/parser/index.js';
 *
 * const buffer = await fs.readFile('presentation.pptx');
 * const presentation = await parsePPTX(buffer);
 *
 * console.log(`幻灯片数量: ${presentation.slides.length}`);
 * ```
 */

import JSZip from 'jszip'
import type {
  PPTXPresentation,
  PPTXSlide,
  PPTXElement,
} from '../../types/pptx.js'
import type {
  XmlObject,
  ParsingContext,
  IndexTables,
} from '../../context/parsing-context.js'
import { createDefaultParsingContext, createEmptyIndexTables } from '../../context/parsing-context.js'
import { getTextByPathList, resolveSolidFill } from '../../resolvers/color-resolver.js'
import { resolveSlideBackgroundFill, type FillStyle } from '../../resolvers/fill-resolver.js'
import { Errors } from '../../../../utils/errors.js'

// 导入拆分的模块
import { readXmlFile, getMimeType } from './utils.js'
import { getContentTypes } from './content-types.js'
import { getSlideInfo } from './slide-info.js'
import { getTheme } from './theme.js'
import { parseRelationships, parseSlideRels } from './relationships.js'
import { parseShape, parseTextBodyToParagraphs } from './elements/shape.js'
import { parsePicture } from './elements/picture.js'
import { parseGraphicFrame } from './elements/graphic-frame.js'
import { parseConnector } from './elements/connector.js'

// 重新导出子模块
export * from './utils.js'
export * from './content-types.js'
export * from './slide-info.js'
export * from './theme.js'
export * from './relationships.js'
export * from './elements/index.js'

/**
 * 索引节点（用于占位符查找）
 *
 * @param content - XML 内容
 * @returns 索引表
 */
function indexNodes(content: XmlObject): IndexTables {
  const keys = Object.keys(content)
  const spTreeNode = content[keys[0]]?.['p:cSld']?.['p:spTree']

  const idTable: Record<string, XmlObject> = {}
  const idxTable: Record<string, XmlObject> = {}
  const typeTable: Record<string, XmlObject> = {}

  if (!spTreeNode) return { idTable, idxTable, typeTable }

  for (const key in spTreeNode) {
    if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') continue

    const targetNode = spTreeNode[key]
    const nodes = Array.isArray(targetNode) ? targetNode : [targetNode]

    for (const node of nodes) {
      const nvSpPrNode = node?.['p:nvSpPr'] || node?.['p:nvPicPr'] || node?.['p:nvGraphicFramePr'] || node?.['p:nvCxnSpPr']
      const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id']) as string
      const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx']) as string
      const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type']) as string

      if (id) idTable[id] = node
      if (idx) idxTable[idx] = node
      if (type) typeTable[type] = node
    }
  }

  return { idTable, idxTable, typeTable }
}

/**
 * 从 notesSlide XML 中提取备注文本
 *
 * @param notesContent - 备注幻灯片 XML 内容
 * @returns 备注文本
 */
function extractNotesText(notesContent: XmlObject): string {
  const spTree = notesContent?.['p:notes']?.['p:cSld']?.['p:spTree']
  if (!spTree) return ''

  const textParts: string[] = []

  for (const key of Object.keys(spTree)) {
    if (key !== 'p:sp') continue
    const shapes = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]

    for (const shape of shapes) {
      const phType = shape?.['p:nvSpPr']?.['p:nvPr']?.['p:ph']?.['attrs']?.['type']
      if (phType !== 'body') continue

      const txBody = shape?.['p:txBody']
      if (!txBody) continue

      const paragraphs = txBody['a:p'] || []
      const pArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs]

      for (const p of pArray) {
        const runs = p?.['a:r'] || []
        const rArray = Array.isArray(runs) ? runs : [runs]
        for (const r of rArray) {
          const text = r?.['a:t']
          if (text) textParts.push(String(text))
        }
        textParts.push('\n')
      }
    }
  }

  return textParts.join('').trim().replace(/\n{3,}/g, '\n\n')
}

/**
 * 解析幻灯片关系并构建上下文
 *
 * @param zip - JSZip 实例
 * @param slideFilename - 幻灯片文件名
 * @param baseContext - 基础上下文
 * @returns 包含上下文和备注文件名的对象
 */
async function buildSlideContext(
  zip: JSZip,
  slideFilename: string,
  baseContext: ParsingContext
): Promise<{ context: ParsingContext; noteFilename?: string }> {
  // 解析幻灯片关系
  const slideName = slideFilename.split('/').pop()?.replace('.xml', '') || 'slide1'
  const relsPath = `ppt/slides/_rels/${slideName}.xml.rels`
  const { resources: slideResObj, layoutFilename, noteFilename } = await parseRelationships(zip, relsPath)

  // 解析布局
  let slideLayoutContent: XmlObject = {}
  let slideLayoutTables: IndexTables = createEmptyIndexTables()
  let layoutResObj: Record<string, any> = {}

  if (layoutFilename) {
    slideLayoutContent = await readXmlFile(zip, layoutFilename)
    slideLayoutTables = indexNodes(slideLayoutContent)

    const layoutRelsPath = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels'
    const { resources, masterFilename } = await parseRelationships(zip, layoutRelsPath)
    layoutResObj = resources

    // 解析母版
    if (masterFilename) {
      const slideMasterContent = await readXmlFile(zip, masterFilename)
      const slideMasterTables = indexNodes(slideMasterContent)
      const slideMasterTextStylesResult = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles'])
      const slideMasterTextStyles = typeof slideMasterTextStylesResult === 'object' ? slideMasterTextStylesResult as XmlObject : undefined

      const masterRelsPath = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels'
      const { resources: masterResObj } = await parseRelationships(zip, masterRelsPath)

      // 解析幻灯片内容
      const slideContent = await readXmlFile(zip, slideFilename)

      return {
        context: {
          ...baseContext,
          slideLayoutContent,
          slideLayoutTables,
          slideMasterContent,
          slideMasterTables,
          slideMasterTextStyles,
          slideResObj,
          layoutResObj,
          masterResObj,
          slideContent,
        },
        noteFilename,
      }
    }
  }

  // 解析幻灯片内容
  const slideContent = await readXmlFile(zip, slideFilename)

  return {
    context: {
      ...baseContext,
      slideLayoutContent,
      slideLayoutTables,
      slideResObj,
      layoutResObj,
      slideContent,
    },
    noteFilename,
  }
}

/**
 * 将 FillStyle 转换为 PPTXSlide.background 格式
 *
 * @param fill - 填充样式
 * @returns 背景格式
 */
function convertFillToBackground(fill: FillStyle): PPTXSlide['background'] {
  switch (fill.type) {
    case 'solid':
      return { type: 'solid', color: fill.color }
    case 'image':
      return { type: 'image', imageRId: fill.src }
    case 'gradient':
      return {
        type: 'gradient',
        gradient: {
          type: fill.gradientType === 'linear' ? 'linear' : 'radial',
          colors: fill.colors.map(c => ({
            pos: parseInt(c.pos) / 100,
            color: c.color,
          })),
          angle: fill.angle,
        },
      }
    case 'none':
    case 'pattern':
    default:
      return undefined
  }
}

/**
 * 解析单张幻灯片
 *
 * @param zip - JSZip 实例
 * @param slideFilename - 幻灯片文件名
 * @param slideIndex - 幻灯片索引
 * @param baseContext - 基础上下文
 * @returns 解析后的幻灯片
 */
async function parseSingleSlide(
  zip: JSZip,
  slideFilename: string,
  slideIndex: number,
  baseContext: ParsingContext
): Promise<PPTXSlide> {
  // 构建幻灯片上下文
  const { context, noteFilename } = await buildSlideContext(zip, slideFilename, baseContext)
  context.slideIndex = slideIndex

  // 获取幻灯片元素树
  const spTree = context.slideContent?.['p:sld']?.['p:cSld']?.['p:spTree']
  if (!spTree) {
    return { id: `slide-${slideIndex}`, elements: [] }
  }

  const elements: PPTXElement[] = []

  // 按原始顺序遍历 spTree 的所有子元素
  for (const key of Object.keys(spTree)) {
    // 跳过非 p: 命名空间的元素
    if (!key.startsWith('p:')) continue

    // 跳过非元素节点
    if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') continue

    const items = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]

    for (const item of items) {
      let element: PPTXElement | null = null

      switch (key) {
        case 'p:sp':
          element = parseShape(item as XmlObject, context)
          break
        case 'p:pic':
          element = parsePicture(item as XmlObject, context)
          break
        case 'p:graphicFrame':
          element = await parseGraphicFrame(item as XmlObject, context)
          break
        case 'p:cxnSp':
          element = parseConnector(item as XmlObject, context)
          break
      }

      if (element) elements.push(element)
    }
  }

  // 解析幻灯片背景
  const backgroundFill = await resolveSlideBackgroundFill(context)
  const background = convertFillToBackground(backgroundFill)

  // 解析备注
  let notes: string | undefined
  if (noteFilename) {
    try {
      const notesContent = await readXmlFile(zip, noteFilename)
      const notesText = extractNotesText(notesContent)
      if (notesText) notes = notesText
    } catch {
      // 备注文件可能不存在，忽略
    }
  }

  return {
    id: `slide-${slideIndex}`,
    elements,
    background,
    notes,
  }
}

/**
 * 解析 PPTX 文件并提取演示文稿数据
 *
 * @description
 * 这是解析器的主入口函数，执行以下步骤：
 * 1. 加载 ZIP 文件并检查密码保护
 * 2. 验证文件结构（检查 presentation.xml）
 * 3. 提取幻灯片信息（尺寸、主题等）
 * 4. 解析每张幻灯片的元素（形状、图片、图表等）
 * 5. 提取媒体文件（图片、视频、音频）
 *
 * @param buffer - PPTX 文件的二进制数据
 * @returns 解析后的演示文稿对象，包含幻灯片、媒体和尺寸信息
 * @throws {ConversionError} 文件受密码保护、损坏或格式无效时抛出
 *
 * @example
 * ```typescript
 * import { readFile } from 'fs/promises';
 * import { parsePPTX } from './services/parser/index.js';
 *
 * const buffer = await readFile('presentation.pptx');
 * const presentation = await parsePPTX(buffer);
 *
 * console.log('幻灯片数量:', presentation.slides.length);
 * console.log('幻灯片尺寸:', presentation.slideSize);
 * console.log('媒体文件数量:', presentation.media.size);
 * ```
 */
export async function parsePPTX(buffer: Buffer): Promise<PPTXPresentation> {
  const zip = new JSZip()
  const contents = await zip.loadAsync(buffer)

  // 检查密码保护
  if (contents.file('EncryptedPackage')) {
    throw Errors.protectedFile()
  }

  // 检查 presentation.xml
  const presentationXml = await contents.file('ppt/presentation.xml')?.async('string')
  if (!presentationXml) {
    throw Errors.corruptedFile()
  }

  // 获取幻灯片信息
  const { width, height, defaultTextStyle } = await getSlideInfo(zip)

  // 获取主题
  const { themeContent, themeColors } = await getTheme(zip)

  // 获取文件信息
  const { slides: slideFiles } = await getContentTypes(zip)

  // 创建基础上下文
  const baseContext: ParsingContext = {
    ...createDefaultParsingContext(zip),
    themeContent,
    themeColors,
    defaultTextStyle,
  }

  // 提取媒体文件
  const media = new Map<string, { data: Buffer; contentType: string }>()
  const mediaFiles = contents.filter((relativePath) => relativePath.startsWith('ppt/media/'))
  for (const file of Object.values(mediaFiles)) {
    const filePath = file.name
    const mediaId = filePath.replace('ppt/media/', '')
    const data = await file.async('nodebuffer')
    const contentType = getMimeType(filePath)
    media.set(mediaId, { data, contentType })
  }

  // 为每个幻灯片存储独立的 rId -> media 映射
  const slideMediaMaps: Map<string, { data: Buffer; contentType: string }>[] = []

  // 解析所有幻灯片
  const slides: PPTXSlide[] = []

  for (let i = 0; i < slideFiles.length; i++) {
    const slideFilename = slideFiles[i]
    const slideNum = i + 1

    // 解析幻灯片
    const slide = await parseSingleSlide(zip, slideFilename, slideNum, baseContext)
    slides.push(slide)

    // 为每个幻灯片创建独立的 rId -> media 映射
    const slideRIdToMedia = new Map<string, { data: Buffer; contentType: string }>()
    const rIdToTarget = await parseSlideRels(zip, slideNum)
    for (const [rId, target] of rIdToTarget) {
      const mediaId = target.replace('media/', '')
      const mediaData = media.get(mediaId)
      if (mediaData) {
        slideRIdToMedia.set(rId, mediaData)
      }
    }
    slideMediaMaps.push(slideRIdToMedia)
  }

  return {
    slides,
    slideSize: { width, height },
    media,
    slideMediaMaps,
  }
}

export default { parsePPTX }
