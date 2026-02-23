/**
 * PPTX 解析器
 * 重构版本：支持完整的表格、图表、样式解析
 */

import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'
import { v4 as uuidv4 } from 'uuid'
import type {
  PPTXPresentation,
  PPTXSlide,
  PPTXElement,
  PPTXTransform,
  PPTXTextRun,
  PPTXParagraph,
} from '../types/pptx.js'
import type {
  XmlObject,
  ParsingContext,
  ResourceMap,
  IndexTables,
} from '../context/parsing-context.js'
import { createDefaultParsingContext, createEmptyIndexTables } from '../context/parsing-context.js'
import { getTextByPathList, resolveSolidFill, resolveSolidFillWithAlpha } from '../resolvers/color-resolver.js'
import { resolveSlideBackgroundFill, type FillStyle } from '../resolvers/fill-resolver.js'
import { parseTable } from '../parsers/table-parser.js'
import { parseChart } from '../parsers/chart-parser.js'
import { generateShapePath } from '../generators/svg-path-generator.js'
import { Errors } from '../../../utils/errors.js'

// EMU 到点的转换比例
const RATIO_EMUs_Points = 1 / 12700

/**
 * 读取 XML 文件
 */
async function readXmlFile(zip: JSZip, path: string): Promise<XmlObject> {
  const content = await zip.file(path)?.async('string')
  if (!content) return {}

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '',
    attributesGroupName: 'attrs',
    textNodeName: '#text',
  })

  return parser.parse(content) as XmlObject
}

/**
 * 解析 [Content_Types].xml 获取文件信息
 */
async function getContentTypes(zip: JSZip): Promise<{
  slides: string[]
  slideLayouts: string[]
}> {
  const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
  const subObj = ContentTypesJson?.['Types']?.['Override'] || []

  const slidesLocArray: string[] = []
  const slideLayoutsLocArray: string[] = []

  const overrides = Array.isArray(subObj) ? subObj : [subObj]

  for (const item of overrides) {
    const contentType = item?.['attrs']?.['ContentType']
    const partName = item?.['attrs']?.['PartName']

    if (!contentType || !partName) continue

    if (contentType === 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml') {
      slidesLocArray.push(partName.substr(1))
    } else if (contentType === 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml') {
      slideLayoutsLocArray.push(partName.substr(1))
    }
  }

  // 按数字排序
  const sortByNumber = (arr: string[]) => {
    return arr.sort((a, b) => {
      const n1 = parseInt(/(\d+)\.xml/.exec(a)?.[1] || '0', 10)
      const n2 = parseInt(/(\d+)\.xml/.exec(b)?.[1] || '0', 10)
      return n1 - n2
    })
  }

  return {
    slides: sortByNumber(slidesLocArray),
    slideLayouts: sortByNumber(slideLayoutsLocArray),
  }
}

/**
 * 获取幻灯片信息
 */
async function getSlideInfo(zip: JSZip): Promise<{
  width: number
  height: number
  defaultTextStyle: XmlObject
}> {
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const sldSzAttrs = content?.['p:presentation']?.['p:sldSz']?.['attrs'] || {}
  const defaultTextStyle = content?.['p:presentation']?.['p:defaultTextStyle'] || {}

  return {
    width: parseInt(sldSzAttrs['cx'] || '9144000', 10),
    height: parseInt(sldSzAttrs['cy'] || '6858000', 10),
    defaultTextStyle,
  }
}

/**
 * 获取主题
 */
async function getTheme(zip: JSZip): Promise<{
  themeContent: XmlObject
  themeColors: string[]
  themePath?: string
}> {
  // 从 presentation.xml.rels 获取主题路径
  const preResContent = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
  const relationshipArray = preResContent?.['Relationships']?.['Relationship'] || []
  const relationships = Array.isArray(relationshipArray) ? relationshipArray : [relationshipArray]

  let themeURI: string | undefined
  for (const rel of relationships) {
    if (rel?.['attrs']?.['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
      themeURI = rel['attrs']['Target']
      break
    }
  }

  if (!themeURI) {
    return { themeContent: {}, themeColors: [] }
  }

  const themePath = `ppt/${themeURI}`
  const themeContent = await readXmlFile(zip, themePath)

  // 提取主题颜色
  const themeColors: string[] = []
  const clrScheme = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:clrScheme'])

  if (clrScheme) {
    for (let i = 1; i <= 6; i++) {
      const colorResult = getTextByPathList(clrScheme as XmlObject, [`a:accent${i}`, 'a:srgbClr', 'attrs', 'val'])
      const color = typeof colorResult === 'string' ? colorResult : undefined
      if (color) {
        themeColors.push(color.startsWith('#') ? color : '#' + color)
      }
    }
  }

  return { themeContent, themeColors, themePath }
}

/**
 * 索引节点（用于占位符查找）
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
 * 解析关系文件
 */
async function parseRelationships(
  zip: JSZip,
  relsPath: string
): Promise<{
  resources: ResourceMap
  layoutFilename?: string
  masterFilename?: string
  themeFilename?: string
  noteFilename?: string
}> {
  const resources: ResourceMap = {}
  let layoutFilename: string | undefined
  let masterFilename: string | undefined
  let themeFilename: string | undefined
  let noteFilename: string | undefined

  try {
    const relsContent = await readXmlFile(zip, relsPath)
    let relationships = relsContent?.['Relationships']?.['Relationship'] || []
    relationships = Array.isArray(relationships) ? relationships : [relationships]

    for (const rel of relationships) {
      const id = rel?.['attrs']?.['Id']
      const type = rel?.['attrs']?.['Type']
      const target = rel?.['attrs']?.['Target']

      if (!id || !type || !target) continue

      const typeName = type.replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
      const normalizedTarget = target.replace('../', 'ppt/')

      resources[id] = { type: typeName, target: normalizedTarget }

      // 识别特殊关系
      if (typeName === 'slideLayout') {
        layoutFilename = normalizedTarget
      } else if (typeName === 'slideMaster') {
        masterFilename = normalizedTarget
      } else if (typeName === 'theme') {
        themeFilename = normalizedTarget
      } else if (typeName === 'notesSlide') {
        noteFilename = normalizedTarget
      }
    }
  } catch {
    // 关系文件可能不存在
  }

  return { resources, layoutFilename, masterFilename, themeFilename, noteFilename }
}

/**
 * 解析幻灯片关系并构建上下文
 */
async function buildSlideContext(
  zip: JSZip,
  slideFilename: string,
  baseContext: ParsingContext
): Promise<ParsingContext> {
  // 解析幻灯片关系
  const slideName = slideFilename.split('/').pop()?.replace('.xml', '') || 'slide1'
  const relsPath = `ppt/slides/_rels/${slideName}.xml.rels`
  const { resources: slideResObj, layoutFilename } = await parseRelationships(zip, relsPath)

  // 解析布局
  let slideLayoutContent: XmlObject = {}
  let slideLayoutTables: IndexTables = createEmptyIndexTables()
  let layoutResObj: ResourceMap = {}

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
      }
    }
  }

  // 解析幻灯片内容
  const slideContent = await readXmlFile(zip, slideFilename)

  return {
    ...baseContext,
    slideLayoutContent,
    slideLayoutTables,
    slideResObj,
    layoutResObj,
    slideContent,
  }
}

/**
 * 解析变换信息
 */
function parseTransform(spPr: XmlObject | undefined): PPTXTransform {
  const xfrm = spPr?.['a:xfrm']
  const off = xfrm?.['a:off']?.['attrs'] || {}
  const ext = xfrm?.['a:ext']?.['attrs'] || {}
  const rot = xfrm?.['attrs']?.['rot']

  return {
    x: parseInt(off['x'] || '0', 10),
    y: parseInt(off['y'] || '0', 10),
    width: parseInt(ext['cx'] || '0', 10),
    height: parseInt(ext['cy'] || '0', 10),
    rotation: rot ? parseInt(rot, 10) / 60000 : undefined,
  }
}

/**
 * 解析文本体
 */
function parseTextBodyToParagraphs(
  txBody: XmlObject | undefined,
  context: ParsingContext
): PPTXParagraph[] {
  if (!txBody) return []

  const pArray = txBody['a:p'] || []
  const paragraphs = Array.isArray(pArray) ? pArray : [pArray]

  return paragraphs.map((p: XmlObject) => {
    const runs: PPTXTextRun[] = []

    // 解析文本运行
    const rArray = p?.['a:r'] || []
    const runArray = Array.isArray(rArray) ? rArray : [rArray]

    for (const r of runArray) {
      const run = r as XmlObject
      const rPr = run?.['a:rPr']?.['attrs'] || {}
      const text = run?.['a:t'] || ''

      // 解析颜色
      let color: string | undefined
      const solidFill = run?.['a:rPr']?.['a:solidFill']
      if (solidFill) {
        color = resolveSolidFill(solidFill, context)
      }

      runs.push({
        text: String(text),
        bold: rPr['b'] === '1',
        italic: rPr['i'] === '1',
        underline: rPr['u'] === 'sng' || rPr['u'] === '1',
        strike: rPr['strike'] === 'sngStrike',
        fontSize: rPr['sz'] ? parseInt(rPr['sz'], 10) / 100 : undefined,
        fontName: rPr['latin'] as string | undefined,
        color,
      })
    }

    // 解析段落属性
    const pPr = p?.['a:pPr']?.['attrs']
    const algn = pPr?.['algn']

    let align: 'left' | 'center' | 'right' | 'justify' | undefined
    switch (algn) {
      case 'l':
        align = 'left'
        break
      case 'r':
        align = 'right'
        break
      case 'ctr':
        align = 'center'
        break
      case 'just':
      case 'dist':
        align = 'justify'
        break
    }

    const bullet = !!(p?.['a:pPr']?.['a:buFont'] || p?.['a:pPr']?.['a:buChar'] || p?.['a:pPr']?.['a:buAutoNum'])
    const level = pPr?.['lvl'] ? parseInt(pPr['lvl'], 10) : undefined

    return { runs, align, bullet, level }
  })
}

/**
 * 解析形状
 */
function parseShape(
  shape: XmlObject,
  context: ParsingContext
): PPTXElement | null {
  const nvSpPr = shape['p:nvSpPr']
  const spPr = shape['p:spPr']
  const txBody = shape['p:txBody']

  const transform = parseTransform(spPr)
  const cNvPr = nvSpPr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  // 检查是否有实际文本内容
  let paragraphs: PPTXParagraph[] = []
  let hasActualText = false

  if (txBody) {
    paragraphs = parseTextBodyToParagraphs(txBody, context)
    // 检查段落中是否有任何非空文本
    hasActualText = paragraphs.some(p => p.runs.some(run => run.text && run.text.trim().length > 0))
  }

  // 如果有实际文本内容，返回文本元素
  if (hasActualText) {
    return {
      type: 'text',
      id: String(id),
      transform,
      name,
      paragraphs,
    }
  }

  // 形状
  const prstGeom = spPr?.['a:prstGeom']?.['attrs']?.['prst'] as string
  const shapeType = prstGeom || 'rect'

  // 解析 avLst 中的调整值（用于 roundRect 等形状的圆角控制）
  const avLst = spPr?.['a:prstGeom']?.['a:avLst']
  const gdList = avLst?.['a:gd']
  const gdArray = Array.isArray(gdList) ? gdList : gdList ? [gdList] : []

  // 解析 adj 参数
  let adj: number | undefined
  for (const gd of gdArray) {
    const gdAttrs = gd?.['attrs']
    if (gdAttrs?.['name'] === 'adj') {
      const fmla = gdAttrs['fmla'] as string
      if (fmla?.startsWith('val ')) {
        adj = parseInt(fmla.substring(4), 10)
      }
      break
    }
  }

  // 解析填充（带透明度）
  let fill: string | undefined
  let fillOpacity: number | undefined
  const solidFill = spPr?.['a:solidFill']
  if (solidFill) {
    const fillResult = resolveSolidFillWithAlpha(solidFill, context)
    fill = fillResult.color
    fillOpacity = fillResult.alpha
  }

  // 解析轮廓
  let outline: { color?: string; width?: number; style?: 'solid' | 'dashed' | 'dotted' } | undefined
  const ln = spPr?.['a:ln']
  if (ln) {
    const lnAttrs = ln?.['attrs'] || {}
    const lnSolidFill = ln['a:solidFill']
    outline = {
      color: lnSolidFill ? resolveSolidFill(lnSolidFill, context) : undefined,
      width: lnAttrs['w'] ? parseInt(lnAttrs['w'], 10) * RATIO_EMUs_Points : undefined,
      style: ln?.['a:prstDash']?.['attrs']?.['val'] === 'solid' ? 'solid' :
             ln?.['a:prstDash']?.['attrs']?.['val'] ? 'dashed' : 'solid',
    }
  }

  return {
    type: 'shape',
    id: String(id),
    transform,
    name,
    shapeType,
    adj,
    fill,
    fillOpacity,
    outline,
    path: generateShapePath(shapeType, transform.width * RATIO_EMUs_Points, transform.height * RATIO_EMUs_Points, adj),
  }
}

/**
 * 解析图片/视频/音频
 */
function parsePicture(
  pic: XmlObject,
  _context: ParsingContext
): PPTXElement | null {
  const nvPicPr = pic['p:nvPicPr']
  const spPr = pic['p:spPr']
  const blipFill = pic['p:blipFill']

  const transform = parseTransform(spPr)
  const cNvPr = nvPicPr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  // 获取 rId
  const rId = blipFill?.['a:blip']?.['attrs']?.['r:embed'] as string
  if (!rId) return null

  // 检查是否是视频/音频
  const nvPr = nvPicPr?.['p:nvPr']
  const videoFile = nvPr?.['a:videoFile']?.['attrs']?.['r:link'] as string
  const audioFile = nvPr?.['a:audioFile']?.['attrs']?.['r:link'] as string

  if (videoFile) {
    return {
      type: 'video',
      id: String(id),
      transform,
      name,
      rId: videoFile,
    }
  }

  if (audioFile) {
    return {
      type: 'audio',
      id: String(id),
      transform,
      name,
      rId: audioFile,
    }
  }

  return {
    type: 'image',
    id: String(id),
    transform,
    name,
    rId,
  }
}

/**
 * 解析图形框架（图表/表格）
 */
async function parseGraphicFrame(
  frame: XmlObject,
  context: ParsingContext
): Promise<PPTXElement | null> {
  const nvGraphicFramePr = frame['p:nvGraphicFramePr']
  const xfrm = frame['p:xfrm']
  const graphic = frame['a:graphic']

  const off = xfrm?.['a:off']?.['attrs'] || {}
  const ext = xfrm?.['a:ext']?.['attrs'] || {}

  const transform: PPTXTransform = {
    x: parseInt(off['x'] || '0', 10),
    y: parseInt(off['y'] || '0', 10),
    width: parseInt(ext['cx'] || '0', 10),
    height: parseInt(ext['cy'] || '0', 10),
  }

  const cNvPr = nvGraphicFramePr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  // 检查图形类型
  const graphicData = graphic?.['a:graphicData']
  const uri = graphicData?.['attrs']?.['uri'] as string

  // 图表
  if (uri?.includes('chart')) {
    const chartRId = graphicData?.['c:chart']?.['attrs']?.['r:id'] as string
    if (chartRId) {
      return parseChart(frame, chartRId, transform, context)
    }
    return {
      type: 'chart',
      id: String(id),
      transform,
      name,
      chartType: 'column',
      rId: '',
    }
  }

  // 表格
  if (uri?.includes('table')) {
    return parseTable(frame, transform, context)
  }

  // 不支持的图形类型
  return null
}

/**
 * 解析连接线
 */
function parseConnector(
  conn: XmlObject,
  context: ParsingContext
): PPTXElement | null {
  const nvCxnSpPr = conn['p:nvCxnSpPr']
  const spPr = conn['p:spPr']

  const transform = parseTransform(spPr)
  const cNvPr = nvCxnSpPr?.['p:cNvPr']?.['attrs']
  const id = cNvPr?.['id'] || uuidv4()
  const name = cNvPr?.['name'] as string | undefined

  const ln = spPr?.['a:ln']
  const solidFill = ln?.['a:solidFill']
  const color = solidFill ? resolveSolidFill(solidFill, context) : undefined
  const width = ln?.['attrs']?.['w'] ? parseInt(ln['attrs']['w'], 10) * RATIO_EMUs_Points : undefined

  return {
    type: 'line',
    id: String(id),
    transform,
    name,
    startX: transform.x,
    startY: transform.y,
    endX: transform.x + transform.width,
    endY: transform.y + transform.height,
    color,
    width,
  }
}

/**
 * 将 FillStyle 转换为 PPTXSlide.background 格式
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
 * 按原始顺序遍历 spTree 的所有子元素，保持 PPTX 中定义的 z-order
 */
async function parseSingleSlide(
  zip: JSZip,
  slideFilename: string,
  slideIndex: number,
  baseContext: ParsingContext
): Promise<PPTXSlide> {
  // 构建幻灯片上下文
  const context = await buildSlideContext(zip, slideFilename, baseContext)
  context.slideIndex = slideIndex

  // 获取幻灯片元素树
  const spTree = context.slideContent?.['p:sld']?.['p:cSld']?.['p:spTree']
  if (!spTree) {
    return { id: `slide-${slideIndex}`, elements: [] }
  }

  const elements: PPTXElement[] = []

  // 按原始顺序遍历 spTree 的所有子元素
  // PPTX 中元素的 z-order 由它们在 p:spTree 中的出现顺序决定
  for (const key of Object.keys(spTree)) {
    // 跳过非 p: 命名空间的元素
    if (!key.startsWith('p:')) continue

    // 跳过非元素节点（这些是容器属性，不是可视化元素）
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

  return {
    id: `slide-${slideIndex}`,
    elements,
    background,
  }
}

/**
 * 解析幻灯片关系文件获取 rId -> media target 映射
 */
async function parseSlideRels(
  zip: JSZip,
  slideNum: number
): Promise<Map<string, string>> {
  const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`
  const relsXml = await zip.file(relsPath)?.async('string')

  const rIdToTarget = new Map<string, string>()

  if (relsXml) {
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '',
      textNodeName: '#text',
    })
    const rels = parser.parse(relsXml) as XmlObject
    const relationships = rels?.['Relationships']?.['Relationship'] || []
    const relArray = Array.isArray(relationships) ? relationships : [relationships]

    for (const rel of relArray) {
      const id = rel['Id'] as string
      const target = rel['Target'] as string
      if (id && target && target.includes('media/')) {
        const normalizedTarget = target.replace(/^\.\.\//, '')
        rIdToTarget.set(id, normalizedTarget)
      }
    }
  }

  return rIdToTarget
}

/**
 * 获取 MIME 类型
 */
function getMimeType(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase()
  const mimeTypes: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    mp4: 'video/mp4',
    avi: 'video/x-msvideo',
    mov: 'video/quicktime',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
  }
  return mimeTypes[ext || ''] || 'application/octet-stream'
}

/**
 * 解析 PPTX 文件并提取演示文稿数据
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
  // 避免 PPTX 中不同幻灯片的 rId 冲突（每张幻灯片的 rId 都从 rId1 开始编号）
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
