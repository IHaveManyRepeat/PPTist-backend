# PPTist-backend 架构设计文档

## 概述

PPTist-backend 是一个将 PPTX 文件转换为 PPTist 格式的后端服务。本服务遵循 ECMA-376 Office Open XML 标准，支持多种元素类型的解析和转换。

## 技术栈

- **运行时**: Node.js 20+ LTS
- **语言**: TypeScript 5+
- **Web 框架**: Fastify 5+
- **文件处理**: JSZip, fast-xml-parser
- **加密**: CryptoJS (AES)
- **验证**: Zod
- **日志**: Pino

## 架构设计

### 分层架构

```
┌─────────────────────────────────────────────────────────────┐
│                      Routes Layer                           │
│  (API 端点定义，请求验证，响应格式化)                          │
├─────────────────────────────────────────────────────────────┤
│                     Services Layer                          │
│  (业务逻辑协调：parser, converter, encryptor, serializer)     │
├─────────────────────────────────────────────────────────────┤
│                    Converters Layer                         │
│  (元素转换器：shape, image, text, table, chart 等)           │
├─────────────────────────────────────────────────────────────┤
│                     Parsers Layer                           │
│  (专用解析器：chart-parser, table-parser)                    │
├─────────────────────────────────────────────────────────────┤
│                     Resolvers Layer                         │
│  (属性解析器：color, fill, border, shadow)                   │
├─────────────────────────────────────────────────────────────┤
│                    Generators Layer                         │
│  (生成器：svg-path-generator, html-text-generator)           │
└─────────────────────────────────────────────────────────────┘
```

### 核心模块

#### 1. 转换模块 (`src/modules/conversion/`)

主要的转换逻辑模块，包含所有转换相关的代码。

**目录结构:**

```
conversion/
├── context/           # 解析上下文定义
│   └── parsing-context.ts
├── converters/        # 元素转换器
│   ├── index.ts       # 转换器注册和分发
│   ├── audio.ts       # 音频转换器
│   ├── chart.ts       # 图表转换器
│   ├── image.ts       # 图片转换器
│   ├── latex.ts       # LaTeX 转换器
│   ├── line.ts        # 线条转换器
│   ├── shape.ts       # 形状转换器
│   ├── table.ts       # 表格转换器
│   ├── text.ts        # 文本转换器
│   └── video.ts       # 视频转换器
├── detectors/         # 检测器
│   ├── password.ts    # 密码保护检测
│   └── unsupported.ts # 不支持元素检测
├── generators/        # 生成器
│   ├── html-text-generator.ts
│   ├── index.ts
│   └── svg-path-generator.ts
├── parsers/           # 专用解析器
│   ├── chart-parser.ts
│   ├── index.ts
│   └── table-parser.ts
├── resolvers/         # 属性解析器
│   ├── border-resolver.ts
│   ├── color-resolver.ts
│   ├── fill-resolver.ts
│   ├── index.ts
│   └── shadow-resolver.ts
├── routes/            # API 路由
│   ├── convert.ts     # 转换 API
│   ├── health.ts      # 健康检查
│   └── ui.ts          # UI 路由
├── services/          # 核心服务
│   ├── converter.ts   # 转换协调器
│   ├── encryptor.ts   # AES 加密
│   ├── parser.ts      # PPTX 解析
│   ├── response.ts    # 响应格式化
│   └── serializer.ts  # PPTist 序列化
├── types/             # 类型定义
│   ├── pptist.ts      # PPTist 格式类型
│   ├── pptx.ts        # PPTX 格式类型
│   └── response.ts    # API 响应类型
└── utils/             # 工具函数
    └── geometry.ts    # 几何计算
```

#### 2. 转换器注册机制

采用策略模式实现可扩展的转换器注册：

```typescript
// 注册转换器
registerConverter(converter, detector, priority)

// 获取匹配的转换器
const converter = getConverter(element)

// 执行转换
const result = convertElement(element, context)
```

**优先级规则:**
- 数字越大，优先级越高
- 高优先级转换器优先匹配
- 用于处理重叠的元素类型（如 shape 和 text）

#### 3. 上下文设计

**ParsingContext** - PPTX 解析阶段状态:

```typescript
interface ParsingContext {
  zip: JSZip
  slideContent: XmlObject
  slideLayoutContent: XmlObject
  slideMasterContent: XmlObject
  themeContent: XmlObject
  themeColors: string[]
  // ...更多字段
}
```

**ConversionContext** - 转换阶段状态:

```typescript
interface ConversionContext {
  requestId: string
  startTime: number
  warnings: WarningInfo[]
  mediaMap: Map<string, MediaInfo>
  slideSize: { width: number; height: number }
  currentSlideIndex: number
}
```

### 数据流

```
┌──────────┐    ┌─────────────┐    ┌─────────────┐    ┌───────────┐
│  PPTX    │───▶│   Parser    │───▶│  Converter  │───▶│ Serializer│
│  File    │    │  (解析 XML)  │    │ (转换元素)   │    │ (序列化)  │
└──────────┘    └─────────────┘    └─────────────┘    └───────────┘
                    │                    │                  │
                    ▼                    ▼                  ▼
               PPTXPresentation    PPTistSlide[]      PPTistData
               (内部数据结构)       (转换后数据)        (最终输出)
                                                              │
                                                              ▼
                                                      ┌───────────┐
                                                      │ Encryptor │
                                                      │ (可选加密) │
                                                      └───────────┘
```

### 错误处理

使用统一的错误处理器 `ConversionErrorHandler`：

```typescript
const errorHandler = new ConversionErrorHandler(requestId, logger);

// 处理元素错误（非致命）
errorHandler.handleElementError(element, error, {
  operation: 'conversion',
  slideIndex: 0,
});

// 添加警告
errorHandler.addWarning('WARN_SMARTART_SKIPPED', 'SmartArt was skipped');

// 获取所有警告
const warnings = errorHandler.getWarnings();
```

**错误类型:**

| 错误代码 | HTTP 状态 | 描述 |
|---------|----------|------|
| ERR_INVALID_FORMAT | 400 | 无效的 PPTX 文件 |
| ERR_FILE_TOO_LARGE | 413 | 文件超过大小限制 |
| ERR_PROTECTED_FILE | 400 | 密码保护的文件 |
| ERR_CORRUPTED_FILE | 400 | 损坏的文件 |
| ERR_EMPTY_FILE | 400 | 空文件 |
| ERR_CONVERSION_FAILED | 500 | 转换失败 |

## 支持的元素类型

| 元素类型 | 支持级别 | 说明 |
|---------|---------|------|
| Text | ✅ 完全支持 | 文本和段落格式 |
| Shape | ✅ 完全支持 | 基本形状和路径 |
| Image | ✅ 完全支持 | 内嵌图片 |
| Video | ✅ 完全支持 | 内嵌视频 |
| Audio | ✅ 完全支持 | 内嵌音频 |
| Line | ✅ 完全支持 | 连接线和箭头 |
| Table | ✅ 基本支持 | 基本表格结构 |
| Chart | ⚠️ 占位数据 | 图表类型识别 |
| LaTeX | ⚠️ 需要 LaTeX 源码 | 公式渲染 |
| SmartArt | ❌ 跳过 | 不支持，显示警告 |
| Macro/VBA | ❌ 跳过 | 不支持，显示警告 |
| ActiveX | ❌ 跳过 | 不支持，显示警告 |

## 性能优化

### 当前优化

1. **流式处理**: 使用 Node.js Stream 处理大文件
2. **增量转换**: 按幻灯片逐个处理，避免内存峰值
3. **媒体按需加载**: 只在需要时读取媒体数据

### 未来优化

1. **LRU 缓存**: 缓存主题颜色、形状路径等
2. **并行处理**: 多幻灯片并行转换
3. **性能监控**: 添加转换时间指标

## 安全考虑

1. **文件大小限制**: 默认 50MB
2. **请求速率限制**: 防止 DDoS 攻击
3. **密码检测**: 拒绝加密文件
4. **内容验证**: 验证 ZIP 和 XML 结构

## 扩展开发

### 添加新的转换器

1. 在 `converters/` 目录创建新文件
2. 实现 `ElementConverter` 和 `ElementTypeDetector`
3. 使用 `registerConverter` 注册
4. 添加单元测试

```typescript
// converters/new-element.ts
export function registerNewElementConverter(): void {
  registerConverter(
    (element, context) => convertNewElement(element, context),
    isNewElement,
    5 // 优先级
  );
}
```

### 添加新的解析器

1. 在 `parsers/` 目录创建新文件
2. 导出解析函数
3. 在主解析器中调用

## 测试策略

```
tests/
├── unit/              # 单元测试
│   ├── converters/    # 转换器测试
│   └── resolvers/     # 解析器测试
├── integration/       # 集成测试
│   └── api.test.ts    # API 测试
├── utils/             # 工具测试
└── fixtures/          # 测试数据
```

**覆盖率目标**: 80%+

## 部署

### Docker

```dockerfile
FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm ci --only=production
COPY dist ./dist
EXPOSE 3000
CMD ["node", "dist/index.js"]
```

### 环境变量

| 变量 | 默认值 | 说明 |
|-----|-------|------|
| PORT | 3000 | 服务端口 |
| HOST | 0.0.0.0 | 服务主机 |
| MAX_FILE_SIZE | 52428800 | 最大文件大小 (50MB) |
| CRYPTO_KEY | pptist | AES 加密密钥 |
| RATE_LIMIT_MAX | 10 | 最大并发请求 |
| LOG_LEVEL | info | 日志级别 |
| DEFAULT_OUTPUT_FORMAT | pptist | 默认输出格式 |
