# PPTist-backend API 文档

## 基础信息

- **Base URL**: `http://localhost:3000/api/v1`
- **Content-Type**: `multipart/form-data` (文件上传), `application/json` (响应)

## 端点

### 转换 API

#### POST /convert

上传并转换 PPTX 文件。

**请求**

- Method: `POST`
- Path: `/api/v1/convert`
- Content-Type: `multipart/form-data`

**查询参数**

| 参数 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| format | string | pptist | 输出格式：`both`, `json`, `pptist` |

**请求体**

| 字段 | 类型 | 必需 | 描述 |
|------|------|------|------|
| file | File | 是 | PPTX 文件 (最大 50MB) |

**cURL 示例**

```bash
# 默认格式 (pptist)
curl -X POST http://localhost:3000/api/v1/convert \
  -F "file=@presentation.pptx" \
  --output presentation.pptist

# JSON 格式
curl -X POST "http://localhost:3000/api/v1/convert?format=json" \
  -F "file=@presentation.pptx" \
  --output presentation.json

# 双输出 (JSON + PPTist)
curl -X POST "http://localhost:3000/api/v1/convert?format=both" \
  -F "file=@presentation.pptx"
```

**响应**

#### format=pptist (默认)

```
HTTP 200 OK
Content-Type: application/octet-stream
Content-Disposition: attachment; filename="pptist-Conversion.pptist"

<加密的二进制数据>
```

#### format=json

```json
HTTP 200 OK
Content-Type: application/json
Content-Disposition: attachment; filename="pptist-Conversion.json"

{
  "slides": [
    {
      "id": "slide-1",
      "elements": [...],
      "background": {...}
    }
  ],
  "media": {
    "0_rId1": {
      "type": "image",
      "data": "base64...",
      "mimeType": "image/png"
    }
  },
  "metadata": {
    "slideCount": 10,
    "slideSize": { "width": 1920, "height": 1080 },
    "conversionTime": 1234
  },
  "warnings": []
}
```

#### format=both

```json
HTTP 200 OK
Content-Type: application/json

{
  "json": {
    "slides": [...],
    "media": {...},
    "metadata": {...},
    "warnings": []
  },
  "pptist": "U2FsdGVkX1..."
}
```

### 健康检查 API

#### GET /health

服务健康状态检查。

**响应**

```json
{
  "status": "ok",
  "memory": {
    "heapUsed": 12345678,
    "heapTotal": 23456789,
    "rss": 34567890
  },
  "uptime": 3600
}
```

#### GET /ready

Kubernetes 就绪探针。

**响应**

```
HTTP 200 OK
ready
```

#### GET /live

Kubernetes 存活探针。

**响应**

```
HTTP 200 OK
alive
```

## 错误响应

所有错误响应遵循统一格式：

```json
{
  "success": false,
  "error": {
    "code": "ERR_INVALID_FORMAT",
    "message": "File is not a valid PPTX",
    "suggestion": "Please upload a PowerPoint (.pptx) file"
  }
}
```

### 错误代码

| 代码 | HTTP 状态 | 描述 | 建议 |
|------|----------|------|------|
| ERR_INVALID_FORMAT | 400 | 文件不是有效的 PPTX | 上传 PowerPoint (.pptx) 文件 |
| ERR_FILE_TOO_LARGE | 413 | 文件超过大小限制 | 压缩演示文稿或减少媒体内容 |
| ERR_PROTECTED_FILE | 400 | 密码保护的文件 | 移除密码保护后重试 |
| ERR_CORRUPTED_FILE | 400 | 文件损坏或不可读 | 验证文件可以在 PowerPoint 中打开 |
| ERR_EMPTY_FILE | 400 | 文件为空或没有幻灯片 | 上传至少包含一张幻灯片的演示文稿 |
| ERR_CONVERSION_FAILED | 500 | 转换过程中发生意外错误 | 重试或联系支持 |

## 警告

转换过程中可能产生非致命警告：

| 代码 | 描述 |
|------|------|
| WARN_SMARTART_SKIPPED | SmartArt 元素被跳过 |
| WARN_MACRO_SKIPPED | 宏/VBA 元素被跳过 |
| WARN_ACTIVEX_SKIPPED | ActiveX 控件被跳过 |
| WARN_FONT_FALLBACK | 某些字体被替换为系统默认 |
| WARN_ELEMENT_FAILED | 某些元素转换失败 |

## 速率限制

- **窗口**: 60 秒
- **最大请求**: 10 (可配置)
- **白名单**: 127.0.0.1 (本地开发)

超出限制时返回 429 状态码。

## 文件限制

- **最大文件大小**: 50MB (可配置)
- **支持格式**: .pptx (Office Open XML)
- **密码保护**: 不支持

## JavaScript SDK 示例

```javascript
async function convertPPTX(file) {
  const formData = new FormData();
  formData.append('file', file);

  const response = await fetch('http://localhost:3000/api/v1/convert?format=both', {
    method: 'POST',
    body: formData,
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error.message);
  }

  return response.json();
}

// 使用示例
const fileInput = document.getElementById('file-input');
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  try {
    const result = await convertPPTX(file);
    console.log('转换成功:', result);
  } catch (error) {
    console.error('转换失败:', error.message);
  }
});
```

## TypeScript 类型定义

```typescript
interface ConversionResponse {
  json?: PPTistPresentation;
  pptist?: string;
}

interface PPTistPresentation {
  slides: Slide[];
  media: Record<string, MediaInfo>;
  metadata: ConversionMetadata;
  warnings: WarningInfo[];
}

interface Slide {
  id: string;
  elements: PPTElement[];
  background?: SlideBackground;
  remark?: string;
}

interface MediaInfo {
  type: 'image' | 'video' | 'audio';
  data: string; // base64
  mimeType: string;
}

interface WarningInfo {
  code: string;
  message: string;
  count?: number;
}

interface ErrorResponse {
  success: false;
  error: {
    code: string;
    message: string;
    suggestion?: string;
  };
}
```
