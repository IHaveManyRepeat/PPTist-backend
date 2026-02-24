# PPTist Backend

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js Version](https://img.shields.io/badge/node-%3E%3D20.0.0-brightgreen)](https://nodejs.org/)
[![TypeScript](https://img.shields.io/badge/typescript-5%2B-blue)](https://www.typescriptlang.org/)

PPTX to PPTist encrypted format conversion service with dual output support.

## Features

- 📄 **PPTX 解析** - 完整支持 Office Open XML (ECMA-376) 标准
- 🔄 **格式转换** - 将 PPTX 转换为 PPTist 兼容格式
- 🔒 **AES 加密** - CryptoJS 兼容的加密输出
- 📦 **双输出** - 支持 JSON 和加密格式同时输出
- 🚀 **高性能** - 流式处理，支持大文件
- 🎨 **元素支持** - 文本、形状、图片、视频、音频、表格、图表等
- 🛡️ **安全防护** - 文件验证、速率限制、大小限制

## Supported Elements

| Element | Support | Description |
|---------|---------|-------------|
| Text | ✅ Full | Text with formatting, paragraphs |
| Shape | ✅ Full | Basic shapes, paths, fills |
| Image | ✅ Full | Embedded images (PNG, JPG, GIF, etc.) |
| Video | ✅ Full | Embedded videos (MP4, etc.) |
| Audio | ✅ Full | Embedded audio (MP3, WAV, etc.) |
| Line | ✅ Full | Connectors with arrows |
| Table | ✅ Basic | Basic table structure |
| Chart | ⚠️ Partial | Chart type detection, placeholder data |
| LaTeX | ⚠️ Partial | Requires LaTeX source |
| SmartArt | ❌ Skipped | Not supported, warning issued |
| Macro/VBA | ❌ Skipped | Not supported, warning issued |

## Quick Start

### Prerequisites

- Node.js 20+ LTS
- npm or pnpm

### Installation

```bash
# Install dependencies
npm install

# Copy environment configuration
cp .env.example .env

# Start development server
npm run dev
```

Server will start at http://localhost:3000

### Production

```bash
# Build
npm run build

# Start production server
npm start
```

## API Endpoints

### POST /api/v1/convert

Upload a PPTX file and receive converted output in your preferred format.

**Query Parameters:**
| Parameter | Values | Default | Description |
|-----------|--------|---------|-------------|
| `format` | `both`, `json`, `pptist` | `pptist` | Output format |

**Request:**
```
POST /api/v1/convert?format=both
Content-Type: multipart/form-data

file: <PPTX file>
```

**Response by Format:**

#### format=both (Dual Output)
```json
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

#### format=json (JSON Only)
```
HTTP 200 OK
Content-Type: application/json
Content-Disposition: attachment; filename="pptist-Conversion.json"

{
  "slides": [...],
  "media": {...},
  "metadata": {...},
  "warnings": []
}
```

#### format=pptist (Encrypted Only - Default)
```
HTTP 200 OK
Content-Type: application/octet-stream
Content-Disposition: attachment; filename="pptist-Conversion.pptist"

<encrypted binary data>
```

### Health Endpoints

- `GET /api/v1/health` - Health check with memory status
- `GET /api/v1/ready` - Readiness probe
- `GET /api/v1/live` - Liveness probe

## Error Codes

| Code | HTTP Status | Description |
|------|-------------|-------------|
| `ERR_INVALID_FORMAT` | 400 | File is not a valid PPTX |
| `ERR_FILE_TOO_LARGE` | 413 | File exceeds 50MB limit |
| `ERR_PROTECTED_FILE` | 400 | Password-protected files not supported |
| `ERR_CORRUPTED_FILE` | 400 | File is corrupted or unreadable |
| `ERR_EMPTY_FILE` | 400 | File contains no slides |
| `ERR_CONVERSION_FAILED` | 500 | Internal conversion error |

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | 3000 | Server port |
| `HOST` | 0.0.0.0 | Server host |
| `MAX_FILE_SIZE` | 52428800 | Max file size (50MB) |
| `CRYPTO_KEY` | pptist | AES encryption key |
| `RATE_LIMIT_MAX` | 10 | Max concurrent requests |
| `RATE_LIMIT_WINDOW` | 60000 | Rate limit window (ms) |
| `LOG_LEVEL` | info | Log level |
| `DEFAULT_OUTPUT_FORMAT` | pptist | Default output format (both, json, pptist) |

## Scripts

```bash
npm run dev          # Development with hot reload
npm run build        # Build for production
npm start            # Start production server
npm test             # Run tests
npm run typecheck    # TypeScript type check
npm run lint         # ESLint check
npm run format       # Prettier format
```

## Project Structure

```
src/
├── app.ts                    # Fastify application entry
├── index.ts                  # Server entry point
├── config/                   # Configuration management
│   └── index.ts
├── modules/                  # Business modules
│   └── conversion/           # PPTX conversion module
│       ├── context/          # Parsing context
│       ├── converters/       # Element converters
│       ├── detectors/        # File/content detectors
│       ├── generators/       # SVG/HTML generators
│       ├── parsers/          # Specialized parsers
│       ├── resolvers/        # Property resolvers
│       ├── routes/           # API routes
│       ├── services/         # Core services
│       ├── types/            # Type definitions
│       └── utils/            # Utility functions
├── types/                    # Global type definitions
│   └── index.ts
└── utils/                    # Global utilities
    ├── crypto.ts
    ├── errors.ts
    ├── error-handler.ts
    └── logger.ts
```

## Documentation

- [Architecture Design](docs/architecture.md) - Detailed architecture documentation
- [API Reference](docs/api.md) - Complete API documentation
- [Contributing Guide](docs/contributing.md) - How to contribute

## Testing

```bash
# Run all tests
npm test

# Run with coverage
npm run test:coverage

# Watch mode
npm run test:watch
```

## Importing into PPTist

1. Download the converted `.pptist` file
2. Open PPTist application
3. Go to **File** → **Import**
4. Select the `pptist-Conversion.pptist` file
5. The presentation will be loaded

## License

MIT
