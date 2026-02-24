# 贡献指南

感谢您对 PPTist-backend 项目的关注！本文档将帮助您参与项目开发。

## 开发环境设置

### 环境要求

- Node.js 20+ LTS
- npm 或 pnpm
- Git

### 安装步骤

```bash
# 克隆仓库
git clone https://github.com/your-org/pptist-backend.git
cd pptist-backend

# 安装依赖
npm install

# 复制环境配置
cp .env.example .env

# 启动开发服务器
npm run dev
```

## 项目结构

```
pptist-backend/
├── src/
│   ├── app.ts                    # Fastify 应用入口
│   ├── index.ts                  # 服务入口
│   ├── config/                   # 配置管理
│   ├── modules/                  # 业务模块
│   │   └── conversion/           # 转换模块
│   ├── types/                    # 全局类型定义
│   └── utils/                    # 工具函数
├── tests/                        # 测试文件
│   ├── unit/                     # 单元测试
│   ├── integration/              # 集成测试
│   └── utils/                    # 测试工具
├── docs/                         # 文档
├── specs/                        # 功能规范
└── contracts/                    # API 契约
```

## 开发规范

### 代码风格

- 使用 ESLint 和 Prettier 进行代码格式化
- 遵循 TypeScript 最佳实践
- 使用 ES Modules (ESM) 格式

```bash
# 检查代码风格
npm run lint

# 自动修复
npm run lint:fix

# 格式化代码
npm run format
```

### 提交规范

使用 Conventional Commits 格式：

```
<type>(<scope>): <description>

[optional body]

[optional footer]
```

**类型:**
- `feat`: 新功能
- `fix`: Bug 修复
- `docs`: 文档更新
- `style`: 代码格式（不影响功能）
- `refactor`: 重构
- `test`: 测试相关
- `chore`: 构建/工具相关

**示例:**

```
feat(converter): 添加圆角矩形支持

- 实现 roundRect 形状路径生成
- 支持 adj 参数控制圆角大小

Closes #123
```

### 分支策略

- `main`: 主分支，稳定版本
- `develop`: 开发分支
- `feature/*`: 功能分支
- `fix/*`: 修复分支
- `release/*`: 发布分支

## 测试

### 运行测试

```bash
# 运行所有测试
npm test

# 监视模式
npm run test:watch

# 测试覆盖率
npm run test:coverage
```

### 测试规范

1. **单元测试**: 测试单个函数或类
2. **集成测试**: 测试 API 端点
3. **覆盖率目标**: 80%+

**测试文件命名:**
- 单元测试: `*.test.ts`
- 集成测试: `*.integration.test.ts`

**测试结构:**

```typescript
describe('模块名称', () => {
  describe('函数名称', () => {
    it('应该正确处理正常输入', () => {
      // arrange
      const input = 'test';

      // act
      const result = functionUnderTest(input);

      // assert
      expect(result).toBe(expected);
    });

    it('应该正确处理边界情况', () => {
      // ...
    });
  });
});
```

## 添加新功能

### 1. 创建功能规范

在 `specs/` 目录下创建功能规范文档：

```
specs/
└── 003-feature-name/
    ├── requirements.md
    ├── plan.md
    ├── spec.md
    └── tasks.md
```

### 2. 添加新转换器

1. 在 `src/modules/conversion/converters/` 创建新文件
2. 实现转换器和检测器函数
3. 注册转换器
4. 添加测试

```typescript
// src/modules/conversion/converters/new-element.ts
import { registerConverter } from './index.js';
import type { PPTXElement } from '../types/pptx.js';
import type { PPTElement } from '../types/pptist.js';
import type { ConversionContext } from '../../../types/index.js';

function isNewElement(element: PPTXElement): boolean {
  return element.type === 'new-element';
}

function convertNewElement(
  element: PPTXElement,
  context: ConversionContext
): PPTElement | null {
  // 转换逻辑
  return {
    id: element.id,
    type: 'new-element',
    // ...
  };
}

export function registerNewElementConverter(): void {
  registerConverter(
    (element, context) => convertNewElement(element, context),
    isNewElement,
    5 // 优先级
  );
}
```

3. 在 `src/modules/conversion/index.ts` 中导入和调用：

```typescript
import { registerNewElementConverter } from './converters/new-element.js';

// 初始化
registerNewElementConverter();
```

### 3. 添加测试

```typescript
// tests/unit/converters/new-element.test.ts
import { describe, it, expect, beforeEach } from 'vitest';
import { clearConverters, getConverter } from '../../../src/modules/conversion/converters/index.js';
import { registerNewElementConverter } from '../../../src/modules/conversion/converters/new-element.js';

describe('NewElement Converter', () => {
  beforeEach(() => {
    clearConverters();
  });

  it('should register converter', () => {
    registerNewElementConverter();

    const element = { type: 'new-element', id: 'test' };
    const converter = getConverter(element);

    expect(converter).toBeDefined();
  });
});
```

## 文档

### 代码注释

使用 JSDoc 格式添加注释：

```typescript
/**
 * 解析 PPTX 文件中的形状元素
 *
 * @description
 * 从 PPTX 的 spTree 中提取形状信息，包括：
 * - 位置和尺寸 (transform)
 * - 填充样式 (fill)
 * - 边框样式 (outline)
 * - 文本内容 (textBody)
 *
 * @param spTree - PPTX 形状树节点
 * @param context - 解析上下文
 * @returns 解析后的形状元素，如果解析失败返回 null
 *
 * @example
 * ```typescript
 * const shape = parseShape(spTreeNode, context);
 * if (shape?.type === 'shape') {
 *   console.log(shape.shapeType);
 * }
 * ```
 */
export function parseShape(
  spTree: XmlObject,
  context: ParsingContext
): PPTXShapeElement | null {
  // ...
}
```

### 更新文档

添加新功能时，请更新以下文档：

1. `README.md` - 添加功能说明
2. `docs/api.md` - 添加 API 变更
3. `docs/architecture.md` - 添加架构变更

## 发布流程

1. 更新版本号
2. 更新 CHANGELOG.md
3. 创建 release 分支
4. 运行完整测试
5. 合并到 main
6. 创建 Git 标签
7. 部署

## 问题反馈

如果您发现 Bug 或有功能建议：

1. 搜索现有 Issues
2. 创建新 Issue，包含：
   - 问题描述
   - 复现步骤
   - 预期结果
   - 实际结果
   - 环境信息

## 代码审查

所有提交都需要通过代码审查：

1. 确保所有测试通过
2. 代码风格符合规范
3. 添加必要的测试
4. 更新相关文档
5. 提交信息格式正确

## 许可证

本项目采用 MIT 许可证。贡献的代码将采用相同的许可证。
