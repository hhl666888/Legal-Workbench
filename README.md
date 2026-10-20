# 合同审阅智能体 · 法务工作台

面向公司法务人员的本地合同审阅工作台。用户上传合同后，系统使用 DeepSeek 完成合同风险审查、法律法规引用审查，并生成可下载的结构化风险报告；最终可在原合同 DOCX 上仅替换已采纳条款，输出一份保留原排版的新合同。

## 当前能力

### 合同审阅

- 支持 `DOCX`、可复制文字的 `PDF`、`TXT`（单个文件 ≤ 20 MB）
- 常规合同风险审查：付款、违约责任、知识产权、保密、自动续约、争议解决、单方解除、主体与签署、交付验收、数据与合规等 10 类固定风险类别
- 法律法规引用审查：识别法律名称、条文、废止或被替代的法律
- 甲方信息提取：从合同原文中提取甲方公司名称与合同标题，不做推测
- 长合同分段处理：按约 20,000 字切分，逐段审阅后汇总去重
- DeepSeek 请求失败自动重试：覆盖 `429`、`500`、`502`、`503`
- 异步任务状态：`queued`、`processing`、`completed`、`failed`，前端实时轮询进度
- 报告支持手工编辑、AI 重新生成、一键采纳全部 AI 意见

### 报告与条款采纳

- 每条建议可单独「采纳 / 不采纳」；状态存储在前端 `window.currentReport.issues[i]._status`
- 生成新 DOCX 时仅替换已采纳条款，使用「原文证据 + 条款位置」在原合同 DOCX 中精准定位段落，保留原合同标题、段落、表格、页眉页脚等结构
- 无法准确定位的采纳条款不会强制替换，会计入「待人工确认」并在响应中列出 `pendingItems`
- 替换算法采用 `split().join()` 方式，避免 `$` 等特殊字符导致的替换错误

### 新合同持久化

- 新生成的 DOCX 由后端持久化：元数据存储在 `data/legal-review.db`（SQLite），文件存储在 `data/generated/`
- 通过登录鉴权接口 `/api/generated/:id/download` 下载，不再依赖浏览器 IndexedDB
- 审阅记录页可根据元数据标志条件性显示「下载新合同」按钮

### 鉴权与角色权限

- 员工登录与角色权限：管理员（admin）、法务专员（legal）、只读员工（viewer）
- 会话 token：登录后颁发随机 Bearer Token，有效期 8 小时
- 权限规则：
  - 只读员工（viewer）无权发起审阅、生成新合同、编辑报告
  - 法务专员（legal）可发起审阅、生成新合同、编辑报告
  - 管理员（admin）可查看所有人的审阅记录与新合同
- 登录页提供「法务专员账号密码」快捷登录按钮，一键填充并提交

## 快速启动

### 1. 配置 API Key

复制 `.env.example` 为 `.env`，并填入真实 DeepSeek API Key：

```env
PORT=8789
DEEPSEEK_API_KEY=你的 DeepSeek API Key
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
DEEPSEEK_MODEL=deepseek-chat
```

> 服务优先读取 `.env`；若不存在，则读取 `.env.example`。
> 默认占位文字（如 `替换为你的...`）不会被识别为有效配置；可通过 `GET /api/health` 查看 `deepseekConfigured` 标志。

### 2. 启动服务

双击 [`启动合同审阅智能体.cmd`](./启动合同审阅智能体.cmd)，或在终端运行：

```text
node server.js
```

### 3. 访问工作台

打开浏览器：<http://localhost:8789>

### 默认账号

| 账号 | 密码 | 角色 | 权限说明 |
| --- | --- | --- | --- |
| `admin` | `admin123` | 系统管理员 | 查看所有人审阅记录、生成新合同、编辑报告 |
| `legal` | `legal123` | 法务专员 | 发起审阅、生成新合同、编辑报告 |
| `viewer` | `viewer123` | 只读员工 | 仅查看已有记录 |

> 仅演示用，正式部署请修改 `AUTH_USERS` 环境变量或后端 `users` 表。

### 系统要求

- Node.js ≥ 22（使用内置 `node:sqlite` 模块）
- Windows 系统（DOCX 解析使用 PowerShell `System.IO.Compression`；PDF 解析需系统已安装 `pdftotext`）

## 使用流程

1. 上传合同文件（DOCX / PDF / TXT）
2. 输入本次审阅要求，可选用预设模板（全面审查 / 采购合同 / 技术服务 / 高风险条款）
3. 点击「开始智能审阅」
4. 前端创建审阅任务并轮询任务状态
5. 后端提取合同文本，按约 20,000 字分段调用 DeepSeek
6. 后端进行法律法规引用审查、标题提取和结果汇总
7. 前端展示最终报告，可下载为 Markdown、编辑或交由 AI 重新生成
8. 在报告中点击每条建议的「采纳 / 不采纳」
9. 点击「生成新合同」：后端在原 DOCX 上仅替换已采纳条款，输出保留排版的新版本
10. 在「审阅完成记录」或「审阅记录」页下载新合同

## 目录说明

| 文件 | 作用 |
| --- | --- |
| `index.html` | 法务审阅工作台页面（登录卡片、上传、审阅流程、报告弹窗） |
| `styles.css` | 页面样式与响应式布局 |
| `app.js` | 上传、任务轮询、报告预览、采纳状态、新合同下载 |
| `server.js` | 本地 HTTP 服务、文件解析、DeepSeek 调用、DOCX 生成、SQLite 持久化、登录鉴权 |
| `.env.example` | DeepSeek 配置文件模板（**不应提交真实 Key**） |
| `启动合同审阅智能体.cmd` | Windows 一键启动脚本 |
| `合同审阅智能体.workflow.json` | 原始工作流节点、变量和提示词定义 |
| `data/legal-review.db` | SQLite 数据库，存储生成合同元数据（运行时生成） |
| `data/generated/` | 生成的 DOCX 合同文件目录（运行时生成） |

## 接口

### 健康检查

```text
GET /api/health
```

返回服务状态、端口和 DeepSeek 是否已配置。

### 登录

```text
POST /api/login
Content-Type: application/json

{ "username": "legal", "password": "legal123" }
```

返回 `token`、`user` 与 `expiresIn`（8 小时）。

### 当前用户

```text
GET /api/me
Authorization: Bearer <token>
```

### 创建审阅任务

```text
POST /api/review
Authorization: Bearer <token>
Content-Type: application/json
```

请求字段：

- `fileName`：文件名
- `fileBase64`：文件 Base64 内容
- `request`：用户审阅要求

返回 `taskId`，状态为 `queued`。

### 查询任务

```text
GET /api/review/:taskId
```

状态包括：`queued`、`processing`、`completed`、`failed`。完成后返回 `report`、`audit`、`meta` 等字段。

### 获取合同原文

```text
GET /api/review/:taskId/text
```

返回 `{ taskId, text }`，用于在弹窗中查看合同原文。

### AI 重新生成报告

```text
POST /api/refine
Authorization: Bearer <token>
Content-Type: application/json

{ "content": "<当前报告内容>", "instruction": "<可选修改意见>" }
```

### 生成新合同 DOCX

```text
POST /api/generate
Authorization: Bearer <token>
Content-Type: application/json

{
  "fileName": "原合同.docx",
  "fileBase64": "<原 DOCX Base64>",
  "reviewId": "<可选>",
  "issues": [{ "_status": "adopted", "evidence": "...", "draft_clause": "..." }]
}
```

返回 `generatedId`、`fileName`、`downloadUrl`、`applied`、`pending`、`pendingItems`。

### 列出生成记录

```text
GET /api/generated
Authorization: Bearer <token>
```

管理员可查看全部记录，其他角色仅查看自己的记录。

### 下载新合同

```text
GET /api/generated/:id/download
Authorization: Bearer <token>
```

直接下载 DOCX 文件。

## 文件解析说明

- `DOCX` 使用 Windows PowerShell 的 `System.IO.Compression` 读取 `word/document.xml`；中文临时路径用 UTF-16LE Base64 编码命令避免 ByteString 错误
- `PDF` 依赖系统中的 `pdftotext`，扫描版 PDF 需后续接入 OCR
- 不支持旧版二进制 `DOC`，请先另存为 `DOCX` 或 PDF
- 加密文件、损坏文件和没有文本层的 PDF 会返回明确错误，不会让模型猜测内容

## 安全与限制

- **不要把 API Key 提交到公开仓库或发送给他人**。若 Key 已暴露，应立即撤销并重新生成
- 默认账号（admin / legal / viewer）仅用于演示，正式部署请通过 `AUTH_USERS` 环境变量覆盖
- 当前服务面向本机使用，未做 HTTPS、CSRF、限流等公网防护
- 合同原文与任务结果在服务运行期间保留在内存中，服务重启后任务结果会丢失；已生成的新合同 DOCX 持久化在 `data/generated/` 不会丢失
- 企业风险核查暂时关闭，报告会明确标注这一状态
- 最终报告是 AI 辅助结果，重大合同事项仍需法务人工复核

## 后续可扩展项

- OCR 扫描件识别
- 元典法律检索正式接入（`.env.example` 已预留 `YUANDIAN_API_KEY` / `YUANDIAN_SEARCH_URL`）
- 历史报告持久化与审计日志
- 条款定位高亮
- 风险统计图
- Prompt 版本管理
- 人工复核状态
- 企业核查开关
- 合同模板比对
- PostgreSQL 替换 SQLite 用于多用户共享部署
