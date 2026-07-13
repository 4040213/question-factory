# 试题批量生成工具

从 Excel 表格批量导入选择题/填空题到「莫顿教学系统」，支持飞书文档一键导入，自动完成题目创建、画布排版、通用反馈和条件反馈配置。

---

## 目录

- [这是什么？](#这是什么)
- [快速上手](#快速上手)
- [使用流程](#使用流程)
- [飞书文档导入 🪶](#飞书文档导入-)
- [Excel 表格怎么写？](#excel-表格怎么写)
- [功能说明](#功能说明)
- [项目文件结构](#项目文件结构)
- [自动化测试](#自动化测试)
- [常见问题 & 已知限制](#常见问题--已知限制)

---

## 这是什么？

莫顿教学系统（tms-mx.xueqiulearning.com）里手动创建题目费时费力。本工具支持三种导入方式：

1. **Excel 导入** — 在 Excel 模板里填好题目，上传即解析
2. **飞书导入** 🪶 — 粘贴飞书文档链接，自动拉取表格数据生成 Excel
3. **手动导入** — 网页填写题目信息后批量生成

**支持题型**：选择题（2~8 选项）、填空题（单空/双空，内容型/选项型）

**项目规模**：Python + HTML/CSS/JS + Node.js，86 个自动化测试。

---

## 快速上手

```bash
# 1. 安装依赖
pip install fastapi uvicorn httpx openpyxl python-multipart pyyaml pytest pytest-asyncio

# 2. 启动
cd "C:/Users/XQ/Desktop/question-factory"
python main.py

# 3. 打开浏览器 → http://localhost:8900
```

> 截图预览功能（可选）：`npm install`，使用系统 Chrome。

---

## 使用流程

### 步骤 1：验证 Cookie
登录莫顿 → F12 → Application → Cookies → 复制 `tokenCrm` → 粘贴验证

### 步骤 2：导入题目（三选一）

| 方式 | 入口 | 说明 |
|---|---|---|
| 📄 Excel 导入 | 点击「选择 Excel 文件」 | 支持多文件、选择题/填空题自动识别 |
| 🪶 飞书导入 | 粘贴飞书文档链接 → 点「转换」 | 自动拉取表格数据 → 生成 Excel → 加载到题目列表 |
| ✏️ 手动填写 | 预览表编辑 | 逐题修改反馈、条件反馈等 |

### 步骤 3：预览 & 批量生成
- 🔍 **布局预览**（生成前）：画布坐标预览
- 📸 **真实预览**（生成后）：Playwright 截取莫顿实际效果
- **批量生成** / **全部生成**：SSE 实时进度条

---

## 飞书文档导入 🪶

> 粘贴飞书文档链接 → 点「转换」→ 题目自动出现在网页列表里，跟手动上传 Excel 完全一样。

### 原理：数据怎么从飞书到莫顿？

```
飞书百科文档
  ↓ ① 飞书 Open API (Blocks API)
结构化块数据（table / text / image block）
  ↓ ② python 解析
question dicts（内存）
  ↓ ③ openpyxl 写入
本地 Excel 文件（D:\A工作\题目生成工作流\选择题_生成结果.xlsx / 填空题_生成结果.xlsx）
  ↓ ④ 前端自动 fetch → handleFileUpload()
前端题目列表（跟手动上传 Excel 完全一致）
  ↓ ⑤ 用户点「批量生成」
莫顿教学系统 API（逐题创建 + 画布排版 + 反馈配置）
```

### 文件储存

| 文件 | 路径 | 说明 |
|---|---|---|
| 选择题模板 | `D:\A工作\题目生成工作流\选择题(模板) - 副本.xlsx` | 31 列模板，只读 |
| 填空题模板 | `D:\A工作\题目生成工作流\填空题（模板）.xlsx` | 37 列模板，只读 |
| **选择题生成结果** | `D:\A工作\题目生成工作流\选择题_生成结果.xlsx` | 每次转换覆盖 |
| **填空题生成结果** | `D:\A工作\题目生成工作流\填空题_生成结果.xlsx` | 每次转换覆盖 |
| 飞书图片 | `D:\A工作\题目生成工作流\飞书图片\` | 按 image_token 命名 |

> 生成的 Excel 和模板同目录，方便你手动核对或修改后重新上传。

### 解析步骤详解

**① Blocks API 结构化解析**（非文本切割）

旧方案用 `raw_content` 文本 + 空行分割 → 题目含内部空行时（如"拖拽效果"前空行、"image.png"后空行）8 列对齐崩溃。v2 改用 Blocks API：

- 递归遍历文档 block 树（block_type: 1=页面 2=文本 4=分隔线 27=图片 31=表格 32=单元格）
- 找出所有 `block_type=31` 表格 → 每 8 个单元格一行 → 第一行是表头，后面是数据
- 递归读取单元格内容：文本块（block_type=2）+ 图片块（block_type=27）+ 行内图片（inline_image）
- 子 block 间自动加 `\n`，保留文档段落结构

**② 题目分类 & 属性提取**

- **选项识别**：统一扫描所有 `[A-H][.、．]` 标记（支持同行多选项、跨行选项、中文全角标点）
- **选择题**：≥2 个选项 → 题干 = 第一个选项之前的文本
- **填空题**：无选项 → 统计 `(   )` 或 `______` 数量 → 单空 / 双空
- **答案处理**：自动去空格 → 支持 BD 多选、跨行答案（如 `蓝球\n白球`）
- **跳关条件**：「错一次过」→ min=1,max=1；「错三次过」→ min=1,max=3
- **名称**：飞书「位置」列（如「关卡1-1」「轻互动2-1」）→ Excel 第 12 列「题目名称」

**③ 负反馈 → 条件反馈映射**

飞书「负反馈」列含多级文本（负反馈一/二/三），对应每次答错的不同提示。写入 Excel 时：

| 飞书内容 | Excel 列 | 对应莫顿字段 |
|---|---|---|
| 负反馈一（第一段提示） | 21「答错通用反馈」 | wrongFeedback（通用，任何错次都显示） |
| 全部负反馈（清洗标签后合并） | 29「条件反馈文本」 | conditionFeedbacks[0].content |
| 跳关条件 | 27「最少错次」/ 28「最多错次」 | minWrong / maxWrong |

**④ 图片处理**

block_type=27 或 inline_image → 提取 `image_token` → 调用 `GET /drive/v1/medias/{token}/download` 下载到本地 `飞书图片/` 目录 → token 写入 Excel 第 19 列「题目图片」

### 前端展示流程

```
点「转换」→ POST /api/feishu/convert（等待飞书 API + Excel 生成）
  ↓ 返回 {success, summary: {total, choice, fill, items}, files}
前端展示：
  ├─ 摘要：共 X 题（选择 X，填空 X） + 每题明细
  ├─ 📥 下载按钮（可点击保存 Excel 到本地）
  └─ 自动加载 Excel → 题目列表（可直接预览/生成）
```

### 飞书文档格式要求

| 位置 | 知识点 | 题目 | 答案 | 正反馈 | 负反馈 | 跳关条件 | 场景设置 |
|---|---|---|---|---|---|---|---|
| 关卡1-1 | 概率 | 下列事件中...<br>A.xxx B.xxx C.xxx D.xxx | D | 太棒了 | 负反馈一：再想想<br>负反馈二：提示<br>负反馈三：解释 | 错一次过 | |

- 8 列顺序固定，每道题一个表格
- 题目列支持 A.xxx / B.xxx 等多行选项或同行选项
- 填空题为 `(    )` 或 `______` 格式
- 负反馈支持多级（负反馈一/二/三），自动拆分到通用答错 + 条件反馈
- 支持嵌入图片（block 图片或行内图片）

### 配置飞书 API

环境变量方式（推荐）：
```bash
export FEISHU_APP_ID=cli_xxxx
export FEISHU_APP_SECRET=xxxx
```

所需权限：`doc:document:readonly` + `wiki:wiki:readonly` + `drive:drive:readonly`（图片下载）

### API 端点

| 方法 | 路径 | 说明 |
|---|---|---|
| POST | `/api/feishu/convert` | 提交飞书文档 URL，返回解析结果 + Excel 文件路径 |
| GET | `/api/feishu/download?path=...` | 下载生成的 Excel（也用于前端自动加载） |

---

## Excel 表格怎么写？

### 选择题模板（31 列）

| 列 | 列名 | 必填 | 示例 |
|---|---|---|---|
| A | 题干 | ✅ | `1+1等于几？` |
| B | 文本内容 | ❌ | `1+1 = ?` |
| C~J | 选项A~H | A/B 必填 | `1`, `2` |
| K | 答案 | ✅ | `B` |
| L | 题目名称 | ❌ | `加法练习` |
| M | 学科 | ❌ | `数学` |
| N | 分辨率 | ❌ | `1334*750` |
| O | 年级 | ❌ | `一年级` |
| P | 阶段 | ❌ | `春` |
| Q | 难度 | ❌ | `3` |
| R | 知识点 | ❌ | `加法;进位` |
| S | 题目图片 | ❌ | URL |
| T | 画布背景 | ❌ | URL |
| U~V | 答错/答对反馈 | ❌ | 文本 |
| W~Z | 答错/答对图片/音频 | ❌ | URL |
| AA~AE | 条件反馈(最少/最多错次、文本、图片、音频) | ❌ | |

### 填空题模板（37 列）

| 列 | 列名 | 必填 | 示例 |
|---|---|---|---|
| A | 题型 | ✅ | `COMPLETION` |
| B | 名称 | ❌ | |
| C | 学科 | ❌ | `数学` |
| D | 分辨率 | ❌ | `1334*750` |
| E | 年级 | ❌ | |
| F | 阶段 | ❌ | |
| G | 难度 | ❌ | |
| H | 知识点 | ❌ | |
| I | 题干文本 | ❌ | |
| J | 文本 | ✅ | `1+1=______，2+2=______。` |
| K | 题干音频 | ❌ | |
| L~X | 空1配置(名称/类型/答案类型/输入框样式/答案/别名/键盘/序号/底框/选项1~4) | 至少1空 | |
| Y~AK | 空2配置（同上） | 可选 | |

---

## 功能说明

### 智能画布布局引擎

候选评分制：生成 5 种布局候选（TEXT_ONLY / TEXT_LEFT_IMAGE_RIGHT / IMAGE_TOP_TEXT_BELOW / STACK_WITH_SCALE / IMAGE_VERTICAL），按溢出度和宽度适应性评分，选取最优。

- **绿框区域**：x=182~1152（宽970），y=130~580（高450）
- **字号自适应**：≤15字→50px, ≤40字→45px, >40字→40px
- **逐字测量**：中文 0.85×字号, 英文 0.55×字号, 标点 0.40×字号
- **选项**：弹簧公式等分 + 高度分级(100~130px)
- **降级兜底**：溢出时文本→图片上下堆叠，图片循环缩小

### 通用反馈 & 条件反馈

答错/答对支持文字+图片+音频。条件反馈基于错误次数触发（如错1次→"再试试"，错2~3次→"提示"）。

### SSE 实时进度

批量生成时每道题完成后立即推送进度事件，前端逐条渲染。

### 浏览器缓存已禁用

`Cache-Control: no-cache, no-store, must-revalidate`，更新代码刷新即可。

---

## 项目文件结构

```
question-factory/
├── main.py                      # 入口：python main.py, 8900 端口
├── requirements.txt             # Python 依赖
├── package.json                 # Node.js 依赖 (Playwright)
├── screenshot.js                # Playwright 截图脚本
├── README.md
│
├── config/
│   ├── mappings.yaml            # 学科/年级/阶段/难度 映射表
│   └── __init__.py              # 自动加载
│
├── app/
│   ├── __init__.py              # FastAPI 工厂, 5个路由
│   ├── client.py                # Morton HTTP 客户端
│   └── routers/
│       ├── auth.py              # Cookie/JWT 验证
│       ├── question.py          # ★ 核心：布局引擎、Excel解析、批量生成
│       ├── feishu.py            # 🪶 飞书导入
│       ├── metadata.py          # 映射表 API
│       └── screenshot.py        # Playwright 截图
│
├── static/
│   └── index.html               # 前端单页
│
└── tests/
    ├── test_question.py         # 40 用例
    ├── test_auth.py             # 20 用例
    ├── test_client.py           # 10 用例
    └── test_mappings.py         # 16 用例
```

---

## 自动化测试

```bash
cd "C:/Users/XQ/Desktop/question-factory"
python -m pytest tests/ -v
```

| 文件 | 用例数 | 检查范围 |
|---|---|---|
| test_question.py | 40 | Excel解析、布局引擎、Payload构建 |
| test_auth.py | 20 | Cookie/JWT/过期检测 |
| test_client.py | 10 | HTTP重试机制（含异步） |
| test_mappings.py | 16 | 映射表正确性 |
| **合计** | **86** | `python -m pytest tests/ -v` |

---

## 常见问题

| 问题 | 解决 |
|---|---|
| Excel 解析失败 | 检查表头列数/顺序 |
| Token 过期 | 重新登录莫顿，复制新 Cookie |
| 飞书转换无结果 | 检查文档权限 + 表头列名"位置/知识点/题目/答案/正反馈/负反馈/跳关条件/场景设置" |
| 截图不可用 | `npm install` 并确认 Chrome 已安装 |

---

## 技术栈

| 层 | 技术 |
|---|---|
| 后端 | FastAPI + uvicorn |
| HTTP | httpx (异步) |
| Excel | openpyxl |
| 截图 | Playwright (Node.js) + 系统 Chrome |
| 前端 | 原生 HTML/CSS/JS + SSE |
| 配置 | YAML (mappings.yaml) |
| 测试 | pytest |
