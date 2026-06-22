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

**项目规模**：Python + HTML/CSS/JS + Node.js，40+ 自动化测试。

---

## 快速上手

```bash
# 1. 安装依赖
pip install fastapi uvicorn httpx openpyxl python-multipart pyyaml

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

> 从飞书文档一键导入题目，自动解析表格并填入 Excel 模板。

### 文档格式要求

飞书文档中需包含 8 列表格：

| 位置 | 知识点 | 题目 | 答案 | 正反馈 | 负反馈 | 跳关条件 | 场景设置 |
|---|---|---|---|---|---|---|---|
| 关卡1-1 | 概率 | 下列事件中... | D | 太棒了 | 再想想 | 错一次过 | |

- 每道题一个表格，8 列顺序固定
- 题目列可包含多行（选项 A.xxx / B.xxx）
- 支持选择题（自动识别 A/B/C/D 选项）和填空题（含 ______ 标记）
- 跳关条件支持「错一次过」「错三次过」及数字范围

### 配置飞书 API（首次使用需 3 分钟）

1. 访问 [飞书开放平台](https://open.feishu.cn/app) → 创建企业自建应用
2. 添加权限：`doc:document:readonly`、`wiki:wiki:readonly`
3. 发布应用 → 管理后台审批
4. 在 `app/routers/feishu.py` 中填入 `APP_ID` 和 `APP_SECRET`

### API 端点

| 端点 | 用途 |
|---|---|
| `POST /api/feishu/convert` | 飞书文档 → Excel 转换 |
| `GET /api/feishu/download` | 下载生成的 Excel 文件 |

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
| test_client.py | 10 | HTTP重试机制 |
| test_mappings.py | 16 | 映射表正确性 |

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
