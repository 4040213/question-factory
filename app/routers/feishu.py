"""飞书文档 → Excel 转换路由（v3 — 维度感知 + 表头别名映射）。

核心改进：
1. 读取 table.property.column_size/row_size 按真实列数构建网格
2. 按表头名字（别名映射）取字段，不再硬编码列位置
3. 优先采信"题型"列判定题型；选项正则兼容"字母+空格"
4. 合并单元格续行（位置为空的行并入上题）
5. 图片按题聚合，下载时用编码文件名标注归属
6. summary 告警缺答案/纯图题/不支持题型
7. 新增莫顿素材库 URL 读回端点
"""

import re
import os
import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import httpx
import openpyxl
from fastapi import APIRouter, Query
from fastapi.responses import FileResponse
from pydantic import BaseModel

router = APIRouter(prefix="/api/feishu", tags=["飞书导入"])

APP_ID = os.getenv("FEISHU_APP_ID", "")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")

# ★ 环境变量未设置时，从配置文件读取
if not APP_ID or not APP_SECRET:
    _cred_file = os.path.join(os.path.dirname(__file__), "..", "..", "config", "feishu_credentials.json")
    try:
        with open(_cred_file, "r", encoding="utf-8") as f:
            _creds = json.load(f)
            APP_ID = _creds.get("feishu_app_id", "")
            APP_SECRET = _creds.get("feishu_app_secret", "")
    except Exception:
        pass

if not APP_ID or not APP_SECRET:
    print("[feishu] ⚠ 未配置 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")
else:
    print(f"[feishu] ✓ 已加载飞书凭证 APP_ID={APP_ID[:8]}***")

# ★ 模板路径
_WORKFLOW_DIR = os.getenv("FEISHU_WORKFLOW_DIR", r"D:\A工作\题目生成工作流")
CHOICE_TEMPLATE = os.path.join(_WORKFLOW_DIR, "选择题(模板) - 副本.xlsx")
FILL_TEMPLATE = os.path.join(_WORKFLOW_DIR, "填空题（模板）.xlsx")
OUTPUT_CHOICE = os.path.join(_WORKFLOW_DIR, "选择题_生成结果.xlsx")
OUTPUT_FILL = os.path.join(_WORKFLOW_DIR, "填空题_生成结果.xlsx")
IMAGE_DIR = os.path.join(_WORKFLOW_DIR, "飞书图片")

DEFAULT_RESOLUTION = "1334*750"
FEISHU_API = "https://open.feishu.cn/open-apis"

# ============================================================
#  表头别名 → 规范字段 映射表
#  按表头名字取字段，不再按位置。
#  同一个逻辑字段在不同文档里可能叫不同名字。
# ============================================================

HEADER_ALIASES: Dict[str, str] = {}

def _register_aliases(canonical: str, *aliases: str):
    for a in aliases:
        HEADER_ALIASES[a.strip()] = canonical

_register_aliases("location", "位置", "关卡", "题号", "名称")
_register_aliases("question_type", "题型", "类型")
_register_aliases("question", "题目", "题干", "内容")
_register_aliases("description", "配题说明", "说明", "配置")
_register_aliases("answer", "答案")
_register_aliases("feedback_combined", "正负反馈", "反馈")
_register_aliases("feedback_negative", "负反馈")
_register_aliases("feedback_positive", "正反馈")
_register_aliases("knowledge", "知识点")
_register_aliases("pass_condition", "跳关条件", "跳关")
_register_aliases("scene", "场景设置", "场景")
# 以下列存在但当前忽略
_register_aliases("_ignore", "小灯泡", "语音引导", "空")
# ★ 题干音频（朗读/配音）：独立于反馈音频，由 _write_*_excel 写入「题干音频」列
_register_aliases("stem_audio", "题干音频", "配音", "朗读音频", "题干配音", "narration", "朗读")
# ★ 标准编码（公司素材编码，如 26sp_G2W13-P1-1）：与「名称」(友好名) 互为别名，
#   用于匹配按公司编码命名的图片/音频文件（如 26sp_G2W13-P1-1_tips_1.mp3）。
#   文档里若已有该列，解析时自动带入；生成的 Excel 也会预留此列供填写。
_register_aliases("code", "标准编码", "素材编码", "题号编码", "编码", "公司编码")
# ★ 自动提示（小灯泡）：最多 3 个提示，每个 图片 + 音频 + 最多6个变量（每变量『状态:时间』）
#   与 question.py 解析 / 前端 collectAutoHints 共用同一命名约定，不写死。
#   第 1 个提示额外兼容旧单提示文档的无序号别名（自动提示图片 / 自动提示音频 / 自动提示变量N），
#   保证旧需求文档仍能把数据落到第 1 个提示。
_AUTO_HINT_COUNT = 3
_AUTO_HINT_VAR_COUNT = 6
_AUTO_HINT_CN = ["一", "二", "三", "四", "五", "六"]
for _h in range(1, _AUTO_HINT_COUNT + 1):
    _hcn = _AUTO_HINT_CN[_h - 1]
    _img_aliases = [f"自动提示{_h}图片", f"自动提示第{_hcn}图片", f"自动提示{_h}图"]
    _aud_aliases = [f"自动提示{_h}音频", f"自动提示第{_hcn}音频", f"自动提示{_h}音"]
    if _h == 1:  # 旧单提示文档兼容
        _img_aliases += ["自动提示图片", "自动提示图", "自动提示背景图", "自动提示图片url"]
        _aud_aliases += ["自动提示音频", "自动提示音", "自动提示音频url"]
    _register_aliases(f"auto_hint_{_h}_image", *_img_aliases)
    _register_aliases(f"auto_hint_{_h}_audio", *_aud_aliases)
    for _i in range(1, _AUTO_HINT_VAR_COUNT + 1):
        _icn = _AUTO_HINT_CN[_i - 1]
        _register_aliases(
            f"auto_hint_{_h}_var{_i}",
            f"自动提示{_h}变量{_i}", f"自动提示{_h}变量{_icn}",
            f"自动提示第{_hcn}个提示变量{_i}", f"自动提示第{_hcn}提示第{_icn}变量",
        )


class FeishuConvertRequest(BaseModel):
    url: str


class MortonReadbackRequest(BaseModel):
    """莫顿素材库 URL 读回请求。"""
    cookie: str
    token: str
    folder_id: int


# ============================================================
#  认证 & 文档定位
# ============================================================

def _get_client() -> Tuple[httpx.Client, str]:
    """获取 httpx 客户端和 tenant_access_token。"""
    client = httpx.Client(timeout=30)
    resp = client.post(
        f"{FEISHU_API}/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
    )
    data = resp.json()
    if data.get("code") != 0:
        client.close()
        raise Exception(f"飞书认证失败: {data}")
    return client, data["tenant_access_token"]


def _resolve_doc_id(client: httpx.Client, token: str, wiki_url: str) -> str:
    """从 wiki URL 解析出 document_id。"""
    wiki_token = wiki_url.rstrip("/").split("/wiki/")[-1]
    resp = client.get(
        f"{FEISHU_API}/wiki/v2/spaces/get_node?token={wiki_token}",
        headers={"Authorization": f"Bearer {token}"},
    ).json()
    if resp.get("code") != 0:
        raise Exception(f"Wiki 节点获取失败: {resp}")
    return resp["data"]["node"]["obj_token"]


# ============================================================
#  Blocks API — 维度感知的表格解析
# ============================================================

def _get_block_content(
    client: httpx.Client, token: str, doc_id: str, block_id: str
) -> Tuple[str, List[str]]:
    """递归获取一个 block 的文本内容和图片 token 列表。"""
    resp = client.get(
        f"{FEISHU_API}/docx/v1/documents/{doc_id}/blocks/{block_id}",
        headers={"Authorization": f"Bearer {token}"},
    ).json()
    blk = resp.get("data", {}).get("block", {})
    bt = blk.get("block_type")

    text = ""
    images: List[str] = []

    elements: List[Dict] = []
    if bt == 2:  # 文本块
        elements = blk.get("text", {}).get("elements", [])
    elif bt == 32:  # 表格单元格
        for body in blk.get("table_cell", []):
            elements.extend(body.get("elements", []))

    for elem in elements:
        tr = elem.get("text_run", {})
        if tr.get("content"):
            text += tr["content"]
        ie = elem.get("inline_image")
        if ie:
            it = ie.get("image_token", "")
            if it:
                images.append(it)

    if bt == 27:  # 独立图片块
        img = blk.get("image", {})
        it = img.get("token", "")
        if it:
            images.append(it)

    for cid in blk.get("children", []):
        ct, ci = _get_block_content(client, token, doc_id, cid)
        if ct:
            if text and not text.endswith("\n"):
                text += "\n"
            text += ct
        images.extend(ci)

    return text, images


def _fetch_all_tables(
    client: httpx.Client, token: str, doc_id: str
) -> List[Dict[str, Any]]:
    """通过 Blocks API 获取文档中所有表格，返回维度感知的表格数据。

    每个表格返回:
      {
        "table_id": str,
        "col_size": int,          # 从 table.property 读取
        "row_size": int,
        "grid": List[List[Dict]], # grid[row][col] = {text, images}
      }
    """
    # Step 1 — 获取文档根 block 的子节点
    root_resp = client.get(
        f"{FEISHU_API}/docx/v1/documents/{doc_id}/blocks/{doc_id}/children",
        headers={"Authorization": f"Bearer {token}"},
        params={"page_size": 100},
    ).json()
    items = root_resp.get("data", {}).get("items", [])
    table_ids = [it["block_id"] for it in items if it.get("block_type") == 31]

    all_tables: List[Dict[str, Any]] = []

    for tid in table_ids:
        # 获取表格 block 本体（读 property.column_size / row_size）
        tbl_resp = client.get(
            f"{FEISHU_API}/docx/v1/documents/{doc_id}/blocks/{tid}",
            headers={"Authorization": f"Bearer {token}"},
        ).json()
        tb_block = tbl_resp.get("data", {}).get("block", {})
        tbl_prop = tb_block.get("table", {}).get("property", {})
        col_size = tbl_prop.get("column_size", 0)
        row_size = tbl_prop.get("row_size", 0)

        if not col_size or not row_size:
            print(f"[feishu] ⚠ 表格 {tid} 无 col_size/row_size，跳过")
            continue

        # 获取表格的所有子 block（单元格）
        cells_resp = client.get(
            f"{FEISHU_API}/docx/v1/documents/{doc_id}/blocks/{tid}/children",
            headers={"Authorization": f"Bearer {token}"},
            params={"page_size": 500},
        ).json()
        raw_cells = cells_resp.get("data", {}).get("items", [])

        # 过滤出表格单元格（block_type == 32）
        cell_blocks = [c for c in raw_cells if c.get("block_type") == 32]

        # ★ 按 row_size × col_size 构建网格
        # 飞书表格单元格按行优先顺序排列
        grid: List[List[Dict[str, Any]]] = []
        cell_idx = 0
        for ri in range(row_size):
            row_cells: List[Dict[str, Any]] = []
            for ci in range(col_size):
                if cell_idx < len(cell_blocks):
                    cb = cell_blocks[cell_idx]
                    txt, imgs = _get_block_content(client, token, doc_id, cb["block_id"])
                    row_cells.append({"text": txt.strip(), "images": imgs})
                else:
                    row_cells.append({"text": "", "images": []})
                cell_idx += 1
            grid.append(row_cells)

        all_tables.append({
            "table_id": tid,
            "col_size": col_size,
            "row_size": row_size,
            "grid": grid,
        })
        print(f"[feishu] 表格 {tid}: {col_size}列×{row_size}行, {len(cell_blocks)} cells")

    return all_tables


# ============================================================
#  表头解析 + 题目组装
# ============================================================

def _parse_header(grid: List[List[Dict]]) -> Dict[str, int]:
    """解析表头行，返回 {canonical_field: column_index}。

    遍历表头每个单元格的文本，通过 HEADER_ALIASES 映射到规范字段。
    同一个规范字段如果出现多次，取第一次。
    """
    if not grid:
        return {}
    header_row = grid[0]
    mapping: Dict[str, int] = {}
    for ci, cell in enumerate(header_row):
        h = cell["text"].strip()
        if not h:
            continue
        canonical = HEADER_ALIASES.get(h)
        if canonical and canonical != "_ignore":
            if canonical not in mapping:
                mapping[canonical] = ci
    return mapping


def _assemble_questions(tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """将网格数据组装为题目 dict 列表。

    每个 table: 表头行 → 别名映射，数据行 → 按映射取字段。
    位置为空的行视为续行，合并到上一题。
    """
    questions: List[Dict[str, Any]] = []

    for table in tables:
        grid = table["grid"]
        col_size = table["col_size"]
        row_size = table["row_size"]

        if row_size < 2:
            continue

        field_map = _parse_header(grid)
        if not field_map:
            print(f"[feishu] ⚠ 表格 {table['table_id']} 无可识别表头，跳过")
            continue

        # 逐行解析数据（跳过表头）
        for ri in range(1, row_size):
            row = grid[ri]
            row_data: Dict[str, Any] = {}

            for canonical, ci in field_map.items():
                if ci < len(row):
                    cell = row[ci]
                    row_data[canonical] = cell["text"]
                    row_data[f"_images_{canonical}"] = cell["images"]

            # 检查是否是续行（位置为空）
            location = row_data.get("location", "").strip()
            if not location and questions:
                # 续行：合并到上一题
                _merge_continuation(questions[-1], row_data)
                continue

            # 汇总所有图片 token
            all_imgs: List[str] = []
            for key in list(row_data.keys()):
                if key.startswith("_images_"):
                    all_imgs.extend(row_data.pop(key))

            row_data["_images"] = all_imgs
            row_data["_image_paths"] = []
            row_data["_warnings"] = []

            _classify(row_data)
            questions.append(row_data)

    return questions


def _merge_continuation(prev_q: Dict, cont: Dict):
    """将续行数据合并到上一题。

    - 题目文本追加
    - 答案追加为新行
    - 图片汇总
    """
    # 追加题目文本
    cont_q = cont.get("question", "").strip()
    if cont_q:
        prev_q["question"] = (prev_q.get("question", "") + "\n" + cont_q).strip()

    # 追加答案
    cont_ans = cont.get("answer", "").strip()
    if cont_ans:
        prev_q["answer"] = (prev_q.get("answer", "") + "\n" + cont_ans).strip()

    # 追加配题说明
    cont_desc = cont.get("description", "").strip()
    if cont_desc:
        prev_q["description"] = (prev_q.get("description", "") + "\n" + cont_desc).strip()

    # 汇总图片
    for key in list(cont.keys()):
        if key.startswith("_images_"):
            prev_q["_images"].extend(cont.pop(key))


# ============================================================
#  题目分类（v3 — 优先采信题型列 + 扩展选项正则）
# ============================================================

def _clean_answer_line(text: str) -> str:
    """清洗答案行：去除 A./B./C./D. 或 1./2./3. 等序号前缀。"""
    text = text.strip()
    text = re.sub(r'^[（(]\s*[ABCDEFGH1-9]\s*[)）]\s*', '', text)
    text = re.sub(r'^[ABCDEFGH1-9]\s*[.、．)）]\s*', '', text)
    return text.strip()


def _classify(row: Dict):
    """分类题目类型，提取题干 / 选项 / 答案 / 跳关条件。

    优先级：
    1. "题型"列明确标注 → 直接采信
    2. 正则扫描选项标记 → 检测到 2+ 选项 → CHOICE
    3. 默认 → FILL
    """
    q_text = row.get("question", "")
    q_type_raw = row.get("question_type", "").strip()

    # ---------- 优先采信"题型"列 ----------
    if q_type_raw:
        type_lower = q_type_raw.lower()
        if "选择" in q_type_raw or "choice" in type_lower:
            row["_type"] = "CHOICE"
            _extract_options(row, q_text)
            _process_answer(row)
            _process_pass_condition(row)
            return
        elif "填空" in q_type_raw or "completion" in type_lower or "blank" in type_lower:
            row["_type"] = "FILL"
            row["_stem"] = q_text.split("\n")[0].strip() if q_text else ""
            row["_options"] = []
            _process_answer(row)
            _process_pass_condition(row)
            return
        elif "连线" in q_type_raw or "connect" in type_lower or "link" in type_lower:
            row["_type"] = "UNSUPPORTED"
            row["_stem"] = q_text.split("\n")[0].strip() if q_text else ""
            row["_options"] = []
            row["_answer"] = row.get("answer", "").strip()
            row["_warnings"].append(f"不支持题型「{q_type_raw}」，已跳过")
            return

    # ---------- 正则扫描选项标记 ----------
    # 兼容三种分隔符：
    #   A.牛奶盒     (字母+点号, 行首)
    #   B 可乐       (字母+空格, 行首)
    #   A. xxx  B. yyy  (同行用2+空格分隔)
    opt_markers: List[Tuple[int, str]] = []
    for m in re.finditer(r'(?:^|\n|\s{2,})\s*([ABCDEFGH])\s*[.、．]\s*', q_text):
        opt_markers.append((m.start(1), m.group(1)))
    for m in re.finditer(r'(?:^|\n|\s{2,})\s*([ABCDEFGH])\s+(\S)', q_text):
        pos, label = m.start(1), m.group(1)
        if not any(p == pos for p, _ in opt_markers):
            opt_markers.append((pos, label))

    # 验证选项标签按 A→B→C 顺序
    if len(opt_markers) >= 2:
        opt_markers.sort(key=lambda x: x[0])
        labels = [m[1] for m in opt_markers]
        expected = [chr(ord('A') + i) for i in range(len(labels))]
        if labels == expected:
            row["_type"] = "CHOICE"
            _extract_options_from_markers(row, q_text, opt_markers)
            _process_answer(row)
            _process_pass_condition(row)
            return

    # ---------- 默认填空题 ----------
    row["_type"] = "FILL"
    row["_stem"] = q_text.split("\n")[0].strip() if q_text else ""
    row["_options"] = []
    _process_answer(row)
    _process_pass_condition(row)

    # ★ 文本兜底：题目中含"连线"但无题型列 → 标记不支持
    if "连线" in q_text and not q_type_raw:
        row["_type"] = "UNSUPPORTED"
        row["_warnings"].append("题目含「连线」关键词，标记为不支持")
        return

    # 检查纯图题（题目为空但有图片）
    if not q_text.strip() and row.get("_images"):
        row["_warnings"].append("纯图题（题目为空，仅有图片）")


def _extract_options(row: Dict, q_text: str):
    """从题目文本中提取选项（题型列已明确选择题时调用）。"""
    opt_markers: List[Tuple[int, str]] = []
    for m in re.finditer(r'(?:^|\n|\s{2,})\s*([ABCDEFGH])\s*[.、．]\s*', q_text):
        opt_markers.append((m.start(1), m.group(1)))
    for m in re.finditer(r'(?:^|\n|\s{2,})\s*([ABCDEFGH])\s+(\S)', q_text):
        pos, label = m.start(1), m.group(1)
        if not any(p == pos for p, _ in opt_markers):
            opt_markers.append((pos, label))

    if len(opt_markers) >= 2:
        opt_markers.sort(key=lambda x: x[0])
        labels = [m[1] for m in opt_markers]
        expected = [chr(ord('A') + i) for i in range(len(labels))]
        if labels == expected:
            _extract_options_from_markers(row, q_text, opt_markers)
            return

    # 题型列说是选择题但正则没匹配到选项 → 可能选项在单独列或答案列
    row["_stem"] = q_text.strip()
    row["_options"] = [""] * 8
    row["_warnings"].append("题型标注为选择题但未检测到选项标记")


def _extract_options_from_markers(row: Dict, q_text: str, opt_markers: List[Tuple[int, str]]):
    """根据选项标记位置提取题干和选项内容。"""
    first_pos = opt_markers[0][0]
    stem = q_text[:first_pos].strip()
    stem = re.sub(r'[\s\n]*$', '', stem)
    row["_stem"] = stem

    options = [""] * 8
    for i, (pos, label) in enumerate(opt_markers):
        idx = ord(label) - ord("A")
        if idx >= 8:
            continue
        content_start = pos + len(label)
        if i + 1 < len(opt_markers):
            content_end = opt_markers[i + 1][0]
        else:
            content_end = len(q_text)
        content = q_text[content_start:content_end]
        # 清理：去除开头的 .、．、空格
        content = re.sub(r'^[.、．\s]+', '', content).strip()
        content = re.sub(r'\s+', ' ', content)
        options[idx] = content

    row["_options"] = options


def _process_answer(row: Dict):
    """处理答案字段。"""
    answer_raw = row.get("answer", "").strip()
    answer_parts = [a.strip() for a in answer_raw.split("\n") if a.strip()]

    if row["_type"] == "CHOICE":
        row["_answer"] = re.sub(r'\s+', '', answer_raw.upper())
        row["_answer2"] = ""
        row["_all_answers"] = [row["_answer"]]
    else:
        cleaned = [_clean_answer_line(a) for a in answer_parts]
        row["_all_answers"] = cleaned
        row["_answer"] = cleaned[0] if cleaned else ""
        row["_answer2"] = cleaned[1] if len(cleaned) > 1 else ""

    # 告警：缺答案
    if not row["_answer"] and not row.get("_images"):
        row["_warnings"].append("缺少答案")


def _process_pass_condition(row: Dict):
    """处理跳关条件。"""
    jump = row.get("pass_condition", "").strip()
    if "错一次过" in jump:
        row["_min_wrong"], row["_max_wrong"] = 1, 1
    elif "错三次过" in jump:
        row["_min_wrong"], row["_max_wrong"] = 1, 3
    else:
        nums = re.findall(r'\d+', jump)
        row["_min_wrong"] = int(nums[0]) if nums else 1
        row["_max_wrong"] = (
            int(nums[1]) if len(nums) > 1
            else int(nums[0]) if len(nums) == 1
            else 3
        )


# ============================================================
#  反馈清洗
# ============================================================

def _clean_feedback(raw: str) -> str:
    """清洗负反馈文本：去除「负反馈一：」「负反馈3：」等标签（中文/阿拉伯数字均支持）。"""
    if not raw or raw == "无":
        return raw
    # 阿拉伯数字：负反馈3： / 负反馈3:
    cleaned = re.sub(r'负反馈\s*\d+\s*[：:]\s*', '', raw)
    # 中文数字：负反馈一： / 负反馈二：
    cleaned = re.sub(r'负反馈[一二三四五六七八九十]+\s*[：:]\s*', '', cleaned)
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    return cleaned.strip()


def _clean_positive_feedback(raw: str) -> str:
    if not raw or raw == "无":
        return raw
    return raw.strip()


# 中文数字 → 阿拉伯数字（用于「负反馈三」这类写法）
_CN_DIGIT = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
             "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def _build_feedback_columns(neg_raw: str, pos_raw: str,
                            fallback_min: int = 1, fallback_max: int = 3) -> Dict[str, Any]:
    """从飞书负/正反馈文本，构建 Excel 反馈相关列的值。

    ★ 核心修复（消除「通用反馈 vs 条件反馈」重复弹窗）：
      莫顿反馈模型里，`wrongQuestionFeedback`(通用) 与 `questionFeedbacks`(条件)
      是两套独立触发逻辑。组课文档的「负反馈N：文字」语义是【第 N 次错才触发】，
      它本质是一条**条件反馈**，不应再被当作通用反馈写一遍。

      规则：
      - neg_raw 含「负反馈N：」层级 → 只产出【条件反馈】：
          条件反馈文本 = 清洗后文字；最少/最多错次 = N（N-N 单点触发）。
          不再产出 答错通用反馈（否则会与条件反馈内容重复，学生答错连弹两个一样窗口）。
      - neg_raw 是纯文字(无层级) → 产出【通用负反馈】：只写 答错通用反馈。
      - pos_raw = 『无』 → 答对通用反馈留空（不生成正反馈）。

    选择题(_write_choice_excel) 与 填空题(_write_fill_excel) 共用此函数，
    保证两种题型反馈行为完全一致。
    """
    neg = (neg_raw or "").strip()
    pos = (pos_raw or "").strip()

    general_wrong = ""
    cond_text = ""
    cond_min = fallback_min
    cond_max = fallback_max

    if neg and neg != "无":
        m = re.search(r'负反馈\s*(\d+|[一二三四五六七八九十]+)\s*[：:]', neg)
        if m:
            token = m.group(1)
            n = int(token) if token.isdigit() else _CN_DIGIT.get(token, 0)
            if n:
                cond_min = n
                cond_max = n
            # 条件反馈文本 = 清洗掉「负反馈N：」前缀后的纯文字
            cond_text = _clean_feedback(neg)
        else:
            # 无层级的纯文字 → 通用负反馈
            general_wrong = neg

    pos_fb = pos if (pos and pos != "无") else ""
    return {
        "general_wrong": general_wrong,
        "cond_text": cond_text,
        "cond_min": cond_min,
        "cond_max": cond_max,
        "pos_fb": pos_fb,
    }


# ============================================================
#  图片下载（编码文件名标注归属）
# ============================================================

def _download_images(questions: List[Dict], client: httpx.Client, token: str, doc_prefix: str = "") -> int:
    """下载题目中的图片到本地 IMAGE_DIR。

    文件名编码归属：{doc_prefix}-Q{num:02d}-{field}.png
    field 取值: stem(题目图)
    背景图/选项图通常由 Figma 导出、按命名规范 Q{nn}-bg / Q{nn}-A..H 提供，见前端匹配逻辑。
    注意：不存在独立的"答案图"——答案是图片时，它就是某个选项的图，靠 answer 字段标记正确选项。

    返回下载成功的图片数量。
    """
    os.makedirs(IMAGE_DIR, exist_ok=True)
    downloaded = 0

    for qi, q in enumerate(questions):
        paths: List[str] = []
        for img_idx, img_token in enumerate(q.get("_images", [])):
            # 飞书文档内嵌的图统一当题目图(stem)
            # （选项图/背景图来自 Figma，不在此处下载）
            field = "stem"

            filename = f"{doc_prefix}-Q{qi+1:02d}-{field}-{img_token[:8]}.png" if doc_prefix else f"Q{qi+1:02d}-{field}-{img_token[:8]}.png"
            local_path = os.path.join(IMAGE_DIR, filename)

            if os.path.exists(local_path):
                paths.append(local_path)
                downloaded += 1
                continue
            try:
                img_resp = client.get(
                    f"{FEISHU_API}/drive/v1/medias/{img_token}/download",
                    headers={"Authorization": f"Bearer {token}"},
                )
                if img_resp.status_code == 200:
                    with open(local_path, "wb") as f:
                        f.write(img_resp.content)
                    paths.append(local_path)
                    downloaded += 1
            except Exception as e:
                print(f"[feishu] 图片下载失败 {img_token}: {e}")
        q["_image_paths"] = paths

    return downloaded


# ============================================================
#  Excel 写入
# ============================================================

def _ensure_header_col(ws, name: str, after_col: int = 0) -> int:
    """确保工作表表头含指定列；不存在则在末尾追加，返回 1-based 列号。幂等。

    after_col: 若需追加，至少放到 after_col 之后（避免落在已知反馈列区 <=31 内，
    与 _write_choice_excel 的硬编码反馈列 27/28/29 冲突）。默认 0 表示直接追加到末列。
    """
    header = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])]
    if name in header:
        return header.index(name) + 1
    col = max(ws.max_column, after_col) + 1
    ws.cell(row=1, column=col, value=name)
    print(f"[EXCEL] 表头缺失列「{name}」，已在第 {col} 列追加(after_col={after_col})")
    return col


def _write_choice_excel(questions: List[Dict]) -> str:
    """将选择题写入 Excel 模板（31列）。

    列映射:
       1=题干  2=文本内容  3-10=选项A-H  11=答案  12=题目名称
       13=学科  14=分辨率  15=年级  16=阶段  17=难度  18=知识点
       19=题目图片  20=画布背景  21=答错通用反馈  22=答对通用反馈
       23=答错图片  24=答对图片  25=答错音频  26=答对音频
       27=最少错次  28=最多错次  29=条件反馈文本  30=条件图片URL  31=条件音频URL
    """
    wb = openpyxl.load_workbook(CHOICE_TEMPLATE)
    ws = wb.active
    # ★ 清除旧数据行（保留表头），避免残留测试数据
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    sr = 1  # 表头行
    # after_col=31：确保缺失时追加到反馈列区(<=31)之后，避免与硬编码反馈列 27/28/29 冲突
    _stem_audio_col = _ensure_header_col(ws, "题干音频", after_col=31)
    # ★ 标准编码（公司素材编码）：与「名称」(location) 互为别名，供按公司编码命名的素材文件匹配
    _code_col = _ensure_header_col(ws, "标准编码", after_col=_stem_audio_col)
    # ★ 自动提示（小灯泡）：最多 3 个提示，每个 图片/音频 + 变量1~6（追加于末尾）
    _auto_hint_img_cols: List[int] = []
    _auto_hint_aud_cols: List[int] = []
    _auto_hint_var_cols: List[List[int]] = []
    _prev = _stem_audio_col
    for _h in range(1, _AUTO_HINT_COUNT + 1):
        _ic = _ensure_header_col(ws, f"自动提示{_h}图片", after_col=_prev)
        _ac = _ensure_header_col(ws, f"自动提示{_h}音频", after_col=_ic)
        _vc: List[int] = []
        _p = _ac
        for _i in range(1, _AUTO_HINT_VAR_COUNT + 1):
            _c = _ensure_header_col(ws, f"自动提示{_h}变量{_i}", after_col=_p)
            _vc.append(_c)
            _p = _c
        _auto_hint_img_cols.append(_ic)
        _auto_hint_aud_cols.append(_ac)
        _auto_hint_var_cols.append(_vc)
        _prev = _p

    choice_qs = [q for q in questions if q["_type"] == "CHOICE"]
    for qi, q in enumerate(choice_qs):
        r = sr + qi + 1

        # 题干音频（朗读/配音）：独立于反馈音频，写到「题干音频」列
        stem_audio_val = q.get("stem_audio", "") or ""

        # 反馈处理（选择题/填空题共用同一套逻辑，保证一致）
        neg_raw = q.get("feedback_negative", "") or q.get("feedback_combined", "")
        pos_raw = q.get("feedback_positive", "")
        fb = _build_feedback_columns(neg_raw, pos_raw,
                                     q.get("_min_wrong", 1), q.get("_max_wrong", 3))

        # 分辨率
        scene = q.get("scene", "").strip()
        resolution = DEFAULT_RESOLUTION
        res_match = re.search(r'\d{3,4}\*\d{3,4}', scene)
        if res_match:
            resolution = res_match.group(0)

        d = {
            1: "选择题",
            2: q["_stem"],
            11: q["_answer"],
            12: q.get("location", ""),
            14: resolution,
            18: q.get("knowledge", ""),
            19: "\n".join(q.get("_image_paths", [])),
            21: fb["general_wrong"],
            22: fb["pos_fb"],
            27: fb["cond_min"],
            28: fb["cond_max"],
            29: fb["cond_text"],
        }
        if stem_audio_val:
            d[_stem_audio_col] = stem_audio_val
        # ★ 标准编码：从文档「标准编码」列带入（别名已注册），供公司编码命名的素材匹配
        _code_val = q.get("code", "") or ""
        if _code_val:
            d[_code_col] = _code_val

        # ★ 自动提示（小灯泡）：逐提示写入图片/音频/变量1~6（逐题优先；文档未配则留空）
        for _h in range(_AUTO_HINT_COUNT):
            _auto_img = q.get(f"auto_hint_{_h + 1}_image", "") or ""
            _auto_aud = q.get(f"auto_hint_{_h + 1}_audio", "") or ""
            if _auto_img:
                d[_auto_hint_img_cols[_h]] = _auto_img
            if _auto_aud:
                d[_auto_hint_aud_cols[_h]] = _auto_aud
            for _i in range(1, _AUTO_HINT_VAR_COUNT + 1):
                _v = q.get(f"auto_hint_{_h + 1}_var{_i}", "") or ""
                if _v:
                    d[_auto_hint_var_cols[_h][_i - 1]] = _v

        for oi, opt in enumerate(q.get("_options", [])):
            if opt:
                d[3 + oi] = opt

        # ★ 写到 max_column+1（而非固定 32）：覆盖反馈列区之后追加的「题干音频」列(如第32列)，
        #   否则题干音频值虽写入 d 字典却不会被 flush 到单元格（之前固定 range(1,32) 会漏写）。
        for col in range(1, ws.max_column + 1):
            val = d.get(col, "")
            if val or col <= 29:
                ws.cell(row=r, column=col).value = val if val else ""

    wb.save(OUTPUT_CHOICE)
    return OUTPUT_CHOICE


def _write_fill_excel(questions: List[Dict]) -> str:
    """将填空题写入 Excel 模板。

    列映射:
       1=题型(COMPLETION)  2=名称  3=学科  4=分辨率  5=年级  6=阶段  7=难度
       8=知识点  9=题干文本  10=文本  11=题干音频
       空N(从col12开始, 每空13列):
         col(12+(N-1)*13)+0=空N名称
         col(12+(N-1)*13)+1=空N类型
         col(12+(N-1)*13)+2=空N答案类型
         col(12+(N-1)*13)+3=空N输入框样式
         col(12+(N-1)*13)+4=空N答案
         col(12+(N-1)*13)+5=空N答案别名
         col(12+(N-1)*13)+6=空N键盘样式
         col(12+(N-1)*13)+7=空N显示序号
         col(12+(N-1)*13)+8=空N显示选项底框
         col(12+(N-1)*13)+9~12=空N选项1~4
    """
    wb = openpyxl.load_workbook(FILL_TEMPLATE)
    ws = wb.active
    # ★ 清除旧数据行（保留表头），避免残留测试数据
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    sr = 1  # 表头行
    # ★ 确保「题干音频」列存在（与选择题对齐；缺失则追加），供逐题朗读音频写入
    _stem_audio_col = _ensure_header_col(ws, "题干音频", after_col=11)
    # ★ 标准编码（公司素材编码）：与「名称」(location) 互为别名，供按公司编码命名的素材文件匹配
    _code_col = _ensure_header_col(ws, "标准编码", after_col=_stem_audio_col)
    # ★ 自动提示（小灯泡）：最多 3 个提示，每个 图片/音频 + 变量1~6（追加于末尾）
    _auto_hint_img_cols: List[int] = []
    _auto_hint_aud_cols: List[int] = []
    _auto_hint_var_cols: List[List[int]] = []
    _prev = 11
    for _h in range(1, _AUTO_HINT_COUNT + 1):
        _ic = _ensure_header_col(ws, f"自动提示{_h}图片", after_col=_prev)
        _ac = _ensure_header_col(ws, f"自动提示{_h}音频", after_col=_ic)
        _vc: List[int] = []
        _p = _ac
        for _i in range(1, _AUTO_HINT_VAR_COUNT + 1):
            _c = _ensure_header_col(ws, f"自动提示{_h}变量{_i}", after_col=_p)
            _vc.append(_c)
            _p = _c
        _auto_hint_img_cols.append(_ic)
        _auto_hint_aud_cols.append(_ac)
        _auto_hint_var_cols.append(_vc)
        _prev = _p

    # ★ 解析表头建立 列名→列号(1-based) 映射，稳健定位反馈/语音/图片新增列
    #   （模板末尾已追加：题目图片/背景图/答错通用反馈/答对通用反馈/.../条件反馈音频，
    #    与 parse_blank_excel 的 find_col 列名一致。旧模板无这些列时返回 -1，优雅跳过。）
    _header = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])]
    def _fill_col(name: str) -> int:
        return _header.index(name) + 1 if name in _header else -1

    fill_qs = [q for q in questions if q["_type"] == "FILL"]
    for qi, q in enumerate(fill_qs):
        r = sr + qi + 1
        q_text = q.get("question", "")

        # 统计填空数量
        blank_count = len(re.findall(r'(?:_{2,}|（\s*）)', q_text)) or 1

        all_answers = q.get("_all_answers", [])
        if not all_answers:
            ans = q.get("_answer", "")
            all_answers = [ans] if ans else []

        num_blanks = blank_count

        # 分辨率
        scene = q.get("scene", "").strip()
        resolution = DEFAULT_RESOLUTION
        res_match = re.search(r'\d{3,4}\*\d{3,4}', scene)
        if res_match:
            resolution = res_match.group(0)

        d = {
            1: "COMPLETION",
            2: q.get("location", ""),
            4: resolution,
            8: q.get("knowledge", ""),
            9: q.get("_stem", "")[:80],
            10: q_text,
        }

        for i in range(num_blanks):
            base_col = 12 + i * 13
            raw_answer = all_answers[i] if i < len(all_answers) else ""
            is_last = (i == num_blanks - 1)
            remaining = len(all_answers) - i

            opts: List[str] = []
            main_answer = ""
            aliases: List[str] = []

            sep = None
            for s in (";", ",", "，"):
                if s in raw_answer:
                    sep = s
                    break

            if sep is not None:
                parts = [_clean_answer_line(p) for p in raw_answer.split(sep) if p.strip()]
                main_answer = parts[0]
                aliases = parts[1:]
                opts = parts[:4]
                blank_type = "选项型"
                answer_type = "选项"
                input_style = "选项输入框"
            elif num_blanks == 1 and remaining > 1:
                candidates = all_answers[:4]
                opts = candidates
                main_answer = candidates[0]
                aliases = candidates[1:]
                blank_type = "选项型"
                answer_type = "选项"
                input_style = "选项输入框"
            elif is_last and remaining > 1:
                tail = all_answers[i:]
                opts = tail[:4]
                main_answer = tail[0]
                aliases = tail[1:]
                blank_type = "选项型"
                answer_type = "选项"
                input_style = "选项输入框"
            else:
                main_answer = raw_answer
                blank_type = "内容型"
                answer_type = "内容"
                input_style = "文本输入框"

            d.update({
                base_col + 0: f"空{i+1}",
                base_col + 1: blank_type,
                base_col + 2: answer_type,
                base_col + 3: input_style,
                base_col + 4: main_answer,
                base_col + 5: ";".join(aliases),
                base_col + 6: "横向双线键盘",
                base_col + 7: "是",
                base_col + 8: "否",
            })
            for oi in range(4):
                opt_val = opts[oi] if oi < len(opts) else ""
                if opt_val:
                    d[base_col + 9 + oi] = opt_val

        # ★ 反馈/语音/图片：与选择题 _write_choice_excel 共用 _build_feedback_columns，
        #   保证两种题型反馈行为完全一致（飞书文档仅含文字反馈 + 嵌入图，
        #   背景图/反馈图片/音频由前端从上传 CDN 文件自动匹配，此处不写）。
        neg_raw = q.get("feedback_negative", "") or q.get("feedback_combined", "")
        pos_raw = q.get("feedback_positive", "")
        fb = _build_feedback_columns(neg_raw, pos_raw,
                                     q.get("_min_wrong", 1), q.get("_max_wrong", 3))

        _fb_cols = {
            "题目图片": "\n".join(q.get("_image_paths", [])),
            "答错通用反馈": fb["general_wrong"],
            "答对通用反馈": fb["pos_fb"],
            "最少错次": fb["cond_min"],
            "最多错次": fb["cond_max"],
            "条件反馈文本": fb["cond_text"],
        }
        for _name, _val in _fb_cols.items():
            _c = _fill_col(_name)
            if _c > 0 and _val != "":
                d[_c] = _val

        # ★ 题干音频（朗读/配音）：独立于反馈音频，写到「题干音频」列
        _stem_val = q.get("stem_audio", "") or ""
        if _stem_val:
            d[_stem_audio_col] = _stem_val
        # ★ 标准编码：从文档「标准编码」列带入（别名已注册），供公司编码命名的素材匹配
        _code_val = q.get("code", "") or ""
        if _code_val:
            d[_code_col] = _code_val

        # ★ 自动提示（小灯泡）：逐提示写入图片/音频/变量1~6（逐题优先；文档未配则留空）
        for _h in range(_AUTO_HINT_COUNT):
            _auto_img = q.get(f"auto_hint_{_h + 1}_image", "") or ""
            _auto_aud = q.get(f"auto_hint_{_h + 1}_audio", "") or ""
            if _auto_img:
                d[_auto_hint_img_cols[_h]] = _auto_img
            if _auto_aud:
                d[_auto_hint_aud_cols[_h]] = _auto_aud
            for _i in range(1, _AUTO_HINT_VAR_COUNT + 1):
                _v = q.get(f"auto_hint_{_h + 1}_var{_i}", "") or ""
                if _v:
                    d[_auto_hint_var_cols[_h][_i - 1]] = _v

        for col, val in d.items():
            if val:
                ws.cell(row=r, column=col).value = val

    wb.save(OUTPUT_FILL)
    return OUTPUT_FILL


# ============================================================
#  莫顿素材库 URL 读回
# ============================================================

from app.config import get_morton_api  # 莫顿 API 域名（支持运行时切换测试/正式环境）
# ★ 2026-07-16 修正：实际路径拼写是 merton（不是 morton），接口名是 searchMaterials（复数）
MORTON_SEARCH_PATH = "/merton-backend/material/searchMaterials"


def _morton_list_materials(cookie: str, token: str, folder_id: int) -> List[Dict]:
    """调用莫顿素材库列表接口，返回素材记录列表（自动分页，拉全部）。

    接口(2026-07-16 实测修正):
      GET /merton-backend/material/searchMaterials
          ?parentId={folder_id}&pageNum=N&pageSize=24&general=0
    注:
      - 旧路径 /morton-backend/material/searchMaterial 已 404（拼写应为 merton）
      - 参数名: page→pageNum, size→pageSize
      - 响应结构兼容 records / list / content；分页兼容 pages / totalPages
    """
    headers = {
        "Cookie": cookie,
        "Authorization": token,
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Referer": f"{get_morton_api()}/fe_morton/morton/assetManager",
    }
    all_records: List[Dict] = []
    page = 1
    while True:
        try:
            resp = httpx.get(
                f"{get_morton_api()}{MORTON_SEARCH_PATH}",
                params={
                    "parentId": folder_id,
                    "pageNum": page,
                    "pageSize": 24,
                    "general": 0,
                },
                headers=headers,
                timeout=30,
                follow_redirects=True,
            )
        except httpx.RequestError as e:
            raise Exception(f"莫顿API请求失败(网络层): {type(e).__name__}: {e}")

        # ★ 先检查 HTTP 状态码
        if resp.status_code >= 400:
            raw_text = resp.text[:500] if resp.text else "(空响应体)"
            raise Exception(
                f"莫顿API返回 HTTP {resp.status_code}。"
                f"\n  URL: {get_morton_api()}{MORTON_SEARCH_PATH}?parentId={folder_id}&pageNum={page}"
                f"\n  响应前500字: {raw_text}"
            )

        # ★ 防御：确保响应体非空且为 JSON
        raw_body = resp.text
        if not raw_body or not raw_body.strip():
            raise Exception(
                f"莫顿API返回空响应体(HTTP {resp.status_code})。"
                f"\n  请检查: Cookie是否过期、Token是否有效、文件夹ID({folder_id})是否存在"
            )

        try:
            data = resp.json()
        except (ValueError, json.JSONDecodeError) as parse_err:
            preview = raw_body[:300].replace('\n', '\\n').replace('\r', '')
            raise Exception(
                f"莫顿API返回了非JSON内容。"
                f"\n  HTTP {resp.status_code} | Content-Type: {resp.headers.get('content-type', '未知')}"
                f"\n  响应前300字: {preview}"
                f"\n  常见原因: Cookie/Token已过期、被重定向到登录页、或接口路径变更"
            ) from parse_err

        # 兼容多种响应结构
        data_obj = data.get("data", data) if isinstance(data.get("data"), dict) else data
        recs = (
            data_obj.get("records")
            or data_obj.get("list")
            or data_obj.get("content")
            or []
        )
        total_pages = (
            data_obj.get("pages")
            or data_obj.get("totalPages")
            or 1
        )

        # 软校验：code 存在且非 200 且无数据 → 报错
        code = data.get("code")
        if code is not None and code != 200 and not recs:
            raise Exception(
                f"莫顿素材库请求失败(code={code}): {data.get('message', '未知错误')}"
                f" | data={str(data)[:200]}"
            )

        all_records.extend(recs)
        if page >= total_pages or not recs:
            break
        page += 1
    return all_records


# ============================================================
#  路由端点
# ============================================================

@router.post("/convert")
async def feishu_convert(body: FeishuConvertRequest) -> Dict[str, Any]:
    """将飞书文档转换为 Excel 模板。"""
    try:
        client, token = _get_client()
        doc_id = _resolve_doc_id(client, token, body.url)

        # 维度感知解析
        tables = _fetch_all_tables(client, token, doc_id)
        questions = _assemble_questions(tables)

        if not questions:
            return {"success": False, "error": "未解析到任何题目"}

        # 生成文档前缀（用于图片文件名编码）
        doc_prefix = doc_id[:8]

        # 下载图片
        img_count = _download_images(questions, client, token, doc_prefix)

        # 写入 Excel
        files: List[str] = []
        choice_count = sum(1 for q in questions if q["_type"] == "CHOICE")
        fill_count = sum(1 for q in questions if q["_type"] == "FILL")
        unsupported_count = sum(1 for q in questions if q["_type"] == "UNSUPPORTED")

        if choice_count:
            files.append(_write_choice_excel(questions))
        if fill_count:
            files.append(_write_fill_excel(questions))

        # 构建摘要（含告警）
        summary_items = []
        total_warnings = 0
        for q in questions:
            warnings = q.get("_warnings", [])
            total_warnings += len(warnings)
            summary_items.append({
                "位置": q.get("location", ""),
                "题型": q["_type"],
                "题干": (q.get("_stem", "") or "")[:60] + ("..." if len(q.get("_stem", "")) > 60 else ""),
                "答案": q.get("_answer", ""),
                "图片": len(q.get("_images", [])),
                "告警": warnings if warnings else "",
            })

        return {
            "success": True,
            "summary": {
                "total": len(questions),
                "choice": choice_count,
                "fill": fill_count,
                "unsupported": unsupported_count,
                "images_downloaded": img_count,
                "warnings": total_warnings,
                "items": summary_items,
            },
            "files": files,
        }

    except Exception as e:
        return {"success": False, "error": str(e)}


@router.get("/download")
async def feishu_download(path: str = Query(...)):
    """下载已生成的 Excel 文件。"""
    fp = Path(path)
    if not fp.exists():
        return {"error": "文件不存在"}
    return FileResponse(
        fp,
        filename=fp.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@router.post("/morton-readback")
async def morton_readback(body: MortonReadbackRequest) -> Dict[str, Any]:
    """从莫顿素材库读回素材 URL 列表。

    返回 [{name, url, type}, ...]，前端可按 name 反查文件名→URL 映射。
    """
    try:
        records = _morton_list_materials(body.cookie, body.token, body.folder_id)
        materials = []
        for r in records:
            url = r.get("url", "")
            if url and not url.startswith("http"):
                url = "https:" + url
            materials.append({
                "name": r.get("name", ""),
                "url": url,
                "type": r.get("type", ""),
                "id": r.get("id", 0),
            })
        return {
            "success": True,
            "folder_id": body.folder_id,
            "total": len(materials),
            "materials": materials,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# ============================================================
#  莫顿素材库上传（方案 A：拖入文件自动推送到莫顿 CDN）
# ============================================================

# 候选上传路径（按优先级尝试，第一个成功即返回）
# 注意：莫顿可能不支持 API 上传，这些路径是基于搜索接口 /merton-backend/material/searchMaterials 推测的
_MORTON_UPLOAD_PATHS = [
    "/merton-backend/material/upload",
    "/merton-backend/material/addMaterial",
    "/fe_morton/morton/assetManager/upload",
]
# 每条路径的超时（秒）— 短超时确保快速失败
_MORTON_UPLOAD_TIMEOUT = 5


class MortonUploadRequest(BaseModel):
    """莫顿素材库上传请求。"""
    cookie: str
    token: str
    folder_id: int
    file_name: str  # 文件名（不含路径）
    file_data_b64: str  # Base64 编码的文件内容


def _morton_upload_single(
    cookie: str, token: str, folder_id: int,
    file_name: str, file_bytes: bytes,
) -> Dict[str, Any]:
    """尝试将单个文件上传到莫顿素材库。只试第一条候选路径（快速失败）。

    鉴于莫顿大概率没有开放上传 API，多试路径只会增加等待时间，
    故只试最可能的一条，5 秒超时后立即返回结论，便于前端快速引导手动流程。
    """
    headers = {
        "Cookie": cookie,
        "Authorization": token,
        "Referer": f"{get_morton_api()}/fe_morton/morton/assetManager",
    }

    # 只试第一条候选路径（快速失败，不逐个耗尽所有候选）
    path = _MORTON_UPLOAD_PATHS[0]
    url = f"{get_morton_api()}{path}"
    try:
        files = {"file": (file_name, file_bytes, _guess_mime_type(file_name))}
        data = {"parentId": str(folder_id), "general": "0"}
        resp = httpx.post(url, files=files, data=data, headers=headers, timeout=_MORTON_UPLOAD_TIMEOUT)
    except Exception as e:
        # 连接失败（API 不存在）→ 立即返回失败，前端引导手动流程
        return {
            "success": False,
            "url": None,
            "material_id": None,
            "tried_paths": [path],
            "error": f"上传请求失败（{type(e).__name__}）：莫顿可能不支持 API 上传，请走手动流程。",
        }

    if resp.status_code == 200 or resp.status_code == 201:
        try:
            result = resp.json()
            obj = result.get("data", result) if isinstance(result.get("data"), dict) else result
            cdn_url = (
                obj.get("url")
                or (obj.get("data", {}).get("url") if isinstance(obj.get("data"), dict) else None)
            )
            mat_id = obj.get("id") or (obj.get("data", {}).get("id") if isinstance(obj.get("data"), dict) else None)
            return {
                "success": True,
                "url": cdn_url,
                "material_id": mat_id,
                "path_used": path,
                "status_code": resp.status_code,
            }
        except Exception:
            pass

    # 非 200/201 → 莫顿不支持该上传 API
    return {
        "success": False,
        "url": None,
        "material_id": None,
        "tried_paths": [path],
        "error": f"上传返回 HTTP {resp.status_code}：莫顿可能不支持 API 上传，请走手动流程。",
    }


def _guess_mime_type(filename: str) -> str:
    """根据扩展名猜测 MIME 类型。"""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    mimes = {
        "mp3": "audio/mpeg", "wav": "audio/wav", "m4a": "audio/mp4",
        "ogg": "audio/ogg", "aac": "audio/aac", "flac": "audio/flac",
        "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "gif": "image/gif", "svg": "image/svg+xml", "webp": "image/webp",
    }
    return mimes.get(ext, "application/octet-stream")


@router.post("/morton-upload")
async def morton_upload(body: MortonUploadRequest) -> Dict[str, Any]:
    """将单个语音/图片文件上传到莫顿素材库。

    前端流程：用户拖入文件 → 调此接口逐个上传 → 拿回 CDN URL → 自动匹配。
    若莫顿不支持 API 上传则返回友好错误 + 手动操作指引。
    """
    try:
        import base64
        file_bytes = base64.b64decode(body.file_data_b64)

        result = _morton_upload_single(
            cookie=body.cookie, token=body.token,
            folder_id=body.folder_id, file_name=body.file_name,
            file_bytes=file_bytes,
        )

        if result["success"]:
            return {
                "success": True,
                "file_name": body.file_name,
                "url": result["url"],
                "material_id": result["material_id"],
                "path_used": result.get("path_used"),
            }
        else:
            return {
                "success": False,
                "error": result.get("error", "上传失败"),
                "file_name": body.file_name,
                "tried_paths": result.get("tried_paths", []),
                "_manual_fallback": True,
                "_hint": (
                    "莫顿暂不支持自动上传。请手动将文件上传到莫顿素材库对应文件夹，"
                    "然后点击「读回 URL」即可自动识别匹配。"
                ),
            }
    except Exception as e:
        return {"success": False, "error": str(e), "file_name": body.file_name}
