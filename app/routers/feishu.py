"""飞书文档 → Excel 转换路由（v2 — Blocks API 结构化解析）。"""

import re
import os
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

if not APP_ID or not APP_SECRET:
    print("[feishu] ⚠ 未配置 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")

# ★ 模板路径：通过环境变量 FEISHU_WORKFLOW_DIR 配置，默认保留现有路径
_WORKFLOW_DIR = os.getenv("FEISHU_WORKFLOW_DIR", r"D:\A工作\题目生成工作流")
CHOICE_TEMPLATE = os.path.join(_WORKFLOW_DIR, "选择题(模板) - 副本.xlsx")
FILL_TEMPLATE = os.path.join(_WORKFLOW_DIR, "填空题（模板）.xlsx")
OUTPUT_CHOICE = os.path.join(_WORKFLOW_DIR, "选择题_生成结果.xlsx")
OUTPUT_FILL = os.path.join(_WORKFLOW_DIR, "填空题_生成结果.xlsx")
IMAGE_DIR = os.path.join(_WORKFLOW_DIR, "飞书图片")

# ★ 默认分辨率
DEFAULT_RESOLUTION = "1334*750"

HEADER_ORDER = ["位置", "知识点", "题目", "答案", "正反馈", "负反馈", "跳关条件", "场景设置"]

FEISHU_API = "https://open.feishu.cn/open-apis"


class FeishuConvertRequest(BaseModel):
    url: str


# ============================================================
#  认证 & 文档定位
# ============================================================

def _get_client() -> Tuple[httpx.Client, str]:
    """获取 httpx 客户端和 tenant_access_token。"""
    client = httpx.Client(timeout=30)
    resp = client.post(
        f"{FEISHU_API}/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
    ).json()
    if resp.get("code") != 0:
        raise Exception(f"飞书认证失败: {resp}")
    return client, resp["tenant_access_token"]


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
#  Blocks API — 结构化表格解析
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

    # 提取文本元素和行内图片
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

    # 独立图片块（block_type=27）
    if bt == 27:
        img = blk.get("image", {})
        it = img.get("token", "")
        if it:
            images.append(it)

    # 递归处理子 block（用换行分隔，保留文档段落结构）
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
) -> List[List[Dict[str, Any]]]:
    """通过 Blocks API 获取文档中所有表格的单元格数据。

    每个表格返回 cell 列表，每个 cell = {text, images}。
    表格的 cell 按 8 个一组（一行），第一行是表头。
    """
    # Step 1 — 获取文档根 block 的子节点
    root_resp = client.get(
        f"{FEISHU_API}/docx/v1/documents/{doc_id}/blocks/{doc_id}/children",
        headers={"Authorization": f"Bearer {token}"},
        params={"page_size": 100},
    ).json()
    items = root_resp.get("data", {}).get("items", [])
    table_ids = [it["block_id"] for it in items if it.get("block_type") == 31]

    all_tables: List[List[Dict]] = []

    for tid in table_ids:
        tbl_resp = client.get(
            f"{FEISHU_API}/docx/v1/documents/{doc_id}/blocks/{tid}/children",
            headers={"Authorization": f"Bearer {token}"},
            params={"page_size": 200},
        ).json()
        cells = tbl_resp.get("data", {}).get("items", [])

        table_cells: List[Dict[str, Any]] = []
        for cell in cells:
            if cell.get("block_type") != 32:
                continue
            txt, imgs = _get_block_content(client, token, doc_id, cell["block_id"])
            table_cells.append({"text": txt.strip(), "images": imgs})

        if table_cells:
            all_tables.append(table_cells)

    return all_tables


def _assemble_questions(tables: List[List[Dict]]) -> List[Dict]:
    """将表格 cell 数据组装为题目 dict 列表。

    每个表格第一行（前 8 个 cell）是表头，后面的每 8 个 cell 是一行数据。
    """
    questions: List[Dict] = []
    for table_cells in tables:
        num_rows = len(table_cells) // 8
        for ri in range(1, num_rows):  # 跳过表头（ri=0）
            base = ri * 8
            if base + 8 > len(table_cells):
                break
            row: Dict[str, Any] = {}
            for ci, key in enumerate(HEADER_ORDER):
                cell = table_cells[base + ci]
                row[key] = cell["text"]
                row[f"_images_{key}"] = cell["images"]
            # 汇总所有图片 token
            all_imgs: List[str] = []
            for key in HEADER_ORDER:
                all_imgs.extend(row.pop(f"_images_{key}", []))
            row["_images"] = all_imgs
            _classify(row)
            questions.append(row)
    return questions


# ============================================================
#  题目分类（改进版）
# ============================================================

def _classify(row: Dict):
    """分类题目类型，提取题干 / 选项 / 答案 / 跳关条件。"""
    q_text = row.get("题目", "")
    # ★ 默认改为 FILL（填空题），只有明确检测到多选项时才设为 CHOICE
    row["_type"] = "FILL"

    # ---------- 选项提取 ----------
    # 策略：用统一的「选项扫描器」提取所有 A. / B. / C. 开头的选项
    #   1) 找到所有选项标签（位置 + 标签字母）
    #   2) 按标签位置排序
    #   3) 每个选项内容 = 标签位置 → 下一个标签位置（或文本末尾）
    #   4) 题干 = 第一个选项之前的所有文本
    # ★ 关键改进：选项标签必须在行首或单独成行，避免将正文中的
    #    "A. xxx" 类文本误判为选项（常见于填空题的题目描述中）

    opt_markers: List[Tuple[int, str]] = []  # (position, label)
    for m in re.finditer(r'(?:^|\n)([ABCDEFGH])\s*[.、．]', q_text):
        pos = m.start(1)  # 标签字母位置
        label = m.group(1)
        opt_markers.append((pos, label))

    # ★ 额外验证：选项标签必须按 A→B→C→... 顺序出现，防止误判
    if len(opt_markers) >= 2:
        labels = [m[1] for m in opt_markers]
        expected = [chr(ord('A') + i) for i in range(len(labels))]
        if labels == expected:
            row["_type"] = "CHOICE"
            # 题干 = 第一个选项之前的部分
            first_pos = opt_markers[0][0]
            stem = q_text[:first_pos].strip()
            # 去掉末尾的换行和空白
            stem = re.sub(r'[\s\n]*$', '', stem)
            row["_stem"] = stem

            options = [""] * 8
            for i, (pos, label) in enumerate(opt_markers):
                idx = ord(label) - ord("A")
                if idx >= 8:
                    continue
                # 选项内容 = 从标签末尾到下一个标签之前
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
        else:
            # 标签不按顺序 → 不是选择题，按填空题处理
            row["_type"] = "FILL"
            row["_stem"] = q_text.split("\n")[0].strip()
            row["_options"] = []
    else:
        # 无选项 → 填空题
        row["_type"] = "FILL"
        row["_stem"] = q_text.split("\n")[0].strip()
        row["_options"] = []

    # ---------- 答案处理 ----------
    answer_raw = row.get("答案", "").strip()
    # 填空题答案可能是多行（如 关卡2-1: 蓝球 / 白球）
    answer_parts = [a.strip() for a in answer_raw.split("\n") if a.strip()]
    if row["_type"] == "CHOICE":
        row["_answer"] = re.sub(r'\s+', '', answer_raw.upper())
        row["_answer2"] = ""
    else:
        row["_answer"] = answer_parts[0] if answer_parts else ""
        row["_answer2"] = answer_parts[1] if len(answer_parts) > 1 else ""

    # ---------- 跳关条件 ----------
    jump = row.get("跳关条件", "").strip()
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
#  负反馈清洗
# ============================================================

def _clean_feedback(raw: str) -> str:
    """清洗负反馈文本：去除「负反馈一：」「负反馈二：」等标签。

    将多条反馈用空行分隔，保留纯文本内容。
    """
    if not raw or raw == "无":
        return raw
    # 去除前缀标签
    cleaned = re.sub(r'负反馈[一二三四五六七八九十]+[：:]\s*', '', raw)
    # 清理多余空白
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    return cleaned.strip()


def _clean_positive_feedback(raw: str) -> str:
    """清洗正反馈文本。"""
    if not raw or raw == "无":
        return raw
    return raw.strip()


# ============================================================
#  Excel 写入（改进版）
# ============================================================

def _write_choice_excel(questions: List[Dict]) -> str:
    """将选择题写入 Excel 模板。

    模板列布局（1-indexed openpyxl）:
       1=题干  2=文本内容  3-10=选项A-H  11=答案  12=题目名称
       13=学科  14=分辨率  15=年级  16=阶段  17=难度  18=知识点
       19=题目图片  20=画布背景
       21=答错通用反馈  22=答对通用反馈
       23=答错图片  24=答对图片  25=答错音频  26=答对音频
       27=最少错次  28=最多错次  29=条件反馈文本  30=条件图片  31=条件音频
    """
    wb = openpyxl.load_workbook(CHOICE_TEMPLATE)
    ws = wb.active
    sr = ws.max_row

    choice_qs = [q for q in questions if q["_type"] == "CHOICE"]
    for qi, q in enumerate(choice_qs):
        r = sr + qi + 1

        # 负反馈：清洗后放入条件反馈（第29列）+ 通用答错反馈（第21列）
        neg_raw = q.get("负反馈", "")
        pos_raw = q.get("正反馈", "")

        # 条件反馈文本（洗去"负反馈一/二/三"标签）
        cond_text = _clean_feedback(neg_raw) if neg_raw and neg_raw != "无" else ""

        # 正反馈作为通用答对
        pos_fb = pos_raw.strip() if pos_raw and pos_raw != "无" else ""

        # 条件反馈 min/max
        min_w = q.get("_min_wrong", 1)
        max_w = q.get("_max_wrong", 3)

        # ★ 同时写入通用答错（第21列）和条件反馈（第27-29列）
        # 通用答错：取第一级反馈文本（最简短的那条）
        general_wrong = ""
        if neg_raw and neg_raw != "无":
            parts = re.split(r'负反馈[一二三四五六七八九十]+[：:]', neg_raw)
            # parts[0] 是空（前面没内容），parts[1] 是第一级
            if len(parts) > 1:
                general_wrong = parts[1].strip()

        # ★ 分辨率：固定默认 1334*750，场景设置可覆盖
        scene = q.get("场景设置", "").strip()
        resolution = DEFAULT_RESOLUTION
        res_match = re.search(r'\d{3,4}\*\d{3,4}', scene)
        if res_match:
            resolution = res_match.group(0)

        d = {
            1: "选择题",                                  # 题干文本（简短概括）
            2: q["_stem"],                               # 文本内容 = 题目描述
            11: q["_answer"],                             # 答案
            12: q.get("位置", ""),                         # ★ 题目名称 = 位置字段
            14: resolution,                               # ★ 分辨率（默认 1334*750）
            18: q.get("知识点", ""),                       # 知识点
            19: "\n".join(q.get("_image_paths", [])),     # 题目图片（本地路径，下载失败的已跳过）
            21: general_wrong,                             # 答错通用反馈
            22: pos_fb,                                   # 答对通用反馈
            27: min_w,                                    # 最少错次
            28: max_w,                                    # 最多错次
            29: cond_text,                                # ★ 条件反馈文本
        }

        # 选项 A-H → 列 3-10
        for oi, opt in enumerate(q.get("_options", [])):
            if opt:
                d[3 + oi] = opt

        # 写入所有列（含空值，确保 openpyxl 不截断）
        for col in range(1, 32):
            val = d.get(col, "")
            if val or col <= 29:  # 前29列至少写空值保证列宽
                ws.cell(row=r, column=col).value = val if val else ""

    wb.save(OUTPUT_CHOICE)
    return OUTPUT_CHOICE


def _write_fill_excel(questions: List[Dict]) -> str:
    """将填空题写入 Excel 模板。

    列映射: 1=题型, 2=名称, 3=学科, 4=分辨率, 5=年级, 6=阶段, 7=难度, 8=知识点,
             9=题干文本, 10=文本,
             12-23=空1配置, 25-36=空2配置
    """
    wb = openpyxl.load_workbook(FILL_TEMPLATE)
    ws = wb.active
    sr = ws.max_row

    fill_qs = [q for q in questions if q["_type"] == "FILL"]
    for qi, q in enumerate(fill_qs):
        r = sr + qi + 1
        q_text = q.get("题目", "")
        # 统计填空数量（__ 或 （空格））
        blank_count = len(re.findall(r'(?:_{2,}|（\s*）)', q_text)) or 1

        # ★ 分辨率：固定默认 1334*750，场景设置可覆盖
        scene = q.get("场景设置", "").strip()
        resolution = DEFAULT_RESOLUTION
        res_match = re.search(r'\d{3,4}\*\d{3,4}', scene)
        if res_match:
            resolution = res_match.group(0)

        d = {
            1: "COMPLETION",
            2: q.get("位置", ""),           # ★ 名称
            3: q.get("知识点", ""),          # 学科（用知识点填充）
            4: resolution,                  # ★ 分辨率
            9: q["_stem"][:80],            # 题干文本
            10: q_text,                      # 文本
        }

        # 空1 配置
        if blank_count >= 1:
            d.update({
                12: "空1", 13: "内容型", 14: "内容", 15: "文本输入框",
                16: q["_answer"],
                18: "横向双线键盘", 19: "是", 20: "否",
            })

        # 空2 配置
        if blank_count >= 2:
            d.update({
                25: "空2", 26: "内容型", 27: "内容", 28: "文本输入框",
                29: q.get("_answer2", ""),
                31: "横向双线键盘", 32: "是", 33: "否",
            })

        for col, val in d.items():
            if val:
                ws.cell(row=r, column=col).value = val

    wb.save(OUTPUT_FILL)
    return OUTPUT_FILL


# ============================================================
#  图片下载
# ============================================================

def _download_images(questions: List[Dict], client: httpx.Client, token: str) -> int:
    """下载题目中的图片到本地 IMAGE_DIR，并将本地路径存入 _image_paths。

    返回下载成功的图片数量。
    """
    os.makedirs(IMAGE_DIR, exist_ok=True)
    downloaded = 0

    for q in questions:
        paths: List[str] = []
        for img_token in q.get("_images", []):
            local_path = os.path.join(IMAGE_DIR, f"{img_token}.png")
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
#  路由端点
# ============================================================

@router.post("/convert")
async def feishu_convert(body: FeishuConvertRequest) -> Dict[str, Any]:
    """将飞书文档转换为 Excel 模板。"""
    try:
        client, token = _get_client()
        doc_id = _resolve_doc_id(client, token, body.url)

        # 结构化解析
        tables = _fetch_all_tables(client, token, doc_id)
        questions = _assemble_questions(tables)

        if not questions:
            return {"success": False, "error": "未解析到任何题目"}

        # 下载图片
        img_count = _download_images(questions, client, token)

        # 写入 Excel
        files: List[str] = []
        choice_count = sum(1 for q in questions if q["_type"] == "CHOICE")
        fill_count = sum(1 for q in questions if q["_type"] == "FILL")

        if choice_count:
            files.append(_write_choice_excel(questions))
        if fill_count:
            files.append(_write_fill_excel(questions))

        # 构建摘要
        summary_items = []
        for q in questions:
            summary_items.append({
                "位置": q.get("位置", ""),
                "题型": q["_type"],
                "题干": q["_stem"][:60] + ("..." if len(q.get("_stem", "")) > 60 else ""),
                "答案": q.get("_answer", ""),
                "图片": len(q.get("_images", [])),
            })

        return {
            "success": True,
            "summary": {
                "total": len(questions),
                "choice": choice_count,
                "fill": fill_count,
                "images_downloaded": img_count,
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
