"""飞书文档 → Excel 转换路由。"""

import re
import os
from pathlib import Path
from typing import Any, Dict, List

import requests
import openpyxl
from fastapi import APIRouter, Query
from fastapi.responses import FileResponse
from pydantic import BaseModel

router = APIRouter(prefix="/api/feishu", tags=["飞书导入"])

APP_ID = os.getenv("FEISHU_APP_ID", "")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")

if not APP_ID or not APP_SECRET:
    print("[feishu] ⚠ 未配置 FEISHU_APP_ID / FEISHU_APP_SECRET 环境变量")

CHOICE_TEMPLATE = r"D:\A工作\题目生成工作流\选择题(模板) - 副本.xlsx"
FILL_TEMPLATE = r"D:\A工作\题目生成工作流\填空题（模板）.xlsx"
OUTPUT_CHOICE = r"D:\A工作\题目生成工作流\选择题_生成结果.xlsx"
OUTPUT_FILL = r"D:\A工作\题目生成工作流\填空题_生成结果.xlsx"

HEADER_ORDER = ["位置", "知识点", "题目", "答案", "正反馈", "负反馈", "跳关条件", "场景设置"]


class FeishuConvertRequest(BaseModel):
    url: str


def _fetch_raw_content(wiki_url: str) -> str:
    wiki_token = wiki_url.rstrip("/").split("/wiki/")[-1]
    token_resp = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}, timeout=15,
    ).json()
    if token_resp["code"] != 0:
        raise Exception(f"飞书认证失败: {token_resp}")
    token = token_resp["tenant_access_token"]

    wiki_resp = requests.get(
        f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}",
        headers={"Authorization": f"Bearer {token}"}, timeout=15,
    ).json()
    if wiki_resp["code"] != 0:
        raise Exception(f"Wiki 节点获取失败: {wiki_resp}")
    doc_id = wiki_resp["data"]["node"]["obj_token"]

    raw_resp = requests.get(
        f"https://open.feishu.cn/open-apis/docx/v1/documents/{doc_id}/raw_content",
        headers={"Authorization": f"Bearer {token}"}, timeout=15,
    ).json()
    return raw_resp["data"]["content"]


def _parse_questions(raw_text: str) -> List[Dict]:
    chunks = re.split(r'\n\s*\n', raw_text)
    questions = []
    i = 0
    while i < len(chunks):
        chunk = chunks[i].strip()
        if not chunk or re.match(r'^(P\d+|20\d{2})', chunk):
            i += 1; continue
        if chunk == "位置":
            if not all(i+j<len(chunks) and chunks[i+j].strip()==HEADER_ORDER[j] for j in range(8)):
                i += 1; continue
            i += 8
            row: Dict[str, str] = {}
            for j in range(8):
                while i < len(chunks) and not chunks[i].strip(): i += 1
                row[HEADER_ORDER[j]] = chunks[i].strip() if i < len(chunks) else ""
                if i < len(chunks): i += 1
            _classify(row); questions.append(row); continue
        i += 1
    return questions


def _classify(row: Dict):
    q_text = row.get("题目", ""); row["_type"] = "CHOICE"
    opt_pattern = re.search(r'\n\s*[ABCDEFGH]\s*[.、．]', q_text)
    if opt_pattern:
        fp = q_text.index("\n" + opt_pattern.group().strip()[0]); row["_stem"] = q_text[:fp].strip()
        options = [""] * 8
        for label, content in re.findall(r'\n\s*([ABCDEFGH])\s*[.、．]\s*(.+?)(?=\n\s*[ABCDEFGH]\s*[.、．]|\Z)', q_text, re.DOTALL):
            idx = ord(label.strip()) - ord('A')
            if idx < 8: options[idx] = re.sub(r'\s+', ' ', content.strip())
        row["_options"] = options
    else:
        inline = re.findall(r'([ABCDEFGH])\s*[.、．]?\s*(\S.*?)(?=\s+[ABCDEFGH]\s*|\Z)', q_text)
        if len(inline) >= 2:
            options = [""] * 8
            for label, content in inline:
                idx = ord(label) - ord('A')
                if idx < 8: options[idx] = content.strip()
            row["_options"] = options; e = q_text.index(inline[0][0]+inline[0][1][:2]); row["_stem"] = q_text[:e].strip()
        else:
            row["_type"] = "FILL"; row["_stem"] = q_text.split("\n")[0]; row["_options"] = []

    answer = re.sub(r'\s+', '', row.get("答案", "").strip().upper())
    parts = [a for a in answer.split("\n") if a.strip()]
    row["_answer"] = parts[0] if parts else ""; row["_answer2"] = parts[1] if len(parts) > 1 else ""
    jump = row.get("跳关条件", "").strip()
    if "错一次过" in jump: row["_min_wrong"], row["_max_wrong"] = 1, 1
    elif "错三次过" in jump: row["_min_wrong"], row["_max_wrong"] = 1, 3
    else:
        nums = re.findall(r'\d+', jump)
        row["_min_wrong"] = int(nums[0]) if nums else 1
        row["_max_wrong"] = int(nums[1]) if len(nums)>1 else int(nums[0]) if len(nums)==1 else 3


def _write_choice_excel(questions):
    wb = openpyxl.load_workbook(CHOICE_TEMPLATE); ws = wb.active; sr = ws.max_row
    for qi, q in enumerate(q for q in questions if q["_type"]=="CHOICE"):
        r = sr + qi + 1
        d = {1: q["_stem"], 2: q.get("题目",""), 11: q["_answer"], 18: q.get("知识点",""),
             21: q.get("负反馈",""), 22: q.get("正反馈",""), 27: q.get("_min_wrong",1), 28: q.get("_max_wrong",3)}
        for oi, opt in enumerate(q.get("_options",[])):
            if opt: d[3+oi] = opt
        for c, v in d.items():
            if v: ws.cell(row=r, column=c).value = v
    wb.save(OUTPUT_CHOICE); return OUTPUT_CHOICE


def _write_fill_excel(questions):
    wb = openpyxl.load_workbook(FILL_TEMPLATE); ws = wb.active; sr = ws.max_row
    for qi, q in enumerate(q for q in questions if q["_type"]=="FILL"):
        r = sr + qi + 1; bl = len(re.findall(r'_{2,}', q.get("题目",""))) or 1
        d = {1: "COMPLETION", 2: q.get("位置",""), 3: q.get("知识点",""), 9: q["_stem"][:80], 10: q.get("题目","")}
        if bl >= 1: d.update({12:"空1",13:"内容型",14:"内容",15:"文本输入框",16:q["_answer"],18:"横向双线键盘",19:"是",20:"否"})
        if bl >= 2: d.update({25:"空2",26:"内容型",27:"内容",28:"文本输入框",29:q.get("_answer2",""),31:"横向双线键盘",32:"是",33:"否"})
        for c, v in d.items():
            if v: ws.cell(row=r, column=c).value = v
    wb.save(OUTPUT_FILL); return OUTPUT_FILL


@router.post("/convert")
async def feishu_convert(body: FeishuConvertRequest) -> Dict[str, Any]:
    try:
        raw = _fetch_raw_content(body.url); qs = _parse_questions(raw)
        if not qs: return {"success": False, "error": "未解析到任何题目"}
        files = []
        if any(q["_type"]=="CHOICE" for q in qs): files.append(_write_choice_excel(qs))
        if any(q["_type"]=="FILL" for q in qs): files.append(_write_fill_excel(qs))
        return {"success": True, "summary": {"total": len(qs)}, "files": files}
    except Exception as e:
        return {"success": False, "error": str(e)}


@router.get("/download")
async def feishu_download(path: str = Query(...)):
    fp = Path(path)
    if not fp.exists(): return {"error": "文件不存在"}
    return FileResponse(fp, filename=fp.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
