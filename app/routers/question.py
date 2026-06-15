"""试题批量生成路由。"""

import asyncio
import io
import json
import uuid as uuid_mod
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook
from pydantic import BaseModel

from ..client import MortonClient, MortonClientError
from config import SUBJECT_MAP, GRADE_MAP, STAGE_MAP, DIFFICULTY_MAP

router = APIRouter(prefix="/api/question", tags=["试题"])

QUESTION_ROOT = "/merton-backend/question"
PREVIEW_NAME_PREFIX = "【预览-可删除】"


class ChoiceOption(BaseModel):
    label: str
    content: str


class ConditionFeedback(BaseModel):
    """条件反馈 - 基于错误次数触发"""
    minWrong: int = 1          # 最小错误次数
    maxWrong: int = 999        # 最大错误次数
    content: str = ""          # 反馈文本
    imageUrl: str = ""         # 图片URL
    audioUrl: str = ""         # 音频URL


class QuestionItem(BaseModel):
    stem: str
    contexts: List[str] = []
    options: List[ChoiceOption]
    answer: str
    name: str = ""
    # 学科/分辨率/年级/阶段/难度/知识点
    subjectId: int = 0
    resolution: str = "1334*750"
    grade: int = 0
    stage: int = 0
    difficulty: int = 0
    knowledge: str = ""              # 分号分隔的知识点文本
    # 题目图片 & 背景（画布布局用）
    questionImages: List[dict] = []   # [{url, width, height}, ...]
    backgroundUrl: str = ""
    # 选项图片（按选项索引，如 [{"index":0, "url":"..."}, ...]）
    optionImages: List[dict] = []     # [{index, url}, ...]
    # 反馈相关
    wrongFeedback: str = ""
    correctFeedback: str = ""
    wrongFeedbackImg: str = ""
    correctFeedbackImg: str = ""
    wrongFeedbackAudio: str = ""
    correctFeedbackAudio: str = ""
    conditionFeedbacks: List[ConditionFeedback] = []   # 条件反馈列表（基于错误次数）


class BatchRequest(BaseModel):
    questions: List[QuestionItem]
    subjectId: int = 0
    cookie: str
    token: str
    defaultName: str = "新建选择题"
    # 新增统一默认值
    defaultResolution: str = "1334*750"
    defaultGrade: int = 0
    defaultStage: int = 0
    defaultDifficulty: int = 0
    defaultKnowledge: str = ""
    # 反馈默认值
    defaultWrongFeedback: str = ""
    defaultCorrectFeedback: str = ""
    defaultWrongImg: str = ""
    defaultCorrectImg: str = ""
    defaultWrongAudio: str = ""
    defaultCorrectAudio: str = ""
    defaultQuestionImages: str = ""
    defaultBackgroundUrl: str = ""
    defaultConditionFeedback: Optional[ConditionFeedback] = None


class PreviewCreateRequest(BaseModel):
    question: QuestionItem
    subjectId: int = 0
    cookie: str
    token: str
    defaultName: str = "新建选择题"
    defaultResolution: str = "1334*750"
    defaultGrade: int = 0
    defaultStage: int = 0
    defaultDifficulty: int = 0
    defaultKnowledge: str = ""
    defaultWrongFeedback: str = ""
    defaultCorrectFeedback: str = ""
    defaultWrongImg: str = ""
    defaultCorrectImg: str = ""
    defaultWrongAudio: str = ""
    defaultCorrectAudio: str = ""
    defaultQuestionImages: str = ""
    defaultBackgroundUrl: str = ""
    defaultConditionFeedback: Optional[ConditionFeedback] = None


class PreviewPayloadRequest(BaseModel):
    """创建前本地预览请求：只构建最终 questionData，不调用莫顿 insert/update。"""
    question: QuestionItem
    subjectId: int = 0
    defaultName: str = "新建选择题"
    defaultResolution: str = "1334*750"
    defaultGrade: int = 0
    defaultStage: int = 0
    defaultDifficulty: int = 0
    defaultKnowledge: str = ""
    defaultWrongFeedback: str = ""
    defaultCorrectFeedback: str = ""
    defaultWrongImg: str = ""
    defaultCorrectImg: str = ""
    defaultWrongAudio: str = ""
    defaultCorrectAudio: str = ""
    defaultQuestionImages: str = ""
    defaultBackgroundUrl: str = ""
    defaultConditionFeedback: Optional[ConditionFeedback] = None


def _apply_defaults(item: QuestionItem, body) -> None:
    """应用全局默认值，确保正式生成和真实预览使用同一套字段回填规则。"""
    if not item.name or item.name.strip() == "":
        item.name = body.defaultName
    if not item.resolution or item.resolution.strip() == "":
        item.resolution = body.defaultResolution
    if item.grade == 0 and body.defaultGrade != 0:
        item.grade = body.defaultGrade
    if item.stage == 0 and body.defaultStage != 0:
        item.stage = body.defaultStage
    if item.difficulty == 0 and body.defaultDifficulty != 0:
        item.difficulty = body.defaultDifficulty
    if not item.knowledge or item.knowledge.strip() == "":
        item.knowledge = body.defaultKnowledge
    if item.subjectId == 0 and body.subjectId != 0:
        item.subjectId = body.subjectId
    if (not item.wrongFeedback or item.wrongFeedback.strip() == "") and body.defaultWrongFeedback:
        item.wrongFeedback = body.defaultWrongFeedback
    if (not item.correctFeedback or item.correctFeedback.strip() == "") and body.defaultCorrectFeedback:
        item.correctFeedback = body.defaultCorrectFeedback
    if (not item.wrongFeedbackImg or item.wrongFeedbackImg.strip() == "") and body.defaultWrongImg:
        item.wrongFeedbackImg = body.defaultWrongImg
    if (not item.correctFeedbackImg or item.correctFeedbackImg.strip() == "") and body.defaultCorrectImg:
        item.correctFeedbackImg = body.defaultCorrectImg
    if (not item.wrongFeedbackAudio or item.wrongFeedbackAudio.strip() == "") and body.defaultWrongAudio:
        item.wrongFeedbackAudio = body.defaultWrongAudio
    if (not item.correctFeedbackAudio or item.correctFeedbackAudio.strip() == "") and body.defaultCorrectAudio:
        item.correctFeedbackAudio = body.defaultCorrectAudio

    # 题目图片全局默认（分号分隔，仅当 Excel 未提供时生效）
    if (not item.questionImages or len(item.questionImages) == 0) and body.defaultQuestionImages:
        images = []
        for url in body.defaultQuestionImages.split(";"):
            url = url.strip()
            if url:
                images.append({"url": url, "width": 0, "height": 0})
        if images:
            item.questionImages = images

    # 画布背景全局默认
    if (not item.backgroundUrl or item.backgroundUrl.strip() == "") and body.defaultBackgroundUrl:
        item.backgroundUrl = body.defaultBackgroundUrl

    # 条件反馈全局默认（仅当逐题未设置时生效）
    if (not item.conditionFeedbacks or len(item.conditionFeedbacks) == 0) and body.defaultConditionFeedback:
        cf = body.defaultConditionFeedback
        # 只有填写了反馈文本时才添加
        if cf.content and cf.content.strip():
            item.conditionFeedbacks = [cf]


# ══════════════════════════════════════════════════════════════════
#  智能画布布局引擎 — 4 阶段核心逻辑
# ══════════════════════════════════════════════════════════════════
#  绿框内容区: x=182~1152（宽970）, y=130~580（高450）
#  布局对象: 题目文本 | 题目图片 | 选项框 | 背景图
#
#  阶段 1 — MEASURE:  精确测量每个元素的实际像素尺寸
#  阶段 2 — PARTITION: 根据内容组合划分绿框空间
#  阶段 3 — PLACE:    在各自分区内放置元素
#  阶段 4 — ADJUST:   微调间距 / 避免越界 / 保证美感
# ══════════════════════════════════════════════════════════════════

# 字符宽度系数（以字号为基准）
_CW_CN = 0.85    # 中文字符
_CW_EN = 0.55    # 英文/数字
_CW_PUNCT = 0.40  # 标点/空格
_LH_RATIO = 1.35  # 行高系数


def _char_width(ch: str, font_size: int) -> float:
    """精确测量单个字符的渲染宽度（CJK 字体环境）。

    规则：ASCII 范围用窄系数，其余一律按中文字符宽度处理。
    这样 ∠、△、∥、＝ 等数学符号不会因块区遗漏而被低估。
    """
    code = ord(ch)
    # ASCII 窄字符
    if code < 0x80:
        # 英文标点最窄
        if ch in ',.!?;:()[]{}':
            return font_size * _CW_PUNCT
        # 英文/数字
        return font_size * _CW_EN
    # 中文标点（U+3000 区块 + 全角标点区块：逗号、顿号、引号等）
    lo = 0x3000 <= code <= 0x303F or 0xFF01 <= code <= 0xFF5E
    # 特殊窄标点：省略号、弯引号、中点、中文空格
    special = code in (0x2026, 0x2018, 0x2019, 0x201C, 0x201D, 0x00B7, 0x3000)
    if lo or special:
        return font_size * _CW_PUNCT
    # 其余非 ASCII 字符一律按中文字符宽度
    return font_size * _CW_CN


def _wrap_text_to_lines(text: str, font_size: int, max_width: int) -> list:
    """将文本按 max_width 智能换行，返回每行文本列表。

    核心规则：
    1. 逐字累加宽度，超 max_width 时换行
    2. 优先在标点/空格处断行（向前找最近标点，最多回溯 5 字）
    3. 中文标点不允许作为行首（挂在上一行末尾）
    4. 显式 \\n 强制换行
    """
    # 断行候选取值 — 在这些字符后断行最自然
    BREAK_CHARS = set('，。！？；：、\n\r\t ')

    if not text:
        return []

    all_lines = []
    for paragraph in text.split('\n'):
        if not paragraph:
            all_lines.append('')
            continue

        current_line = ''
        current_width = 0.0

        for i, ch in enumerate(paragraph):
            cw = _char_width(ch, font_size)
            # 检查是否需要换行
            if current_width + cw > max_width and current_line:
                # 向前找回溯：优先在标点后断行
                best_break = len(current_line)  # 默认在当前位置断
                lookback = min(10, len(current_line))
                for j in range(len(current_line) - 1, max(len(current_line) - lookback - 1, 0), -1):
                    if current_line[j] in BREAK_CHARS:
                        best_break = j + 1  # 在标点之后断
                        break

                if best_break < len(current_line):
                    # 把断点后的字符移到下一行
                    carry = current_line[best_break:]
                    all_lines.append(current_line[:best_break])
                    current_line = carry
                else:
                    all_lines.append(current_line)
                    current_line = ''

                current_width = sum(_char_width(c, font_size) for c in current_line)

                # 行首不能是中文标点
                if current_line and current_line[0]:
                    fc = ord(current_line[0])
                    if 0x3000 <= fc <= 0x303F or 0xFF00 <= fc <= 0xFFEF:
                        # 把标点挂到上一行末尾
                        if all_lines:
                            all_lines[-1] += current_line[0]
                            current_line = current_line[1:]
                            current_width = sum(_char_width(c, font_size) for c in current_line)

            current_line += ch
            current_width += cw

        if current_line:
            all_lines.append(current_line)

    return all_lines


def _measure_paragraph(text: str, font_size: int, max_width: int) -> dict:
    """测量一段文本：返回换行后的行列表、总高度、最大行宽。"""
    lines = _wrap_text_to_lines(text, font_size, max_width)
    result = {
        "lines": lines,
        "line_count": len(lines),
        "max_line_width": 0,
        "total_height": 0,
    }
    if not lines:
        return result

    max_lw = 0
    for line in lines:
        lw = sum(_char_width(c, font_size) for c in line)
        max_lw = max(max_lw, lw)

    line_height = font_size * _LH_RATIO
    result["max_line_width"] = round(max_lw)
    result["total_height"] = round(len(lines) * line_height)
    return result


def _measure_image(img: dict, max_w: int, max_h: int) -> dict:
    """缩放图片以适配最大尺寸，保持宽高比。"""
    iw = img.get("width", 200)
    ih = img.get("height", 200)
    if iw <= 0 or ih <= 0:
        iw, ih = 200, 200
    scale = min(max_w / iw, max_h / ih, 1.0)
    return {
        "url": img.get("url", ""),
        "width": round(iw * scale),
        "height": round(ih * scale),
    }


def _calculate_smart_layout(
    options_count: int,
    contexts_count: int,
    context_texts: List[str] = None,
    question_images: List[dict] = None,
    background_url: str = None,
) -> dict:
    """智能画布布局引擎 — 4 阶段核心逻辑。

    参数:
        options_count: 选项数量 (2~4)
        contexts_count: 文本块数量
        context_texts: 文本块内容列表
        question_images: 题目图片 [{url, width, height}, ...]
        background_url: 画布背景图 URL

    返回:
        {
            "option_positions": [{x,y,width,height}, ...],
            "text_positions": [{x,y,width,align}, ...],
            "image_positions": [{x,y,width,height,url}, ...],
            "background_url": str,
            "font_size": int,
        }
    """
    # ── 内容区域常量 ──
    CL = 182          # Content Left
    CR = 1152         # Content Right
    CT = 130          # Content Top
    CB = 580          # Content Bottom
    CW = CR - CL      # 970
    CH = CB - CT      # 450
    CCX = (CL + CR) // 2  # 667

    # ── 边距常量 ──
    PAD_TOP = 20
    PAD_BOTTOM = 20
    MIN_GAP = 20      # 元素间最小间距

    # ══════════════════════════════════════════════
    #  阶段 1: MEASURE — 精确测量每个元素尺寸
    # ══════════════════════════════════════════════

    # 1a. 确定最优字号
    total_chars = 0
    if context_texts:
        total_chars = sum(len(t) for t in context_texts)

    if total_chars <= 15:
        FONT_SIZE = 50
    elif total_chars <= 40:
        FONT_SIZE = 45
    else:
        FONT_SIZE = 40

    # 1b. 测量文本：逐字换行
    text_available_w = CW  # 默认全宽
    text_paragraphs = []
    if context_texts and contexts_count > 0:
        for text in context_texts:
            m = _measure_paragraph(text, FONT_SIZE, text_available_w)
            text_paragraphs.append(m)

    # 1c. 测量图片
    images = []
    if question_images:
        img_max_w = int(CW * 0.45)   # 图片最大宽度不超过绿框 45%
        img_max_h = int(CH * 0.35)   # 图片最大高度不超过绿框 35%
        for img in question_images:
            images.append(_measure_image(img, img_max_w, img_max_h))

    # ══════════════════════════════════════════════
    #  阶段 2: PARTITION — 划分绿框空间
    # ══════════════════════════════════════════════

    # 2a. 计算文本总高度
    text_total_h = 0
    text_max_w = 0
    for m in text_paragraphs:
        text_total_h += m["total_height"]
        text_max_w = max(text_max_w, m["max_line_width"])
    if len(text_paragraphs) > 1:
        text_total_h += (len(text_paragraphs) - 1) * 16  # 段间距

    # 2b. 计算图片占用高度
    img_total_h = 0
    img_total_w = 0
    if images:
        # 图片横向排列，间距 MIN_GAP
        img_total_w = sum(i["width"] for i in images) + (len(images) - 1) * MIN_GAP
        img_total_h = max(i["height"] for i in images)

    # 2c. 空间分配
    # 决定是否有图片区
    has_images = len(images) > 0

    if has_images:
        # ── 有图片的布局 ──
        # 图片和文本可以并排（图片在左/右，文本在另一边）
        # 或者上下排列（图片在上，文本在下）
        img_text_gap = MIN_GAP

        # 判断模式：图片总宽 + 文本最小宽 能否并排
        text_min_w = min(text_max_w, CW * 0.35) if text_max_w > 0 else 0
        can_inline = (img_total_w + text_min_w + img_text_gap * 2) <= CW

        if can_inline and len(images) == 1:
            # 模式 A: 图文并排 — 文本左侧，图片右侧（莫顿几何题标准布局）
            layout_mode = "TEXT_LEFT_IMAGE_RIGHT"
            text_zone_w = CW - images[0]["width"] - img_text_gap * 2
            text_top_area_h = max(text_total_h, images[0]["height"])
            # 重新测量文本（用缩小的宽度）
            text_paragraphs = []
            for text in (context_texts or []):
                text_paragraphs.append(_measure_paragraph(text, FONT_SIZE, text_zone_w))
            text_total_h = sum(m["total_height"] for m in text_paragraphs)
            if len(text_paragraphs) > 1:
                text_total_h += (len(text_paragraphs) - 1) * 16
            text_max_w = max((m["max_line_width"] for m in text_paragraphs), default=0)
            text_top_area_h = max(text_total_h, images[0]["height"])
        else:
            # 模式 B: 图文上下 — 图片在上，文本居中在下
            layout_mode = "IMAGE_TOP_TEXT_BELOW"
            text_zone_w = CW
            text_top_area_h = text_total_h
    else:
        # ── 纯文本布局 ──
        layout_mode = "TEXT_ONLY"
        text_zone_w = CW
        text_top_area_h = text_total_h
        img_total_h = 0

    # 2d. 选项可用空间 — 无内容时自动垂直居中
    upper_content_h = (img_total_h + MIN_GAP if has_images else 0) + text_top_area_h
    option_available_h = CH - PAD_TOP - upper_content_h - MIN_GAP - PAD_BOTTOM
    OPTION_H = max(100, min(130, option_available_h))

    if upper_content_h == 0:
        option_y_base = CT + PAD_TOP + (CH - PAD_TOP - PAD_BOTTOM - OPTION_H) // 2
    else:
        option_y_base = CB - PAD_BOTTOM - OPTION_H

    # ══════════════════════════════════════════════
    #  阶段 3: PLACE — 在分区内放置元素
    # ══════════════════════════════════════════════

    # 3a. 放置图片
    image_positions = []
    if has_images:
        if layout_mode == "TEXT_LEFT_IMAGE_RIGHT":
            # 图片放右侧，与文本顶部对齐
            img_x = CR - MIN_GAP - images[0]["width"]
            img_y = CT + PAD_TOP
            image_positions.append({
                "x": img_x,
                "y": img_y,
                "width": images[0]["width"],
                "height": images[0]["height"],
                "url": images[0]["url"],
            })
        elif layout_mode == "IMAGE_TOP_TEXT_BELOW":
            # 图片顶部居中排列
            img_start_x = CCX - img_total_w // 2
            cx = img_start_x
            img_y = CT + PAD_TOP
            for im in images:
                image_positions.append({
                    "x": cx,
                    "y": img_y,
                    "width": im["width"],
                    "height": im["height"],
                    "url": im["url"],
                })
                cx += im["width"] + MIN_GAP

    # 3b. 放置文本
    text_positions = []
    if context_texts and contexts_count > 0:
        text_start_y = CT + PAD_TOP

        if layout_mode == "TEXT_LEFT_IMAGE_RIGHT":
            # 文本从左侧开始，占据图片左边的剩余宽度
            text_zone_x = CL + MIN_GAP
            text_zone_w_actual = CW - images[0]["width"] - MIN_GAP * 2
            current_y = text_start_y
        elif layout_mode == "IMAGE_TOP_TEXT_BELOW":
            # 文本在图片下方，加边距
            text_zone_x = CL + MIN_GAP
            text_zone_w_actual = CW - MIN_GAP * 2
            current_y = CT + PAD_TOP + img_total_h + MIN_GAP
        else:
            # 纯文本：文本从绿框左侧+边距开始，x=202（匹配莫顿实测）
            text_zone_x = CL + MIN_GAP
            text_zone_w_actual = CW - MIN_GAP * 2
            current_y = text_start_y

        for m in text_paragraphs:
            text_positions.append({
                "x": round(text_zone_x),                  # 向后兼容
                "zone_x": round(text_zone_x),             # 文本 zone 左边缘
                "y": round(current_y),                     # 文本顶部 y
                "width": text_zone_w_actual,               # zone 全宽（displayBoxWeight）
                "align": "center",                         # 莫顿内部居中
            })
            current_y += m["total_height"] + 16

    # 3c. 放置选项 — 无内容时垂直居中
    if upper_content_h == 0:
        option_y_base = CT + PAD_TOP + (CH - PAD_TOP - PAD_BOTTOM - OPTION_H) // 2
    else:
        option_y_base = CB - PAD_BOTTOM - OPTION_H

    # 动态宽度与间距
    if options_count == 2:
        opt_w, gap = 220, max(180, (CW - 2 * 220) // 2)
    elif options_count == 3:
        opt_w, gap = 210, max(80, (CW - 3 * 210) // 3)
    elif options_count == 4:
        opt_w, gap = 190, 50
    else:
        opt_w, gap = 160, 40

    # 按比例缩放适配绿框宽度
    total_opt_w = options_count * opt_w + (options_count - 1) * gap
    if total_opt_w > CW:
        scale = CW / total_opt_w
        opt_w = max(140, int(opt_w * scale))
        gap = max(30, int(gap * scale))

    opt_start_x = CL + (CW - (options_count * opt_w + (options_count - 1) * gap)) // 2

    option_positions = []
    for i in range(options_count):
        option_positions.append({
            "x": round(opt_start_x + i * (opt_w + gap)),
            "y": round(option_y_base),
            "width": round(opt_w),
            "height": round(OPTION_H),
        })

    # ══════════════════════════════════════════════
    #  阶段 4: ADJUST — 微调防止越界
    # ══════════════════════════════════════════════

    for tp in text_positions:
        if tp["x"] < CL:
            tp["x"] = CL + 10
        if tp["x"] + tp["width"] > CR:
            tp["width"] = CR - tp["x"] - 10

    for op in option_positions:
        op["x"] = max(CL + 5, min(op["x"], CR - op["width"] - 5))

    return {
        "option_positions": option_positions,
        "text_positions": text_positions,
        "image_positions": image_positions,
        "background_url": background_url or "",
        "font_size": FONT_SIZE,
        "layout_mode": layout_mode,
    }


def _merge_update_fields(question_data: Dict[str, Any], item) -> tuple[bool, bool]:
    """把 /update 阶段保存的元数据和反馈字段合并进 questionData。"""
    has_update = False

    # ===== 元数据部分 =====
    if item.subjectId:
        question_data["subjectId"] = item.subjectId
        has_update = True
    if item.difficulty:
        question_data["difficulty"] = item.difficulty
        has_update = True
    if item.grade:
        question_data["grade"] = item.grade
        has_update = True
    if item.stage:
        question_data["stage"] = item.stage
        has_update = True

    # 画布背景 — 莫顿字段: webBgmUrl（BGM = background）
    if item.backgroundUrl:
        question_data["webBgmUrl"] = item.backgroundUrl
        has_update = True
        print(f"[UPDATE] 设置画布背景: {item.backgroundUrl[:80]}...")

    # ===== 反馈部分 =====
    wrong = item.wrongFeedback
    correct = item.correctFeedback
    wrong_img = item.wrongFeedbackImg
    correct_img = item.correctFeedbackImg
    wrong_audio = item.wrongFeedbackAudio
    correct_audio = item.correctFeedbackAudio

    has_feedback = bool(wrong or correct or wrong_img or correct_img or wrong_audio or correct_audio or item.conditionFeedbacks)

    if wrong or wrong_img or wrong_audio:
        fb = {"answerType": "PICTURE_CONTENT"}
        if wrong: fb["content"] = wrong
        if wrong_img: fb["webUrl"] = wrong_img; fb["webName"] = wrong_img.rsplit("/",1)[-1] if "/" in wrong_img else wrong_img
        if wrong_audio: fb["webAudioUrl"] = wrong_audio
        question_data["wrongQuestionFeedback"] = fb
        has_update = True

    if correct or correct_img or correct_audio:
        fb = {"answerType": "PICTURE_CONTENT"}
        if correct: fb["content"] = correct
        if correct_img: fb["webUrl"] = correct_img; fb["webName"] = correct_img.rsplit("/",1)[-1] if "/" in correct_img else correct_img
        if correct_audio: fb["webAudioUrl"] = correct_audio
        question_data["rightQuestionFeedback"] = fb
        has_update = True

    # 条件反馈（基于错误次数）
    if item.conditionFeedbacks:
        cfbs = []
        for cf in item.conditionFeedbacks:
            fb = {
                "uuid": str(uuid_mod.uuid4()),
                "answerType": "PICTURE_CONTENT",
                "webListenList": [{
                    "uuid": str(uuid_mod.uuid4()),
                    "variableKey": "WrongNumber",
                    "variableType": "int",
                }],
                "logicExpressList": [],
                "logicTypeList": [],
            }
            if cf.content: fb["content"] = cf.content
            if cf.imageUrl:
                fb["webUrl"] = cf.imageUrl
                fb["webName"] = cf.imageUrl.rsplit("/",1)[-1] if "/" in cf.imageUrl else cf.imageUrl
            if cf.audioUrl: fb["webAudioUrl"] = cf.audioUrl

            # 区间逻辑
            listen = fb["webListenList"][0]
            if cf.minWrong == cf.maxWrong:
                listen["compare"] = "EQUAL"
                listen["value"] = cf.minWrong
            else:
                listen["compare"] = "BETWEEN"
                listen["min"] = cf.minWrong
                listen["max"] = cf.maxWrong

            cfbs.append(fb)
        question_data["questionFeedbacks"] = cfbs
        has_update = True

    return has_update, has_feedback


async def _update_question(client, question_data: Dict[str, Any], question_id, item) -> str:
    """通过单次 /update 写入学科、年级、阶段、难度 + 反馈（合并两次调用为一次）。"""
    update_data = dict(question_data)
    update_data["questionId"] = question_id
    has_update, has_feedback = _merge_update_fields(update_data, item)

    if not has_update:
        return ""

    try:
        print(f"[UPDATE] 第{question_id}题: subjectId={item.subjectId} grade={item.grade} stage={item.stage} difficulty={item.difficulty} hasFeedback={has_feedback}")
        result = await client.post(
            f"{QUESTION_ROOT}/update",
            json={"questionData": update_data, "id": question_id},
        )
        if result.get("code") != 200:
            print(f"[UPDATE] 失败: {result}")
            return f"更新失败: {result.get('message', '未知')}"
        print(f"[UPDATE] 成功")
        return ""
    except Exception as e:
        return f"更新异常: {e}"


def _build_insert_payload(item: QuestionItem, subject_id: int) -> Dict[str, Any]:
    """构建 insert 接口的完整 payload。"""
    FONT_SIZE = 32
    RGB = "#9f765b"

    # 根据分辨率动态设置画布尺寸
    res = item.resolution or "1334*750"
    if "*" in res:
        parts = res.replace("高清", "").strip().split("*")
        CANVAS_WIDTH = int(parts[0]) if parts[0].isdigit() else 1334
        CANVAS_HEIGHT = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 750
    else:
        CANVAS_WIDTH, CANVAS_HEIGHT = 1334, 750

    # 智能布局引擎：4 阶段核心逻辑 — 测量→分区→放置→微调
    layout = _calculate_smart_layout(
        options_count=len(item.options),
        contexts_count=len(item.contexts),
        context_texts=item.contexts if item.contexts else None,
        question_images=item.questionImages if item.questionImages else None,
        background_url=item.backgroundUrl if item.backgroundUrl else None,
    )
    opt_positions = layout["option_positions"]
    txt_positions = layout["text_positions"]
    text_font_size = layout["font_size"]
    image_positions = layout["image_positions"]
    background_url = layout["background_url"]

    # 构建选项
    # 选项图片映射表 {index: url}
    opt_img_map = {}
    for oi in (item.optionImages or []):
        idx = oi.get("index", -1)
        url = oi.get("url", "")
        if idx >= 0 and url:
            opt_img_map[idx] = url

    options: List[Dict[str, Any]] = []
    for i, opt in enumerate(item.options):
        pos = opt_positions[i]
        is_answer = opt.label.upper() == item.answer.upper()
        opt_payload = {
            "uuid": str(uuid_mod.uuid4()),
            "answer": is_answer,
            "content": opt.content,
            "fontSize": FONT_SIZE,
            "rgb": RGB,
            "displayBoxWeight": pos["width"],
            "displayBoxHeight": pos["height"],
            "coordinate": {"x": pos["x"], "y": pos["y"]},
            "isAppear": True,
            "scaleValue": 100,
            "tmpStr": f"选项 {opt.label.upper()}",
            "optionIndex": opt.label.upper(),
            "isDrag": False,
        }
        # 选项图片
        if i in opt_img_map:
            opt_payload["imageUrl"] = opt_img_map[i]
        options.append(opt_payload)

    # 画布文本 contexts — x 固定在绿框左边缘，displayBoxWeight 占满可用宽度
    contexts: List[Dict[str, Any]] = []
    for i, text_block in enumerate(item.contexts):
        pos = txt_positions[i]
        # 文本 zone 可用宽度（layout 引擎已根据不同模式计算好）
        zone_width = pos.get("width", 970)
        # 按 zone 宽度换行，插入 \n
        lines = _wrap_text_to_lines(text_block, text_font_size, zone_width)
        content_with_newlines = "\n".join(lines)
        print(f"[CONTEXT] zone_w={zone_width} x={pos.get('zone_x', pos.get('x', '?'))} "
              f"lines={len(lines)} has_nl={'\\n' in content_with_newlines} "
              f"content_preview={repr(content_with_newlines[:60])}")
        contexts.append({
            "content": content_with_newlines,
            "size": text_font_size,
            "rgb": RGB,
            "coordinate": {"x": pos.get("zone_x", pos["x"]), "y": pos["y"]},
            "uuid": str(uuid_mod.uuid4()),
        })

    question_data: Dict[str, Any] = {
        "answerFeedbackType": "ALL",
        "skinVersion": 1,
        "resolution": res.replace("高清", "").strip(),
        "questionItem": {
            "displayNumber": False,
            "displayBox": True,
            "newDisplayBox": True,
            "options": options,
        },
        "questionType": "CHOICE",
        "autoWidth": CANVAS_WIDTH,
        "autoHeight": CANVAS_HEIGHT,
        "name": item.name,
        "stem": {"content": item.stem},
        "subjectId": item.subjectId or subject_id,
    }

    # 知识点：纯字符串数组格式 ["数学", "几何"]
    if item.knowledge:
        question_data["knowledge"] = [k.strip() for k in item.knowledge.split(";") if k.strip()]

    # 年级/阶段/难度（从莫顿 API 实测确认字段名）
    if item.difficulty:
        question_data["difficulty"] = item.difficulty
    if item.grade:
        question_data["grade"] = item.grade
    if item.stage:
        question_data["stage"] = item.stage

    # 打印关键字段供验证
    print(f"[INSERT] name={item.name} subjectId={question_data['subjectId']} "
          f"resolution={question_data['resolution']} "
          f"knowledge={question_data.get('knowledge','')} "
          f"grade={item.grade} stage={item.stage} difficulty={item.difficulty}")

    if contexts:
        question_data["contexts"] = contexts

    # 画布图片 — 莫顿字段: pictures[].{uuid,webPictureUrl,coordinate,scaleValue,isDrag}
    if image_positions:
        pics: List[Dict[str, Any]] = []
        for ip in image_positions:
            pics.append({
                "uuid": str(uuid_mod.uuid4()),
                "webPictureUrl": ip["url"],
                "coordinate": {"x": ip["x"], "y": ip["y"]},
                "scaleValue": 200,   # 默认缩放 200%（匹配莫顿手动添加行为）
                "isDrag": True,
            })
        question_data["pictures"] = pics

    # 画布背景 — 莫顿字段: webBgmUrl（insert 不支持，统一在 update 阶段写入）
    # 此处不设置 question_data["webUrl"]，背景在 _merge_update_fields 中处理

    return {"questionData": question_data}, question_data


@router.post("/preview-payload")
async def preview_payload(body: PreviewPayloadRequest) -> Dict[str, Any]:
    """构建创建前预览用的最终 questionData；不会在莫顿系统创建题目。"""
    try:
        item = body.question
        _apply_defaults(item, body)
        _, question_data = _build_insert_payload(item, body.subjectId)
        _merge_update_fields(question_data, item)
        return {
            "success": True,
            "questionData": question_data,
            "feedbackError": None,
        }
    except Exception as e:
        return {"success": False, "error": f"预览数据构建异常: {e}"}


@router.post("/preview-create")
async def preview_create(body: PreviewCreateRequest) -> Dict[str, Any]:
    """创建一条莫顿预览题，用真实 widget 截图还原解析后的题目效果。"""
    client = MortonClient(cookie=body.cookie, token=body.token)
    try:
        item = body.question
        _apply_defaults(item, body)

        original_name = item.name or body.defaultName or "新建选择题"
        if not original_name.startswith(PREVIEW_NAME_PREFIX):
            item.name = f"{PREVIEW_NAME_PREFIX}{original_name}"

        payload, question_data = _build_insert_payload(item, body.subjectId)
        # DEBUG: 打印发送给莫顿的选项坐标
        qd = payload.get("questionData", {})
        qi = qd.get("questionItem", {})
        for o in qi.get("options", []):
            coord = o.get("coordinate", {})
            print(f"[INSERT_DEBUG] 选项{o.get('optionIndex','?')}: x={coord.get('x')}, y={coord.get('y')}, w={o.get('displayBoxWeight')}, h={o.get('displayBoxHeight')}")
        result = await client.post(f"{QUESTION_ROOT}/insert", json=payload)

        if result.get("code") != 200:
            print(f"[PREVIEW] 插入失败: {result}")
            return {
                "success": False,
                "error": f"预览题创建失败: {result.get('message', '未知错误')}",
            }

        question_id = result.get("data")
        update_error = await _update_question(client, question_data, question_id, item)

        return {
            "success": True,
            "previewQuestionId": question_id,
            "feedbackError": update_error if update_error else None,
            "warning": "已在莫顿系统创建一条带【预览-可删除】前缀的预览题；当前未接入删除接口，请后续在莫顿后台手动清理。",
        }

    except MortonClientError as e:
        return {"success": False, "error": e.message}
    except Exception as e:
        return {"success": False, "error": f"预览题创建异常: {e}"}
    finally:
        await client.close()


@router.post("/batch-create")
async def batch_create(body: BatchRequest) -> Dict[str, Any]:
    """批量创建选择题。"""
    client = MortonClient(cookie=body.cookie, token=body.token)
    results: List[Dict[str, Any]] = []

    # 应用默认值，确保正式生成和真实预览一致
    for q in body.questions:
        _apply_defaults(q, body)

    try:
        for idx, item in enumerate(body.questions):
            try:
                payload, question_data = _build_insert_payload(item, body.subjectId)
                result = await client.post(f"{QUESTION_ROOT}/insert", json=payload)

                if result.get("code") != 200:
                    # 打印完整响应便于排查
                    print(f"[ERROR] 第{idx+1}题插入失败: {result}")
                    results.append({
                        "index": idx,
                        "success": False,
                        "error": f"创建失败: {result.get('message', '未知错误')}",
                    })
                    continue

                question_id = result.get("data")

                # 通过单次 /update 确保元数据+反馈被正确保存
                update_error = await _update_question(
                    client, question_data, question_id, item
                )

                results.append({
                    "index": idx,
                    "success": True,
                    "questionId": question_id,
                    "feedbackError": update_error if update_error else None,
                })

            except MortonClientError as e:
                results.append({
                    "index": idx,
                    "success": False,
                    "error": e.message,
                })

            if idx < len(body.questions) - 1:
                await asyncio.sleep(0.3)

        success_count = sum(1 for r in results if r["success"])
        return {
            "success": True,
            "total": len(body.questions),
            "successCount": success_count,
            "failCount": len(body.questions) - success_count,
            "results": results,
        }
    except MortonClientError as e:
        return {"success": False, "error": e.message}
    finally:
        await client.close()


@router.post("/batch-create-stream")
async def batch_create_stream(body: BatchRequest):
    """SSE 流式批量创建选择题，实时推送进度。"""

    async def event_generator():
        client = MortonClient(cookie=body.cookie, token=body.token)
        total = len(body.questions)

        # 发送初始事件，确保连接建立
        yield f"data: {json.dumps({'total': total, 'progress': f'0/{total}', 'percent': 0}, ensure_ascii=False)}\n\n"

        # 应用默认值（与 batch_create / 预览创建相同逻辑）
        for q in body.questions:
            _apply_defaults(q, body)

        success_count = 0
        fail_count = 0

        try:
            for idx, item in enumerate(body.questions):
                try:
                    payload, question_data = _build_insert_payload(item, body.subjectId)
                    # DEBUG: 打印选项坐标
                    qd = payload.get("questionData", {})
                    qi = qd.get("questionItem", {})
                    for o in qi.get("options", []):
                        coord = o.get("coordinate", {})
                        print(f"[STREAM_DEBUG] 选项{o.get('optionIndex','?')}: x={coord.get('x')}, y={coord.get('y')}, w={o.get('displayBoxWeight')}, h={o.get('displayBoxHeight')}")
                    # DEBUG: 看发送给莫顿的 content 里换行符是什么格式
                    ctxs = payload.get("questionData", {}).get("contexts", [])
                    if ctxs:
                        raw_content = ctxs[0].get("content", "")
                        print(f"[PAYLOAD] content_repr={repr(raw_content)} "
                              f"has_real_nl={bool(raw_content)} "
                              f"displayBoxWeight={ctxs[0].get('displayBoxWeight')} "
                              f"x={ctxs[0].get('coordinate', {}).get('x')}")
                    result = await client.post(f"{QUESTION_ROOT}/insert", json=payload)
                    percent = round((idx + 1) / total * 100)
                    progress_str = f"{idx+1}/{total}"

                    if result.get("code") != 200:
                        fail_count += 1
                        print(f"[ERROR] 第{idx+1}题插入失败: {result}")
                        error_msg = result.get("message", "未知错误")
                        evt = {"index": idx, "success": False, "error": f"创建失败: {error_msg}", "progress": progress_str, "percent": percent}
                        yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
                        continue

                    question_id = result.get("data")
                    update_error = await _update_question(client, question_data, question_id, item)
                    success_count += 1
                    evt = {"index": idx, "success": True, "questionId": question_id, "feedbackError": update_error or None, "progress": progress_str, "percent": percent}
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"

                except MortonClientError as e:
                    fail_count += 1
                    percent = round((idx + 1) / total * 100)
                    evt = {"index": idx, "success": False, "error": e.message, "progress": f"{idx+1}/{total}", "percent": percent}
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"

                if idx < total - 1:
                    await asyncio.sleep(0.3)

            # 完成事件
            evt = {"done": True, "total": total, "successCount": success_count, "failCount": fail_count}
            yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"

        except MortonClientError as e:
            evt = {"done": True, "total": total, "successCount": success_count, "failCount": fail_count, "fatalError": e.message}
            yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
        finally:
            await client.close()

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@router.post("/parse-excel")
async def parse_excel(file: UploadFile = File(...)) -> Dict[str, Any]:
    """解析上传的 Excel 文件，返回解析后的题目数据。"""
    try:
        content = await file.read()
        wb = load_workbook(io.BytesIO(content), read_only=True)
        ws = wb.active

        # 读取表头，检测格式版本
        all_rows = list(ws.iter_rows(min_row=1, values_only=True))
        header = all_rows[0] if all_rows else []
        # 新格式：表头 col 6/7 是 "选项E"/"选项F" 或 "E"/"F" → 31 列，选项 A-H
        h6 = str(header[6] or "").strip().upper() if len(header) > 6 else ""
        h7 = str(header[7] or "").strip().upper() if len(header) > 7 else ""
        is_new_format = (len(header) >= 10 and 
                         (h6 in ("E", "选项E")) and
                         (h7 in ("F", "选项F")))
        print(f"[PARSE] 检测格式: {'新(E-H,31列)' if is_new_format else '旧(A-D,22列)'} header={len(header)}列 h6={repr(h6)} h7={repr(h7)}")

        rows = list(ws.iter_rows(min_row=2, values_only=True))
        questions: List[Dict[str, Any]] = []

        for i, row in enumerate(rows):
            if not row or not row[0]:
                continue

            stem = str(row[0] or "").strip()
            contexts_text = str(row[1] or "").strip()
            # 选项 & 答案 — 根据格式版本读取不同列号
            if is_new_format:
                option_a = str(row[2] or "").strip()
                option_b = str(row[3] or "").strip()
                option_c = str(row[4] or "").strip() if len(row) > 4 else ""
                option_d = str(row[5] or "").strip() if len(row) > 5 else ""
                option_e = str(row[6] or "").strip() if len(row) > 6 else ""
                option_f = str(row[7] or "").strip() if len(row) > 7 else ""
                option_g = str(row[8] or "").strip() if len(row) > 8 else ""
                option_h = str(row[9] or "").strip() if len(row) > 9 else ""
                answer = str(row[10] or "").strip().upper() if len(row) > 10 else ""
                question_name = str(row[11] or "").strip() if len(row) > 11 else ""
            else:
                option_a = str(row[2] or "").strip()
                option_b = str(row[3] or "").strip()
                option_c = str(row[4] or "").strip() if len(row) > 4 else ""
                option_d = str(row[5] or "").strip() if len(row) > 5 else ""
                option_e = ""; option_f = ""; option_g = ""; option_h = ""
                answer = str(row[6] or "").strip().upper() if len(row) > 6 else ""
                question_name = str(row[7] or "").strip() if len(row) > 7 else ""

            # 学科/分辨率/年级/阶段/难度/知识点/图片/背景/反馈/条件反馈 — 根据格式版本
            if is_new_format:
                subject_raw = str(row[12] or "").strip() if len(row) > 12 else ""
                resolution = str(row[13] or "").strip() if len(row) > 13 else ""
                grade_raw = str(row[14] or "").strip().replace("－", "-").replace("—", "-") if len(row) > 14 else ""
                stage_raw = str(row[15] or "").strip() if len(row) > 15 else ""
                difficulty_raw = str(row[16] or "0").strip() if len(row) > 16 and row[16] else "0"
                knowledge = str(row[17] or "").strip() if len(row) > 17 else ""
                question_images_raw = str(row[18] or "").strip() if len(row) > 18 else ""
                background_url = str(row[19] or "").strip() if len(row) > 19 else ""
                wrong_feedback = str(row[20] or "").strip() if len(row) > 20 else ""
                correct_feedback = str(row[21] or "").strip() if len(row) > 21 else ""
                wrong_fb_img = str(row[22] or "").strip() if len(row) > 22 else ""
                correct_fb_img = str(row[23] or "").strip() if len(row) > 23 else ""
                wrong_fb_audio = str(row[24] or "").strip() if len(row) > 24 else ""
                correct_fb_audio = str(row[25] or "").strip() if len(row) > 25 else ""
                cond_min_raw = str(row[26] or "").strip() if len(row) > 26 else ""
                cond_max_raw = str(row[27] or "").strip() if len(row) > 27 else ""
                cond_fb_text = str(row[28] or "").strip() if len(row) > 28 else ""
                cond_fb_img = str(row[29] or "").strip() if len(row) > 29 else ""
                cond_fb_audio = str(row[30] or "").strip() if len(row) > 30 else ""
            else:
                # 旧格式：反馈在 cols 8-13，学科在 col 14
                wrong_feedback = str(row[8] or "").strip() if len(row) > 8 else ""
                correct_feedback = str(row[9] or "").strip() if len(row) > 9 else ""
                wrong_fb_img = str(row[10] or "").strip() if len(row) > 10 else ""
                correct_fb_img = str(row[11] or "").strip() if len(row) > 11 else ""
                wrong_fb_audio = str(row[12] or "").strip() if len(row) > 12 else ""
                correct_fb_audio = str(row[13] or "").strip() if len(row) > 13 else ""
                subject_raw = str(row[14] or "").strip() if len(row) > 14 else ""
                resolution = str(row[15] or "").strip() if len(row) > 15 else ""
                grade_raw = str(row[16] or "").strip().replace("－", "-").replace("—", "-") if len(row) > 16 else ""
                stage_raw = str(row[17] or "").strip() if len(row) > 17 else ""
                difficulty_raw = str(row[18] or "0").strip() if len(row) > 18 and row[18] else "0"
                knowledge = str(row[19] or "").strip() if len(row) > 19 else ""
                question_images_raw = str(row[20] or "").strip() if len(row) > 20 else ""
                background_url = str(row[21] or "").strip() if len(row) > 21 else ""
                cond_min_raw = ""; cond_max_raw = ""; cond_fb_text = ""; cond_fb_img = ""; cond_fb_audio = ""
            # 学科映射
            if subject_raw.isdigit():
                subject_id = int(subject_raw)
            elif subject_raw in SUBJECT_MAP:
                subject_id = SUBJECT_MAP[subject_raw]
            else:
                subject_id = 0
            # 年级映射
            if grade_raw.isdigit():
                grade_val = int(grade_raw)
            elif grade_raw in GRADE_MAP:
                grade_val = GRADE_MAP[grade_raw]
            else:
                grade_val = 0
            # 阶段映射
            stage_val = 0
            if stage_raw.isdigit():
                stage_val = int(stage_raw)
            elif stage_raw in STAGE_MAP:
                stage_val = STAGE_MAP[stage_raw]
            # 难度映射
            if difficulty_raw.isdigit():
                difficulty = int(difficulty_raw)
            elif difficulty_raw in DIFFICULTY_MAP:
                difficulty = DIFFICULTY_MAP[difficulty_raw]
            else:
                difficulty = 0
            difficulty = max(0, min(5, difficulty))

            cond_min_wrong = 0
            cond_max_wrong = 0
            if cond_min_raw and cond_min_raw.replace("-","").isdigit():
                cond_min_wrong = max(0, int(cond_min_raw))
            if cond_max_raw and cond_max_raw.replace("-","").isdigit():
                cond_max_wrong = max(0, int(cond_max_raw))

            # 构建条件反馈列表
            condition_feedbacks = []
            if cond_fb_text.strip():
                condition_feedbacks.append({
                    "minWrong": cond_min_wrong,
                    "maxWrong": cond_max_wrong,
                    "content": cond_fb_text,
                    "imageUrl": cond_fb_img,
                    "audioUrl": cond_fb_audio,
                })

            # 解析题目图片：分号分隔 → [{url, width, height}]
            question_images = []
            if question_images_raw:
                for url in question_images_raw.split(";"):
                    url = url.strip()
                    if url:
                        question_images.append({"url": url, "width": 0, "height": 0})

            print(f"[PARSE] row{i} stem={repr(stem[:30])} A={repr(option_a[:10])} B={repr(option_b[:10])} ans={repr(answer)} "
                  f"imgs={len(question_images)} bg={'YES' if background_url else 'NO'} "
                  f"name={repr(question_name[:20])}")
            
            if not stem or not option_a or not option_b or not answer:
                miss = []
                if not stem: miss.append("题干")
                if not option_a: miss.append("A")
                if not option_b: miss.append("B")
                if not answer: miss.append("答案")
                print(f"[PARSE] row{i} 跳过 — 缺少: {', '.join(miss)} "
                      f"(row长度={len(row)})")
                continue

            options = [
                {"label": "A", "content": option_a},
                {"label": "B", "content": option_b},
            ]
            if option_c:
                options.append({"label": "C", "content": option_c})
            if option_d:
                options.append({"label": "D", "content": option_d})
            if option_e:
                options.append({"label": "E", "content": option_e})
            if option_f:
                options.append({"label": "F", "content": option_f})
            if option_g:
                options.append({"label": "G", "content": option_g})
            if option_h:
                options.append({"label": "H", "content": option_h})

            # 关键修改：整个文本内容作为一个整体，不按换行拆分
            contexts = [contexts_text] if contexts_text else []

            questions.append({
                "stem": stem,
                "contexts": contexts,
                "options": options,
                "answer": answer,
                "name": question_name,
                "subjectId": subject_id,
                "resolution": resolution,
                "grade": grade_val,
                "stage": stage_val,
                "difficulty": difficulty,
                "knowledge": knowledge,
                "questionImages": question_images,
                "backgroundUrl": background_url,
                "wrongFeedback": wrong_feedback,
                "correctFeedback": correct_feedback,
                "wrongFeedbackImg": wrong_fb_img,
                "correctFeedbackImg": correct_fb_img,
                "wrongFeedbackAudio": wrong_fb_audio,
                "correctFeedbackAudio": correct_fb_audio,
                "conditionFeedbacks": condition_feedbacks,
            })

        wb.close()
        return {"success": True, "questions": questions, "total": len(questions)}

    except Exception as e:
        return {"success": False, "error": f"解析失败: {str(e)}"}