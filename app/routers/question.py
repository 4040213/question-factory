"""试题批量生成路由。"""

import asyncio
import io
import json
import re
import urllib.parse
import uuid as uuid_mod
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook
from pydantic import BaseModel

from ..client import MortonClient, MortonClientError
from config import SUBJECT_MAP, GRADE_MAP, STAGE_MAP, DIFFICULTY_MAP

router = APIRouter(prefix="/api/question", tags=["试题"])

QUESTION_TYPE_CHOICE = "CHOICE"
QUESTION_TYPE_BLANK = "BLANK"
QUESTION_TYPE_COMPLETION = "COMPLETION"

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
    # ★ 标准编码（公司素材编码，如 26sp_G2W13-P1-1）：与文档「名称」(友好名) 互为别名，
    #    用于匹配按公司编码命名的图片/音频文件（如 26sp_G2W13-P1-1_tips_1.mp3）。
    code: str = ""
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
    # ★ 题干音频（朗读/配音）：独立于反馈音频，存 CDN URL。有则 stem.webAudioUrl 设上、文字保留（行为B共存）
    stemAudio: str = ""
    conditionFeedbacks: List[ConditionFeedback] = []   # 条件反馈列表（基于错误次数）
    # ★ AI辅助提示 → 自动提示（全题型通用，不写死）
    # 结构：[{imageUrl, audioUrl, variables: [{operation, timeSeconds}]}]
    autoHints: Optional[List[Dict[str, Any]]] = None


class BatchRequest(BaseModel):
    questions: List[Dict[str, Any]]
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
    # ★ 通用语音反馈（一套复用于全部题目）：{right, wrong1, wrong2, wrong3} = 莫顿CDN音频URL
    # 有语音时反馈强制 answerType="AUDIO"（莫顿规则：有音频不能混搭图文），无语音则回退文字
    defaultVoiceFeedback: Optional[Dict[str, str]] = None
    # ★ 题干音频（全局默认，复用于全部题目，与 defaultVoiceFeedback 同理）；逐题 stemAudio 优先
    defaultStemAudio: str = ""
    # ★ AI辅助提示 → 自动提示（全题型通用，不写死）
    # 结构：[{imageUrl, audioUrl, variables: [{operation, timeSeconds}]}]
    # operation: NO_ACTION / NOT_SUBMITTED / NOT_PASSED
    # timeSeconds: 触发时间（秒）
    defaultAutoHints: Optional[List[Dict[str, Any]]] = None
    # ═══ 优化口径（命名自动化 / 每题各自音频）═══
    # 未来「组课文档→Excel」步骤若能做到「每题各自的反馈音频」，把素材清单（manifest）
    # 一并传上来，结构为：
    #   { "audioMap": { "<题目id/name>": {"right": url, "wrong":[{"minWrong":N,"audio":url}]} } }
    # 此处仅声明字段占位，不改动当前行为。消费点在 _apply_defaults 见 TODO(PER_QUESTION_AUDIO)。
    materials_manifest: Optional[Dict[str, Any]] = None


class PreviewCreateRequest(BaseModel):
    question: Dict[str, Any]
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
    # ★ 通用语音反馈（一套复用于全部题目）：{right, wrong1, wrong2, wrong3} = 莫顿CDN音频URL
    # 有语音时反馈强制 answerType="AUDIO"（莫顿规则：有音频不能混搭图文），无语音则回退文字
    defaultVoiceFeedback: Optional[Dict[str, str]] = None
    # ★ 题干音频（全局默认，复用于全部题目，与 defaultVoiceFeedback 同理）；逐题 stemAudio 优先
    defaultStemAudio: str = ""
    # ★ AI辅助提示 → 自动提示（全题型通用）
    defaultAutoHints: Optional[List[Dict[str, Any]]] = None


class PreviewPayloadRequest(BaseModel):
    """创建前本地预览请求：只构建最终 questionData，不调用莫顿 insert/update。"""
    question: Dict[str, Any]
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
    # ★ AI辅助提示 → 自动提示（全题型通用）
    defaultAutoHints: Optional[List[Dict[str, Any]]] = None


# ============================================================
#  语音反馈工具函数
#  莫顿规则（实测 26018 正常题）：反馈含音频时必须用 answerType="AUDIO"
#  且音频地址放 webUrl 字段（webAudioUrl 反而为 null），不能 PICTURE_CONTENT(图文)
#  里再挂音频（非法混搭 → 渲染器"加载配置失败"）。
#  故策略：有语音 → 纯 AUDIO(webUrl=mp3, 无 content)；无语音 → 回退文字。
# ============================================================

# 负反馈文案回退（仅当对应语音缺失时使用）
_FALLBACK_WRONG = [
    "用一一对应的方法来比较它们的多少吧！",
    "一一对应以后，有剩余的小动物数量多！",
    "好可惜啊，不过没关系，等所有关卡结束后，再来试一试吧！",
]
_FALLBACK_RIGHT = "恭喜你，回答正确~"

# 正反馈"无"判定：文档写"无"/"无正反馈"/空等 → 不生成 rightQuestionFeedback
_NONE_FEEDBACK_VALUES = {"", "无", "无正反馈", "无反馈", "-", "/", "none", "null", "n/a"}
def _is_positive_feedback_none(value: str) -> bool:
    """判断正反馈字段是否表示'不需要答对反馈'（按文档原值，不硬编码题目类型）。"""
    return (value or "").strip().lower() in _NONE_FEEDBACK_VALUES


def _merge_voice_to_existing_cfs(item, vf: dict) -> None:
  """将通用语音音频 URL 合并到已有的 conditionFeedbacks（保留文档的 content）。
  匹配规则（优先级从高到低）：
    1. 精确匹配：CF.minWrong == n → 注入 audioUrl
    2. 范围匹配：CF.minWrong <= n <= CF.maxWrong → 注入并归一为单点(n,n)
       （兼容旧 Excel「最少=1 最多=3」但文本实为"负反馈3"的场景）
  都匹配不到时才追加新条目。
  """
  existing = getattr(item, "conditionFeedbacks", None) or []
  if not existing:
      # 文档没有任何条件反馈 → 按语音角色新建（旧行为，兜底）
      item.conditionFeedbacks = [
          ConditionFeedback(minWrong=n, maxWrong=n, audioUrl=vf[key])
          for n, key in [(1,"wrong1"),(2,"wrong2"),(3,"wrong3")] if vf.get(key)
      ]
      return
  # 已有文档条件反馈 → 对每个语音角色找能接住的 CF
  merged = list(existing)
  audio_injected = set()
  for n, key in [(1,"wrong1"),(2,"wrong2"),(3,"wrong3")]:
      url = vf.get(key, "")
      if not url:
          continue
      matched = False
      for cf in merged:
          mn = getattr(cf, 'minWrong', 0)
          mx = getattr(cf, 'maxWrong', 0)
          # 范围匹配：n 落在 [mn, mx] 内且该 CF 尚无音频
          if mn <= n <= mx and not getattr(cf, 'audioUrl', ''):
              cf.audioUrl = url
              # 归一为单点（文档意图就是第 N 次错的反馈）
              cf.minWrong = n
              cf.maxWrong = n
              audio_injected.add(n)
              matched = True
              break
      # 该语音角色无人认领 → 仅当没有同 minWrong 条目时追加
      if not matched and n not in audio_injected:
          if not any(getattr(c, 'minWrong', 0) == n for c in merged):
              merged.append(ConditionFeedback(minWrong=n, maxWrong=n, audioUrl=url))
  item.conditionFeedbacks = merged


def _clean_feedback_label(text: str) -> str:
    """清洗反馈内容中的标签前缀，如「负反馈3：」「第3次错误反馈」等。"""
    import re as _re
    t = text.strip()
    # 去掉 "负反馈N：" / "负反馈N—" / "第N次错误(反馈)?[：:]\s*"
    t = _re.sub(r'^(负反馈\s*\d+[\s：:—\-]+|(第\s*\d+\s*次?\s*(错误|反馈)[\s：:—\-]+))', '', t).strip()
    # 去掉末尾残留标点
    t = t.rstrip('：:—- ')
    return t


def _audio_fb(url: str, variable_value: Optional[str] = None) -> Dict[str, Any]:
    """构造一个纯音频反馈对象（莫顿 AUDIO 类型）。"""
    fb: Dict[str, Any] = {
        "webUrl": url,
        "webName": url.rsplit("/", 1)[-1] if "/" in url else url,
        "webAudioUrl": None,
        "logicTypeList": [],
        "logicExpressList": [],
        "answerType": "AUDIO",
    }
    if variable_value:
        n = int(variable_value.split(",")[0])
        fb["webListenList"] = [{
            "variableKey": "WrongNumber",
            "variableType": "int",
            "compare": "EQUAL",
            "value": n,
            "variableValue": variable_value,
        }]
    else:
        fb["webListenList"] = []
    return fb


def _wrong_fb(audio_url: str, variable_value: str, idx: int) -> Dict[str, Any]:
    """第 idx 次错误对应的非通用反馈：有音频用 AUDIO，否则回退文字。"""
    if audio_url:
        return _audio_fb(audio_url, variable_value)
    return {
        "webUrl": None,
        "webAudioUrl": None,
        "webListenList": [{
            "variableKey": "WrongNumber",
            "variableType": "int",
            "compare": "EQUAL",
            "value": int(variable_value.split(",")[0]),
            "variableValue": variable_value,
        }],
        "logicTypeList": [],
        "logicExpressList": [],
        "answerType": "PICTURE_CONTENT",
        "content": _FALLBACK_WRONG[idx],
    }


# ============================================================
#  AI辅助提示 → 自动提示 工具函数（全题型通用，字段已对齐莫顿样例）
# ============================================================

# 操作状态：中文标签 → 莫顿 variableKey 枚举
AUTO_HINT_OP_MAP = {
    "无操作": "NO_ACTION",
    "未提交": "NOT_SUBMITTED",
    "未通关": "NOT_PASSED",
}
AUTO_HINT_OP_ENUM = set(AUTO_HINT_OP_MAP.values())

# ★ 自动提示（小灯泡）数量与每个提示的变量列数（Excel 列数，集中定义避免散落写死）
AUTO_HINT_COUNT = 3        # 每题最多 3 个自动提示
AUTO_HINT_VAR_COUNT = 6    # 每个自动提示最多 6 个变量（列：状态:时间）
# 注：feishu.py 侧以同名局部常量保持同步（避免跨模块循环导入）。


def _norm_auto_hint_op(op: str) -> str:
    """中文标签 / 枚举值 → 莫顿 variableKey；无法识别则原样返回便于排查。"""
    op = (op or "").strip()
    if op in AUTO_HINT_OP_MAP:
        return AUTO_HINT_OP_MAP[op]
    if op in AUTO_HINT_OP_ENUM:
        return op
    return op


def _parse_one_auto_hint(auto_img: str, auto_audio: str,
                         var_raws: List[str]) -> Optional[Dict[str, Any]]:
    """从单条自动提示的 Excel 列（图片/音频/变量1~6）构建 autoHints 元素。

    每变量列格式『状态:时间』（冒号可为全角；状态可为中文或 Merton 枚举；时间单位秒）。
    仅当 图片 / 音频 / 任一有效变量 有内容时返回该提示 dict，否则返回 None（不注入）。
    operation 保留用户原值（中文或枚举），由 _build_auto_hint_entries → _norm_auto_hint_op 归一，
    与前端 collectAutoHints 输出的数据结构完全一致。
    """
    auto_img = (auto_img or "").strip()
    auto_audio = (auto_audio or "").strip()
    variables: List[Dict[str, Any]] = []
    for raw in (var_raws or []):
        raw = (raw or "").strip()
        if not raw:
            continue
        m = re.match(r'^(.+?)[\s:：]+(\d+)$', raw)
        if m:
            op = m.group(1).strip()
            secs = int(m.group(2))
        else:
            # 仅状态无时间（或格式异常）→ 保留状态，时间记 0（_build_auto_hint_entries 会过滤无效项）
            op = raw
            secs = 0
        variables.append({"operation": op, "timeSeconds": secs})
    if not auto_img and not auto_audio and not variables:
        return None
    return {
        "imageUrl": auto_img,
        "audioUrl": auto_audio,
        "variables": variables,
    }


def _parse_auto_hint_columns(hint_specs: List[tuple]) -> Optional[List[Dict[str, Any]]]:
    """从 N 条自动提示的列规格（每条 = (图片, 音频, 变量1~6)）构建稠密的 autoHints 列表。

    hint_specs 顺序即提示序号（1-based 在外层循环）。空提示被跳过，最终列表稠密
    （索引即提示序-1），与前端 q.autoHints 及莫顿 autoTip.list 一致。
    全部为空则返回 None（不注入 autoHints）。
    """
    hints: List[Dict[str, Any]] = []
    for spec in hint_specs:
        one = _parse_one_auto_hint(*spec)
        if one:
            hints.append(one)
    return hints if hints else None


def _derive_hint_filename(url: str) -> str:
    """从 CDN URL 提取原始文件名（去 query/hash、URL 解码）。"""
    if not url:
        return ""
    try:
        name = urllib.parse.unquote(url.split("?")[0].split("#")[0].rstrip("/").split("/")[-1])
    except Exception:
        name = url.split("?")[0].rstrip("/").split("/")[-1]
    return name


def _build_auto_hint_entries(hints: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    将前端收集的自动提示 → 莫顿 auxiliaryTips.autoTip.list 元素列表。
    输入：[{imageUrl, audioUrl, variables: [{operation, timeSeconds}]}]
    输出：[{type, uuid, audioName, pictureName, webAudioUrl, webListenList, webPictureUrl, webSpineUrl, webVideoUrl}]
    变量「操作状态 + 时间(秒)」两者都需有内容才有效（对齐用户要求）。
    """
    entries = []
    for h in hints:
        img = (h.get("imageUrl") or "").strip()
        aud = (h.get("audioUrl") or "").strip()
        if not img and not aud:
            continue  # 无媒体则跳过（无效提示）
        # 变量列表：仅保留 操作+时间 都有效的项
        listen = []
        for v in (h.get("variables") or []):
            op = _norm_auto_hint_op(v.get("operation", ""))
            ts = v.get("timeSeconds", 0)
            if op and ts and float(ts) > 0:
                listen.append({
                    "variableKey": op,
                    "variableType": "int",
                    "variableValue": str(int(float(ts))),
                })
        entries.append({
            "type": "PICTURE_POP_WINDOWS",
            "uuid": str(uuid_mod.uuid4()),
            "audioName": _derive_hint_filename(aud),
            "pictureName": _derive_hint_filename(img),
            "webAudioUrl": aud or None,
            "webListenList": listen,
            "webPictureUrl": img or None,
            "webSpineUrl": None,
            "webVideoUrl": None,
        })
    return entries


def _build_auxiliary_tips(entries: List[Dict[str, Any]],
                          coordinate: Optional[Dict[str, int]] = None,
                          direction: str = "down") -> Dict[str, Any]:
    """组装完整 auxiliaryTips（全题型统一结构）。"""
    if coordinate is None:
        coordinate = {"x": 0, "y": 0}
    return {
        "tips": [],
        "autoTip": {
            "coordinate": coordinate,
            "direction": direction,
            "list": entries,
        },
        "parses": [],
        "experiments": [],
    }


class BlankField(BaseModel):
    varName: str
    blankType: str = "CONTENT"
    answerType: str = "CONTENT"
    inputStyle: str = "TEXT"
    answer: str = ""
    aliases: List[str] = []
    keyboardStyle: str = "HORIZONTAL_DOUBLE_LINE"
    showNumber: bool = True
    showBox: bool = False
    options: List[str] = []
    optionImages: List[dict] = []  # [{index, url}, ...] 填空选择题（选项型空）的选项图


class BlankQuestionItem(BaseModel):
    stem: str
    contexts: List[str] = []
    name: str = ""
    # ★ 标准编码（公司素材编码）：与「名称」互为别名，用于匹配按公司编码命名的素材文件
    code: str = ""
    subjectId: int = 0
    resolution: str = "1334*750"
    grade: int = 0
    stage: int = 0
    difficulty: int = 0
    knowledge: str = ""
    questionType: str = QUESTION_TYPE_COMPLETION
    contentMarkup: str = ""
    stemAudio: str = ""
    blanks: List[BlankField] = []
    questionImages: List[dict] = []
    backgroundUrl: str = ""
    wrongFeedback: str = ""
    correctFeedback: str = ""
    wrongFeedbackImg: str = ""
    correctFeedbackImg: str = ""
    wrongFeedbackAudio: str = ""
    correctFeedbackAudio: str = ""
    conditionFeedbacks: List[ConditionFeedback] = []
    # ★ AI辅助提示 → 自动提示（全题型通用，不写死）
    autoHints: Optional[List[Dict[str, Any]]] = None


class BlankBatchRequest(BaseModel):
    questions: List[Dict[str, Any]]
    subjectId: int = 0
    cookie: str = ""
    token: str = ""
    defaultName: str = "新建填空题"
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
    # ★ 通用语音反馈（一套复用于全部题目）
    defaultVoiceFeedback: Optional[Dict[str, str]] = None
    # ★ 题干音频（全局默认，复用于全部题目，与 defaultVoiceFeedback 同理）；逐题 stemAudio 优先
    defaultStemAudio: str = ""


def _normalize_blank_value(value: str, mapping: Dict[str, str], default: str) -> str:
    text = str(value or "").strip()
    if not text:
        return default
    return mapping.get(text, text).upper()


def _to_model(item: Any):
    if isinstance(item, dict):
        if item.get("questionType") in (QUESTION_TYPE_COMPLETION, QUESTION_TYPE_BLANK):
            item = {**item, "questionType": QUESTION_TYPE_COMPLETION}
            return BlankQuestionItem(**item)
        return QuestionItem(**item)
    return item


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
    # ★ 学科防御默认：subjectId 绝对不能为 0
    # 莫顿渲染器靠 subjectId 加载题模板/皮肤，0 会导致预览"加载配置失败"
    # Excel 未填 且 下拉框未选(=0) 时，回退到数学(4)
    if item.subjectId == 0:
        item.subjectId = 4
        print(f"[SUBJECT_DEFAULT] ⚠️ subjectId 回退为 4(数学)：Excel 未填且下拉框未选学科，"
              f"否则莫顿预览会报'加载配置失败'")

    is_completion = getattr(item, "questionType", QUESTION_TYPE_CHOICE) == QUESTION_TYPE_COMPLETION

    if (not getattr(item, "wrongFeedback", "") or getattr(item, "wrongFeedback", "").strip() == "") and body.defaultWrongFeedback:
        item.wrongFeedback = body.defaultWrongFeedback
    if (not getattr(item, "correctFeedback", "") or getattr(item, "correctFeedback", "").strip() == "") and body.defaultCorrectFeedback:
        item.correctFeedback = body.defaultCorrectFeedback
    if (not getattr(item, "wrongFeedbackImg", "") or getattr(item, "wrongFeedbackImg", "").strip() == "") and body.defaultWrongImg:
        item.wrongFeedbackImg = body.defaultWrongImg
    if (not getattr(item, "correctFeedbackImg", "") or getattr(item, "correctFeedbackImg", "").strip() == "") and body.defaultCorrectImg:
        item.correctFeedbackImg = body.defaultCorrectImg
    if (not getattr(item, "wrongFeedbackAudio", "") or getattr(item, "wrongFeedbackAudio", "").strip() == "") and body.defaultWrongAudio:
        item.wrongFeedbackAudio = body.defaultWrongAudio
    if (not getattr(item, "correctFeedbackAudio", "") or getattr(item, "correctFeedbackAudio", "").strip() == "") and body.defaultCorrectAudio:
        item.correctFeedbackAudio = body.defaultCorrectAudio

    # 题目图片全局默认（分号分隔，仅当 Excel 未提供时生效）
    if (not getattr(item, "questionImages", []) or len(getattr(item, "questionImages", [])) == 0) and body.defaultQuestionImages:
        images = []
        for url in body.defaultQuestionImages.split(";"):
            url = url.strip()
            if url:
                images.append({"url": url, "width": 0, "height": 0})
        if images:
            item.questionImages = images

    # 画布背景全局默认
    if (not getattr(item, "backgroundUrl", "") or getattr(item, "backgroundUrl", "").strip() == "") and body.defaultBackgroundUrl:
        item.backgroundUrl = body.defaultBackgroundUrl

    # 条件反馈全局默认（仅当逐题未设置时生效）
    cond_fbs = getattr(item, "conditionFeedbacks", None)
    if (cond_fbs is None or len(cond_fbs) == 0) and body.defaultConditionFeedback:
        cf = body.defaultConditionFeedback
        # 只有填写了反馈文本时才添加
        if cf.content and cf.content.strip():
            item.conditionFeedbacks = [cf]

    # ★ 通用语音反馈（一套复用于全部题目）：defaultVoiceFeedback = {right, wrong1, wrong2, wrong3}
    # ★ 修复：合并到已有 conditionFeedbacks（保留文档 minWrong/maxWrong），不再直接覆盖
    vf = getattr(body, "defaultVoiceFeedback", None)
    if vf:
        if vf.get("right"):
            item.correctFeedbackAudio = vf["right"]
        if any(vf.get(k) for k in ("wrong1","wrong2","wrong3")):
            _merge_voice_to_existing_cfs(item, vf)
            print(f"[VOICE_DEFAULT] 🔗 合并通用语音到条件反馈(保留文档区间): "
                  f"正={'有' if vf.get('right') else '无'}, "
                  f"负={[k for k in ('wrong1','wrong2','wrong3') if vf.get(k)]}")

    # ★ 题干音频全局默认（仅当逐题未设置时生效，与 defaultVoiceFeedback 同理）
    ds = getattr(body, "defaultStemAudio", "")
    if ds and (not getattr(item, "stemAudio", "") or getattr(item, "stemAudio", "").strip() == ""):
        item.stemAudio = ds
        print(f"[STEM_AUDIO] 🔗 注入全局题干音频: {ds[:50]}")

    # ★ AI辅助提示 → 自动提示（全题型通用，不写死）
    # 仅当逐题未设置时生效（与条件反馈/语音同策略）
    dah = getattr(body, "defaultAutoHints", None)
    if dah and not getattr(item, "autoHints", None):
        item.autoHints = dah
        print(f"[AUTO_HINT] 🔗 注入全局自动提示: {len(dah)} 条")

    # TODO(PER_QUESTION_AUDIO) 优化口径：若 body.materials_manifest 带 per-question 音频映射，
    # 在此按 item.name 取出对应音频，覆盖上面「一套复用于全部题」的默认语音：
    #   am = getattr(body, "materials_manifest", None)
    #   if am and am.get("audioMap"):
    #       per = am["audioMap"].get(item.name)
    #       if per and per.get("right"): item.correctFeedbackAudio = per["right"]
    #       if per and per.get("wrong"):
    #           item.conditionFeedbacks = [
    #               ConditionFeedback(minWrong=w["minWrong"], maxWrong=w["minWrong"], audioUrl=w["audio"])
    #               for w in per["wrong"]
    #           ]
    # 当前默认策略仍是「一套语音复用于全部题目」，故此处留空。


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
    resolution: str = "1334*750",
) -> dict:
    """智能画布布局引擎 — 4 阶段核心逻辑。

    参数:
        options_count: 选项数量 (2~4)
        contexts_count: 文本块数量
        context_texts: 文本块内容列表
        question_images: 题目图片 [{url, width, height}, ...]
        background_url: 画布背景图 URL
        resolution: 分辨率字符串，如 "1334*750"

    返回:
        {
            "option_positions": [{x,y,width,height}, ...],
            "text_positions": [{x,y,width,align}, ...],
            "image_positions": [{x,y,width,height,url}, ...],
            "background_url": str,
            "font_size": int,
        }
    """
    # ── 分辨率 → 缩放比例 ──
    if "*" in resolution:
        parts = resolution.replace("高清", "").strip().split("*")
        _rw = int(parts[0]) if parts[0].isdigit() else 1334
        _rh = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 750
    else:
        _rw, _rh = 1334, 750
    scale_x = _rw / 1334.0
    scale_y = _rh / 750.0

    # ── 内容区域：按分辨率等比缩放 ──
    CL = int(182 * scale_x)
    CR = int(1152 * scale_x)
    CT = int(130 * scale_y)
    CB = int(580 * scale_y)
    CW = CR - CL
    CH = CB - CT
    CCX = (CL + CR) // 2

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
    OPTION_H = max(100, min(160, option_available_h))

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
    # ★ 只写入远程可达 URL，本地路径会拖垮渲染器
    if item.backgroundUrl and _is_remote_url(item.backgroundUrl):
        question_data["webBgmUrl"] = item.backgroundUrl
        has_update = True
        print(f"[UPDATE] 设置画布背景(远程): {item.backgroundUrl[:80]}...")
    elif item.backgroundUrl and not _is_remote_url(item.backgroundUrl):
        print(f"[UPDATE] ⚠️ 跳过画布背景(本地路径): {item.backgroundUrl[:80]}...")

    # ===== 反馈部分（音频优先：有音频→AUDIO(webUrl=mp3)，无→文字/图文） =====
    wrong = item.wrongFeedback
    correct = item.correctFeedback
    wrong_img = item.wrongFeedbackImg
    correct_img = item.correctFeedbackImg
    wrong_audio = item.wrongFeedbackAudio
    correct_audio = item.correctFeedbackAudio

    has_feedback = bool(wrong or correct or wrong_img or correct_img or wrong_audio or correct_audio or item.conditionFeedbacks)

    # 通用答错反馈：有音频→纯AUDIO，否则图文/文字
    # ★ 去重：若通用反馈文字已作为某条「条件反馈」出现（典型：组课文档"负反馈N：文字"
    #   被同时写入「答错通用反馈」与「条件反馈文本」两列），则不重复生成
    #   wrongQuestionFeedback，避免学生在答错时连弹两个内容相同的窗口。通用侧的
    #   音频/图片并入匹配到的条件反馈，避免素材丢失。
    _gen_content = _clean_feedback_label(wrong) if wrong else ""
    _match_cf = None
    if _gen_content:
        for _cf in (item.conditionFeedbacks or []):
            if _clean_feedback_label(getattr(_cf, "content", "")) == _gen_content:
                _match_cf = _cf
                break
    if (wrong or wrong_img or wrong_audio) and not _match_cf:
        if wrong_audio:
            question_data["wrongQuestionFeedback"] = _audio_fb(wrong_audio)
        else:
            fb = {"answerType": "PICTURE_CONTENT"}
            if wrong: fb["content"] = wrong
            if wrong_img: fb["webUrl"] = wrong_img; fb["webName"] = wrong_img.rsplit("/",1)[-1] if "/" in wrong_img else wrong_img
            question_data["wrongQuestionFeedback"] = fb
        has_update = True
    elif _match_cf is not None:
        # 通用侧素材并入条件反馈（条件反馈重建时会读取这些字段）
        if wrong_audio and not getattr(_match_cf, "audioUrl", ""):
            _match_cf.audioUrl = wrong_audio
            print(f"[FB_DEDUP] 通用答错音频并入条件反馈: {_gen_content!r}")
        if wrong_img and not getattr(_match_cf, "imageUrl", ""):
            _match_cf.imageUrl = wrong_img
        print(f"[FB_DEDUP] ⏭️ 通用答错反馈与条件反馈文本重复，已去重: {_gen_content!r}")

    # 答对通用反馈（正反馈）：文档优先
    # ★ correctFeedback 写"无"/空 → 不生成 rightQuestionFeedback（莫顿允许缺失）
    if not _is_positive_feedback_none(correct):
        if correct or correct_img or correct_audio:
            if correct_audio:
                question_data["rightQuestionFeedback"] = _audio_fb(correct_audio)
            else:
                fb = {"answerType": "PICTURE_CONTENT"}
                if correct: fb["content"] = correct
                if correct_img: fb["webUrl"] = correct_img; fb["webName"] = correct_img.rsplit("/",1)[-1] if "/" in correct_img else correct_img
                question_data["rightQuestionFeedback"] = fb
            has_update = True
    else:
        print(f"[RIGHT_FB_UPDATE] ⏭️ 跳过答对反馈(文档标记'无'): name={item.name!r}, correctFeedback={correct!r}")

    # 条件反馈（基于错误次数）
    if item.conditionFeedbacks:
        cfbs = []
        for cf in item.conditionFeedbacks:
            # ★ 有音频 → 纯 AUDIO（webUrl=mp3），不能图文挂音频
            if cf.audioUrl:
                n = cf.minWrong
                vv = f"{n},{n}"
                fb = _audio_fb(cf.audioUrl, vv)
            else:
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

            # ★ variableValue 格式：莫顿要求 "N,N"（两数字一致 = 单点触发）
            # 即使 Excel 填了区间(如 最少=1 最多=3)，也取 minWrong 作为单点
            # 因为每条 conditionFeedback 对应一个具体错误次数（第N次错）
            listen = fb["webListenList"][0]
            n = cf.minWrong
            listen["compare"] = "EQUAL"
            listen["value"] = n
            listen["variableValue"] = f"{n},{n}"
            # 清洗 content 标签
            if cf.content:
                fb["content"] = _clean_feedback_label(cf.content)

            cfbs.append(fb)
        question_data["questionFeedbacks"] = cfbs
        has_update = True

    # ===== ★★★ 终极防护：确保 feedback 结构正确（不再强制3条） ★★★ =====
    # 莫顿要求：rightQuestionFeedback=答对通用反馈(纯正反馈) / questionFeedbacks=N条独立判定(纯负反馈)
    # N 可以是 0~任意数，由组课文档/Excel 的 conditionFeedbacks 决定
    existing_fbs = question_data.get("questionFeedbacks", [])

    needs_fix = False
    for idx, fb in enumerate(existing_fbs):
        vv = ""
        listen_list = fb.get("webListenList", [])
        if listen_list and isinstance(listen_list, list) and len(listen_list) > 0:
            vv = listen_list[0].get("variableValue", "")
        content = fb.get("content", "")
        # 缺少 variableValue → 修正
        if not vv:
            needs_fix = True
            print(f"[FB_FIX] ⚠️ 判定{idx+1} 缺少 variableValue，需要修正")
        # 内容含"正反馈"标签残留 → 清洗（不应出现在负反馈中）
        if "正反馈" in content:
            needs_fix = True
            print(f"[FB_FIX] ⚠️ 判定{idx+1} content 含'正反馈'标签残留，将清洗")

    if needs_fix:
        print(f"[FB_FIX] 🔧 修正 {len(existing_fbs)} 条 questionFeedbacks 的结构问题...")
        fixed_fbs = []
        for idx, fb in enumerate(existing_fbs):
            fb_copy = dict(fb)
            # 清洗 content 标签
            raw_content = fb_copy.get("content", "")
            fb_copy["content"] = _clean_feedback_label(raw_content)
            # 确保有 webListenList + variableValue
            listen_list = fb_copy.get("webListenList", [])
            if not (listen_list and isinstance(listen_list, list) and len(listen_list) > 0
                    and listen_list[0].get("variableValue")):
                # 尝试从已有音频推断，或使用默认值
                n = idx + 1  # fallback: 第i条 → 第i次错误
                fb_copy["webListenList"] = [{
                    "variableKey": "WrongNumber",
                    "variableType": "int",
                    "compare": "EQUAL",
                    "value": n,
                    "variableValue": f"{n},{n}",
                }]
            else:
                # 确保 variableValue 格式为 N,N（莫顿要求两数字一致）
                ll = fb_copy["webListenList"][0]
                old_vv = ll.get("variableValue", "")
                if "," in old_vv:
                    parts = old_vv.split(",")
                    if len(parts) == 2 and parts[0] != parts[1]:
                        # 区间格式(如"1,3")→取min作为单点(如"1,1")
                        n = parts[0]
                        ll["variableValue"] = f"{n},{n}"
                        ll["value"] = int(n)
                        print(f"[FB_FIX]   判定{idx+1} variableValue {old_vv} → {ll['variableValue']}（区间→单点）")
            fixed_fbs.append(fb_copy)
        question_data["questionFeedbacks"] = fixed_fbs
        has_update = True

    # ★ 文档标记"无"时，不强制补回 rightQuestionFeedback（莫顿允许缺失）
    correct_fb_text = (getattr(item, 'correctFeedback', '') or '').strip()
    if _is_positive_feedback_none(correct_fb_text):
        # 文档明确写了无正反馈 → 确保不残留
        if "rightQuestionFeedback" in question_data:
            del question_data["rightQuestionFeedback"]
        return has_update, has_feedback

    # 确保 rightQuestionFeedback 存在且不含负反馈（保留已有音频）
    rfq = question_data.get("rightQuestionFeedback", {})
    rfq_is_broken = (not rfq) or ("负反馈" in rfq.get("content", "")) or \
                   (rfq.get("answerType") == "PICTURE_CONTENT" and "正反馈" in rfq.get("content", ""))
    if rfq_is_broken:
        if rfq.get("answerType") == "AUDIO" and rfq.get("webUrl"):
            pass  # 已有音频，保留
        else:
            question_data["rightQuestionFeedback"] = {
                "answerType": "PICTURE_CONTENT",
                "content": _FALLBACK_RIGHT,
            }
            has_update = True
            print("[FB_FIX] 🔧 修正 rightQuestionFeedback（纯正反馈/音频）")

    return has_update, has_feedback


async def _update_question(client, question_data: Dict[str, Any], question_id, item) -> str:
    """通过单次 /update 写入学科、年级、阶段、难度 + 反馈（合并两次调用为一次）。"""
    update_data = dict(question_data)
    update_data["questionId"] = question_id
    has_update, has_feedback = _merge_update_fields(update_data, item)

    if not has_update:
        return ""

    # ★ 诊断：打印 update 时 options + 反馈状态
    qi_options = update_data.get("questionItem", {}).get("options", [])
    for idx, o in enumerate(qi_options):
        wurl = o.get('webOptionUrl', '')
        print(f"[UPDATE_OPT] 选项{idx}: webOptionUrl={'有: ' + wurl[:60] if wurl else '无'}, content='{o.get('content','')[:30]}'")
    print(f"[UPDATE_WEB] webBgmUrl={update_data.get('webBgmUrl', '无')[:60] if update_data.get('webBgmUrl') else '无'}")
    print(f"[UPDATE_PIC] pictures个数={len(update_data.get('pictures', []))}")

    # ★★ 关键诊断：打印即将发给莫顿 update API 的最终反馈结构 ★★
    rfq = update_data.get("rightQuestionFeedback", {})
    fbs = update_data.get("questionFeedbacks", [])
    fb_vv = [f.get("webListenList", [{}])[0].get("variableValue", "❌缺失") for f in fbs]
    fb_content_preview = [f.get("content", "")[:30].replace("\n", "\\n") for f in fbs]
    print(f"[FINAL_FB] → 莫顿 update payload:")
    print(f"  rightQuestionFeedback.content='{rfq.get('content', '(无)')[:40]}'")
    print(f"  questionFeedbacks数={len(fbs)}, variableValues={fb_vv}")
    for i, c in enumerate(fb_content_preview):
        print(f"    判定{i+1} content='{c}...'")

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


def _is_remote_url(url: str) -> bool:
    """检测 URL 是否为可被莫顿云渲染器访问的远程地址。

    莫顿渲染器运行在云端，只能访问 http(s) URL。
    本地文件路径（O:\\、D:\\、C:\\、//server/share 等）会导致图片加载失败 → 渲染崩溃 → "加载配置失败"
    """
    if not url:
        return False
    s = url.strip()
    # 合法的远程 URL 必须以 http:// 或 https:// 开头
    if s.startswith("http://") or s.startswith("https://"):
        return True
    # 其他一切（Windows 路径、UNC 路径、相对路径等）都是本地/不可达的
    return False


def _build_choice_payload(item: QuestionItem, subject_id: int) -> Dict[str, Any]:
    """构建选择题 insert 接口的完整 payload。"""
    # ════════════════════════════════════════════════
    #  阶段 0：图片 URL 预处理 — 过滤本地不可达路径
    # ════════════════════════════════════════════════
    original_qimgs = item.questionImages or []
    remote_question_images = [qi for qi in original_qimgs if _is_remote_url(qi.get("url", ""))]
    local_qimg_count = len(original_qimgs) - len(remote_question_images)

    original_oimgs = item.optionImages or []
    remote_option_images = [oi for oi in original_oimgs if _is_remote_url(oi.get("url", ""))]
    local_oimg_count = len(original_oimgs) - len(remote_option_images)

    remote_bg = item.backgroundUrl if _is_remote_url(item.backgroundUrl or "") else ""
    # ★ 规则：题目有「远程可达的背景图」时，题干不再添加文本（背景图已含内容）
    # 注意：本地背景图(不可达)不触发本规则——否则莫顿无法显示背景图、文本又被清空，题目会变空白
    has_bg = bool(remote_bg)

    if local_qimg_count > 0 or local_oimg_count > 0 or (item.backgroundUrl and not remote_bg):
        print(f"[URL_FILTER] ⚠️ 过滤本地路径: {local_qimg_count}个题目图/{local_oimg_count}个选项图/"
              f"{('1背景图为本地' if item.backgroundUrl and not remote_bg else '背景图OK')}")
        if local_qimg_count > 0:
            for qi in original_qimgs:
                if not _is_remote_url(qi.get("url", "")):
                    print(f"  → 舍弃题目图(本地): {qi.get('url', '')[:60]}")

    # ★ 诊断日志：打印过滤后的图片字段状态
    print(f"[IMG_DEBUG] 题目[{item.name or '?'}] "
          f"questionImages={len(remote_question_images)}个(原{len(original_qimgs)}个), "
          f"backgroundUrl={'有(' + remote_bg[:40] + ')' if remote_bg else ('有(本地!)' if item.backgroundUrl else '无')}, "
          f"optionImages={[(oi.get('index'), oi.get('url','')[:40]) for oi in remote_option_images]}")
    if remote_option_images:
        for oi in remote_option_images:
            print(f"  → optionImages[{oi.get('index')}] = {oi.get('url', '')[:60]}...")
    if remote_question_images:
        for qi in remote_question_images:
            print(f"  → questionImages(远程) url = {qi.get('url', '')[:60]}...")

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
    # 注意：只传入过滤后的远程图片，本地路径图片不参与布局（云端不可达）
    layout = _calculate_smart_layout(
        options_count=len(item.options),
        contexts_count=len(item.contexts),
        context_texts=item.contexts if item.contexts else None,
        question_images=remote_question_images if remote_question_images else None,
        background_url=remote_bg if remote_bg else None,
        resolution=item.resolution or "1334*750",
    )
    opt_positions = layout["option_positions"]
    txt_positions = layout["text_positions"]
    text_font_size = layout["font_size"]
    image_positions = layout["image_positions"]
    background_url = layout["background_url"]

    # 构建选项
    # 选项图片映射表 {index: url} — 只包含远程可访问的 URL
    opt_img_map = {}
    for oi in remote_option_images:
        idx = oi.get("index", -1)
        url = oi.get("url", "")
        if idx >= 0 and url:
            opt_img_map[idx] = url

    options: List[Dict[str, Any]] = []
    for i, opt in enumerate(item.options):
        pos = opt_positions[i]
        is_answer = opt.label.upper() == item.answer.upper()
        has_img = i in opt_img_map
        # ★★★ 关键修复：图片选项绝对不能带 rgb 字段！
        # 对比能正常预览的题目 26018：图片选项结构只有 5 个字段，完全没有 rgb
        # 我们之前错误地发 rgb="#" → 莫顿渲染器解析异常 → "加载配置失败"
        if has_img:
            # 图片选项（对齐 26018）：{webOptionUrl, answer, isAppear, coordinate, scaleValue}
            opt_payload: Dict[str, Any] = {
                "webOptionUrl": opt_img_map[i],
                "answer": is_answer,
                "isAppear": True,
                "coordinate": {"x": pos["x"], "y": pos["y"]},
                "scaleValue": 100,
            }
            print(f"  [OPT_IMG] 选项{opt.label.upper()} has_img=True(无rgb), webOptionUrl={opt_img_map[i][:60]}...")
        else:
            # 文字选项：保留 rgb + content + fontSize
            opt_payload: Dict[str, Any] = {
                "answer": is_answer,
                "rgb": RGB,
                "isAppear": True,
                "coordinate": {"x": pos["x"], "y": pos["y"]},
                "scaleValue": 100,
                "content": opt.content,
                "fontSize": FONT_SIZE,
            }
        options.append(opt_payload)

    # 画布文本 contexts — x 固定在绿框左边缘，displayBoxWeight 占满可用宽度
    # ★ 规则：题目有背景图(远程)时不再添加文本（背景图已含内容），contexts 留空
    contexts: List[Dict[str, Any]] = []
    if has_bg:
        print(f"[BG_NOTEXT] 题目[{item.name or '?'}] 有背景图(远程)，按规则清除题干文本与contexts，不再添加文本")
    else:
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

    # ★ displayBox 只由"题干本身是否纯图片"决定，与"选项图"无关；
    #   但若题目有「远程背景图」，同样关闭文字框（背景图已含内容，不该再叠文字）
    stem_text_present = bool(item.stem) or bool(contexts)
    stem_is_pure_image = bool(remote_question_images) and not stem_text_present
    display_box = not (stem_is_pure_image or has_bg)

    # 图文互斥：题干本身是远程图片(无文字) 或 题目有远程背景图 时，清空题干文字
    stem_content = "" if (remote_question_images or has_bg) else item.stem

    question_data: Dict[str, Any] = {
        "answerFeedbackType": "ALL",
        "skinVersion": 1,
        "resolution": res.replace("高清", "").strip(),
        "questionItem": {
            "displayNumber": False,
            # ★ 关键修复：displayBox 仅由"题干是否纯图片"决定
            #   题干有文字 → True（显示文字框）；题干纯图片 → False（显示图片）
            "displayBox": display_box,
            "newDisplayBox": display_box,
            "options": options,
        },
        "questionType": QUESTION_TYPE_CHOICE,
        "name": item.name,
        # 对齐正常题目 26249：stem 必须带 webAudioUrl 字段（即使为 null）
        # ★ 题干音频（朗读/配音）：有则设 webAudioUrl，文字 content 保留（行为B：两者共存，音频为附加）
        "stem": {"webAudioUrl": item.stemAudio or None, "content": stem_content},
        "subjectId": item.subjectId or subject_id,
        # 注意：正常题目25867 无 autoWidth/autoHeight 顶层字段，移除以完全对齐
    }

    # ★★★ 反馈构造：有语音用 AUDIO(纯语音, webUrl=mp3)，无语音回退文字(PICTURE_CONTENT) ★★★
    # 莫顿规则（实测 26018）：反馈含音频时必须 answerType="AUDIO" 且音频放 webUrl，
    # 不能 PICTURE_CONTENT 里再挂音频（非法混搭 → "加载配置失败"）。
    # 组课文档规范：
    #   rightQuestionFeedback → 答对时统一展示的通用正反馈（音频或文字）
    #   判定1 (variableValue="1,1")   → 第1次错 → 负反馈1
    #   判定2 (variableValue="2,2")   → 第2次错 → 负反馈2
    #   判定3 (variableValue="3,3")   → 第3次错 → 负反馈3

    # 收集各次错误对应的音频 URL（来自 item.conditionFeedbacks 的 audioUrl）
    cond_audios = ["", "", ""]
    for cf in (item.conditionFeedbacks or []):
        n = cf.minWrong
        if 1 <= n <= 3:
            cond_audios[n - 1] = cf.audioUrl

    # ── 答对通用反馈（正反馈独立存放）──
    # ★ 文档优先：correctFeedback 写"无"/空 → 不生成 rightQuestionFeedback（莫顿允许缺失）
    # 有语音→AUDIO，有文字→PICTURE_CONTENT，都没写"无"→不生成
    right_audio = item.correctFeedbackAudio
    right_text = (getattr(item, 'correctFeedback', '') or '').strip()
    if not _is_positive_feedback_none(right_text):
        if right_audio:
            question_data["rightQuestionFeedback"] = _audio_fb(right_audio)
        elif right_text:
            question_data["rightQuestionFeedback"] = {
                "answerType": "PICTURE_CONTENT",
                "content": right_text,
            }
        else:
            # 文档没填正反馈字段（非显式"无"）→ 保留兜底文字（向后兼容）
            question_data["rightQuestionFeedback"] = {
                "answerType": "PICTURE_CONTENT",
                "content": _FALLBACK_RIGHT,
            }
    else:
        print(f"[RIGHT_FB] ⏭️ 跳过答对反馈(文档标记'无'): name={item.name!r}, correctFeedback={right_text!r}")

    # ── 非通用反馈（按条件反馈逐条生成，不再强制3条）──
    # 每条 ConditionFeedback → 1 个 questionFeedback，用其自身的 minWrong/maxWrong
    # 若无 conditionFeedbacks → 空 list（莫顿允许无判定反馈）
    cond_fbs_raw = item.conditionFeedbacks or []
    built_feedbacks = []
    for cf in cond_fbs_raw:
        n = cf.minWrong
        audio_url = getattr(cf, 'audioUrl', '') or ''
        fb_text = (getattr(cf, 'content', '') or '').strip()
        # ★ 清洗 content 中的"负反馈N："/"第N次错误"等标签前缀
        fb_text_clean = _clean_feedback_label(fb_text)
        vv = f"{n},{n}" if n == getattr(cf, 'maxWrong', n) else f"{n},{getattr(cf, 'maxWrong', n)}"
        if audio_url:
            fb_entry = _audio_fb(audio_url, vv)
        else:
            fb_entry = {
                "webUrl": None,
                "webAudioUrl": None,
                "webListenList": [{
                    "variableKey": "WrongNumber",
                    "variableType": "int",
                    "compare": "EQUAL",
                    "value": n,
                    "variableValue": vv,
                }],
                "logicTypeList": [],
                "logicExpressList": [],
                "answerType": "PICTURE_CONTENT",
                "content": fb_text_clean or _FALLBACK_WRONG[min(n-1, len(_FALLBACK_WRONG)-1)],
            }
        built_feedbacks.append(fb_entry)

    question_data["questionFeedbacks"] = built_feedbacks
    # 工具配置：正常题目结构必需（即使为空也是 []）
    question_data["toolConfigList"] = []
    # 知识点：保证始终为数组（正常题即使为空也是 []，缺失会导致结构校验失败）
    if "knowledge" not in question_data:
        question_data["knowledge"] = []

    # ★ 诊断：打印题干和背景的最终状态（过滤后）
    has_remote_qimg = bool(remote_question_images)
    has_bg = bool(remote_bg)
    print(f"[STEM_BG_DEBUG] 远程题干图={has_remote_qimg}(原{len(original_qimgs)}个), "
          f"stem_text_present={stem_text_present}, displayBox={display_box}; "
          f"远程背景图={has_bg}, bg_url='{(remote_bg or '')[:60]}'")

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

    # 画布背景 — 只写入远程可达 URL（本地路径会导致渲染器加载失败）
    if remote_bg:
        question_data["webBgmUrl"] = remote_bg
        print(f"[BG_INSERT] 背景图(远程)已在insert阶段写入: {remote_bg[:60]}...")
    elif item.backgroundUrl and not remote_bg:
        print(f"[BG_INSERT] ⚠️ 背景图是本地路径，已跳过: {item.backgroundUrl[:60]}...")

    # 诊断：打印最终 payload 关键结构
    final_opts = question_data.get("questionItem", {}).get("options", [])
    for idx, o in enumerate(final_opts):
        has_url = o.get("webOptionUrl", "")
        has_content = o.get("content", "")
        print(f"[FINAL_OPT] 选项{idx}: webOptionUrl={'有' if has_url else '无'}, content={'有' if has_content else '无'}, rgb={o.get('rgb')}, coord={o.get('coordinate')}")
    print(f"[FINAL_PIC] pictures数={len(question_data.get('pictures', []))}, webBgmUrl={'有' if question_data.get('webBgmUrl') else '无'}, displayBox={question_data.get('questionItem', {}).get('displayBox')}")

    # ★ 关键诊断：确认 questionFeedbacks + rightQuestionFeedback 结构正确
    fbs = question_data.get("questionFeedbacks", [])
    fb_info = [f.get("webListenList", [{}])[0].get("variableValue", "❌缺失") for f in fbs]
    rfq = question_data.get("rightQuestionFeedback", {})
    print(f"[FB_DEBUG] rightQuestionFeedback={'有' if rfq else '无'}(正反馈分离), "
          f"questionFeedbacks数={len(fbs)}, 各判定variableValue={fb_info}")

    # ★ AI辅助提示 → 自动提示（全题型通用，不写死）
    # 真实字段结构已对齐莫顿样例：questionData.auxiliaryTips.autoTip.list
    hints = getattr(item, "autoHints", None)
    if hints:
        entries = _build_auto_hint_entries(hints)
        if entries:
            question_data["auxiliaryTips"] = _build_auxiliary_tips(entries)
            print(f"[AUTO_HINT] ✅ 写入 {len(entries)} 条自动提示到 auxiliaryTips.autoTip")

    # ★ 完整 JSON 输出 — 如果仍然"加载配置失败"，请把这段 JSON 发给 AI 做逐字段对比
    print(f"[JSON_DUMP] ════════ 完整 questionData ════════")
    print(json.dumps(question_data, ensure_ascii=False, indent=2))
    print(f"[JSON_DUMP] ════════ END ════════")

    return {"questionData": question_data}, question_data


def _calculate_blank_smart_layout(
    stem_text: str,
    blank_count: int,
    resolution: str = "1334*750",
) -> dict:
    """填空题智能布局引擎 — 根据题目文本和空位数自动计算坐标。

    返回:
        {
            "text_position": {"x": int, "y": int, "width": int, "font_size": int},
            "blank_positions": [{"x": int, "y": int, "width": int, "height": int}, ...],
        }
    """
    # ── 解析分辨率 → 画布尺寸 ──
    if "*" in resolution:
        parts = resolution.replace("高清", "").strip().split("*")
        CW = int(parts[0]) if parts[0].isdigit() else 1334
        CH = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 750
    else:
        CW, CH = 1334, 750

    # ── 按比例缩放绿框 ──
    # 基准: 1334*750 → 绿框 x=182~1152 (宽970), y=130~580 (高450)
    scale_x = CW / 1334.0
    scale_y = CH / 750.0
    scale = min(scale_x, scale_y)

    CONTENT_LEFT = int(182 * scale_x)
    CONTENT_RIGHT = int(1152 * scale_x)
    CONTENT_TOP = int(130 * scale_y)
    CONTENT_BOTTOM = int(580 * scale_y)
    CONTENT_W = CONTENT_RIGHT - CONTENT_LEFT
    CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP
    CENTER_X = (CONTENT_LEFT + CONTENT_RIGHT) // 2

    PAD = int(20 * scale)

    # ── 字号：根据文本长度自适应 ──
    text_len = len(stem_text) if stem_text else 0
    if text_len <= 15:
        FONT_SIZE = int(50 * scale)
    elif text_len <= 40:
        FONT_SIZE = int(45 * scale)
    else:
        FONT_SIZE = max(int(40 * scale), 24)

    # ── 文本测量 ──
    text_zone_w = CONTENT_W - PAD * 2
    m = _measure_paragraph(stem_text, FONT_SIZE, text_zone_w)
    text_h = m["total_height"]

    # ── 文本位置：绿框顶部居中 ──
    text_y = CONTENT_TOP + PAD
    text_position = {
        "x": CONTENT_LEFT + PAD,
        "y": text_y,
        "width": text_zone_w,
        "font_size": FONT_SIZE,
    }

    # ── 空位尺寸 ──
    BLANK_W = max(int(146 * scale), 100)
    BLANK_H = max(int(63 * scale), 48)
    BLANK_GAP = max(int(40 * scale), 20)

    # ── 空位区域 ──
    text_bottom = text_y + text_h
    blank_area_top = text_bottom + PAD
    blank_area_bottom = CONTENT_BOTTOM - PAD
    blank_area_h = blank_area_bottom - blank_area_top

    # 如果空位区域不够，向上挤压（减少文本底部间距）
    if blank_area_h < BLANK_H and text_bottom > CONTENT_TOP + PAD * 2:
        blank_area_top = max(CONTENT_TOP + PAD, CONTENT_BOTTOM - PAD - BLANK_H)
        blank_area_h = blank_area_bottom - blank_area_top

    # ── 空位横向排列：居中 ──
    total_blank_w = blank_count * BLANK_W + (blank_count - 1) * BLANK_GAP
    if total_blank_w > CONTENT_W - PAD * 2:
        # 超出则压缩
        available = CONTENT_W - PAD * 2
        BLANK_W = max(80, (available - (blank_count - 1) * BLANK_GAP) // blank_count)
        total_blank_w = blank_count * BLANK_W + (blank_count - 1) * BLANK_GAP

    blank_start_x = CENTER_X - total_blank_w // 2
    blank_y = blank_area_top + (blank_area_h - BLANK_H) // 2  # 垂直居中

    blank_positions = []
    for i in range(blank_count):
        blank_positions.append({
            "x": blank_start_x + i * (BLANK_W + BLANK_GAP),
            "y": blank_y,
            "width": BLANK_W,
            "height": BLANK_H,
        })

    return {
        "text_position": text_position,
        "blank_positions": blank_positions,
    }


def _build_completion_payload(item: BlankQuestionItem, subject_id: int) -> Dict[str, Any]:
    """构建填空题 COMPLETION 接口的完整 payload（智能布局引擎）。"""
    FONT_SIZE = 32
    RGB = "#9f765b"

    res = item.resolution or "1334*750"
    if "*" in res:
        parts = res.replace("高清", "").strip().split("*")
        CANVAS_WIDTH = int(parts[0]) if parts[0].isdigit() else 1334
        CANVAS_HEIGHT = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 750
    else:
        CANVAS_WIDTH, CANVAS_HEIGHT = 1334, 750

    # ★ 智能布局引擎：根据文本长度 + 空位数 + 分辨率自动计算坐标
    # ★ 规则：题目有背景图(远程)时不再添加文本（背景图已含内容）
    has_bg_comp = bool(item.backgroundUrl and _is_remote_url(item.backgroundUrl))
    if has_bg_comp:
        stem_display = ""
        context_display = ""
        print(f"[BG_NOTEXT] 填空题[{item.name or '?'}] 有背景图(远程)，按规则清除题干文本与contexts，不再添加文本")
    else:
        stem_display = item.stem.replace("______", "")
        context_source = item.contexts[0] if item.contexts else item.stem
        context_display = context_source.replace("______", "")
    blank_count = len(item.blanks) if item.blanks else 1

    layout = _calculate_blank_smart_layout(
        stem_text=context_display,
        blank_count=blank_count,
        resolution=res,
    )
    tp = layout["text_position"]
    bp = layout["blank_positions"]

    contexts = [] if has_bg_comp else [{
        "content": context_display,
        "size": tp["font_size"],
        "rgb": RGB,
        "coordinate": {"x": tp["x"], "y": tp["y"]},
        "displayBoxWeight": tp["width"],
        "displayBoxHeight": int(tp["font_size"] * _LH_RATIO * 2),
        "isDrag": True,
        "uuid": str(uuid_mod.uuid4()),
    }]

    blanks: List[Dict[str, Any]] = []
    all_blank_options: List[Dict[str, Any]] = []  # ★ 汇总所有空位选项到 questionItem.options

    for idx, blank in enumerate(item.blanks or []):
        answers = [blank.answer] if blank.answer else []
        if blank.aliases:
            answers.extend([a for a in blank.aliases if a])

        # ★ 使用智能布局计算的坐标
        pos = bp[idx] if idx < len(bp) else {"x": 678 + idx * 168, "y": 171}
        is_option_type = _normalize_blank_value(blank.answerType, {"内容": "CONTENT", "选项": "OPTION"}, "CONTENT") == "OPTION"

        # ── 构建选项列表（含完整元数据 — 参照 CHOICE 选项结构） ──
        options = []
        if is_option_type and blank.options:
            num_opts = len(blank.options)
            opt_gap = 8
            opt_w = max(60, (pos["width"] - opt_gap * (num_opts - 1)) // num_opts)
            opt_h = pos["height"]
            total_opt_w = num_opts * opt_w + (num_opts - 1) * opt_gap
            opt_start_x = pos["x"] + (pos["width"] - total_opt_w) // 2

            for opt_idx, opt_content in enumerate(blank.options):
                opt_x = opt_start_x + opt_idx * (opt_w + opt_gap)
                is_answer = opt_content in answers
                # ★ 填空选择题选项图：若该选项有图片，用 webOptionUrl（对齐选择题图片选项，禁带 rgb/content）
                opt_img_map = {oi.get("index"): oi.get("url") for oi in (blank.optionImages or []) if oi.get("url")}
                img_url = opt_img_map.get(opt_idx)
                if img_url:
                    opt_payload = {
                        "uuid": str(uuid_mod.uuid4()),
                        "answer": is_answer,
                        "webOptionUrl": img_url,
                        "isAppear": True,
                        "coordinate": {"x": opt_x, "y": pos["y"]},
                        "scaleValue": 100,
                        "tmpStr": f"选项 {opt_idx + 1}",
                        "optionIndex": str(opt_idx + 1),
                        "isDrag": False,
                    }
                    print(f"  [COMPLETION_OPT_IMG] 选项型空[{blank.varName}] 选项{opt_idx + 1} 使用图片 {img_url[:50]}...")
                else:
                    opt_payload = {
                        "uuid": str(uuid_mod.uuid4()),
                        "answer": is_answer,
                        "content": opt_content,
                        "fontSize": FONT_SIZE,
                        "rgb": RGB,
                        "displayBoxWeight": opt_w,
                        "displayBoxHeight": opt_h,
                        "coordinate": {"x": opt_x, "y": pos["y"]},
                        "isAppear": True,
                        "scaleValue": 100,
                        "tmpStr": f"选项 {opt_idx + 1}",
                        "optionIndex": str(opt_idx + 1),
                        "isDrag": False,
                    }
                options.append(opt_payload)
                # ★ 同时加入 questionItem.options（莫顿画布渲染依赖此字段）
                all_blank_options.append(opt_payload)

        blanks.append({
            "uuid": str(uuid_mod.uuid4()),
            "blankType": "PICTURE",
            "answerType": "OPTION" if is_option_type else "CONTENT",
            "blankAnswer": answers,
            "coordinate": {"x": pos["x"], "y": pos["y"]},
            "displayFormat": "OPTIONNUMBER",
            "fontSize": FONT_SIZE,
            "frontType": "blank",
            "height": pos["height"],
            "rgb": RGB,
            "styleType": 1,
            "top": False,
            "weight": pos["width"],
            "keyBoardType": _normalize_blank_value(blank.keyboardStyle, {"横向双线键盘": "HORIZONTAL_DOUBLE_LINE"}, "HORIZONTAL_DOUBLE_LINE"),
            "options": options,
            "questionType": QUESTION_TYPE_COMPLETION,
            "action": "edit",
        })

    question_item = {
        "displayNumber": any(blank.showNumber for blank in item.blanks),
        "displayBox": not has_bg_comp,  # ★ 规则：有背景图(远程)时关闭文字框
        "newDisplayBox": True,
        "inputType": 1,
        "keyBoardType": blanks[0]["keyBoardType"] if blanks else "HORIZONTAL_DOUBLE_LINE",
        "completionBlankList": blanks,
        "options": all_blank_options,  # ★ 空位选项（含坐标/字号/颜色等完整元数据）
    }

    display_stem = stem_display
    question_data: Dict[str, Any] = {
        "answerFeedbackType": "ALL",
        "skinVersion": 1,
        "resolution": res.replace("高清", "").strip(),
        "questionItem": question_item,
        "questionType": QUESTION_TYPE_COMPLETION,
        "autoWidth": CANVAS_WIDTH,
        "autoHeight": CANVAS_HEIGHT,
        "name": item.name,
        # ★ 题干音频（朗读/配音）：有则设 webAudioUrl，文字 content 保留（行为B：两者共存，音频为附加）
        "stem": {"content": display_stem, "webAudioUrl": item.stemAudio or None},
        "subjectId": item.subjectId or subject_id,
        "contexts": contexts,
        "questionFeedbacks": [],
        "auxiliaryTips": {"tips": [], "autoTip": {}, "parses": [], "experiments": []},
    }

    if item.knowledge:
        question_data["knowledge"] = [k.strip() for k in item.knowledge.split(";") if k.strip()]
    if item.difficulty:
        question_data["difficulty"] = item.difficulty
    if item.grade:
        question_data["grade"] = item.grade
    if item.stage:
        question_data["stage"] = item.stage

    # ★ 填空题反馈/语音/背景：复用 _merge_update_fields，把 item 上的
    #   conditionFeedbacks / correctFeedbackAudio / backgroundUrl / questionImages
    #   合并进 question_data（与选择题一致）。
    #   前端按题名匹配的图片、全局语音(defaultVoiceFeedback)、全局背景(defaultBackgroundUrl)
    #   都已由 _apply_defaults 注入到 item，这里统一消费，避免填空题"识别不出来"反馈/语音/图片。
    _merge_update_fields(question_data, item)

    print(f"[COMPLETION] name={item.name} subjectId={question_data['subjectId']} resolution={question_data['resolution']} blanks={len(blanks)}")

    # ★ AI辅助提示 → 自动提示（全题型通用，不写死）
    # 真实字段：questionData.auxiliaryTips.autoTip.list（与选择题完全一致）
    hints = getattr(item, "autoHints", None)
    if hints:
        entries = _build_auto_hint_entries(hints)
        if entries:
            # 覆盖下方初始化的空 auxiliaryTips 骨架
            question_data["auxiliaryTips"] = _build_auxiliary_tips(entries)
            print(f"[COMPLETION_AUTO_HINT] ✅ 写入 {len(entries)} 条自动提示")

    return {"questionData": question_data}, question_data


@router.post("/preview-payload")
async def preview_payload(body: PreviewPayloadRequest) -> Dict[str, Any]:
    """构建创建前预览用的最终 questionData；不会在莫顿系统创建题目。"""
    try:
        item = _to_model(body.question)
        _apply_defaults(item, body)
        if getattr(item, 'questionType', QUESTION_TYPE_CHOICE) == QUESTION_TYPE_COMPLETION:
            _, question_data = _build_completion_payload(item, body.subjectId)
        else:
            _, question_data = _build_choice_payload(item, body.subjectId)
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
        item = _to_model(body.question)
        _apply_defaults(item, body)

        original_name = item.name or body.defaultName or "新建选择题"
        if not original_name.startswith(PREVIEW_NAME_PREFIX):
            item.name = f"{PREVIEW_NAME_PREFIX}{original_name}"

        if getattr(item, 'questionType', QUESTION_TYPE_CHOICE) == QUESTION_TYPE_COMPLETION:
            payload, question_data = _build_completion_payload(item, body.subjectId)
        else:
            payload, question_data = _build_choice_payload(item, body.subjectId)
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
    body.questions = [_to_model(q) for q in body.questions]
    for q in body.questions:
        _apply_defaults(q, body)

    try:
        for idx, item in enumerate(body.questions):
            try:
                if getattr(item, 'questionType', QUESTION_TYPE_CHOICE) == QUESTION_TYPE_COMPLETION:
                    payload, question_data = _build_completion_payload(item, body.subjectId)
                else:
                    payload, question_data = _build_choice_payload(item, body.subjectId)
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
        body.questions = [_to_model(q) for q in body.questions]
        for q in body.questions:
            _apply_defaults(q, body)

        success_count = 0
        fail_count = 0

        try:
            for idx, item in enumerate(body.questions):
                try:
                    if getattr(item, 'questionType', QUESTION_TYPE_CHOICE) == QUESTION_TYPE_COMPLETION:
                        payload, question_data = _build_completion_payload(item, body.subjectId)
                    else:
                        payload, question_data = _build_choice_payload(item, body.subjectId)
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
                    evt = {
                        "index": idx, "success": True, "questionId": question_id,
                        "progress": progress_str, "percent": percent,
                        "feedbackError": update_error or None,
                        "warning": f"⚠️ 题目已创建但元数据/反馈更新失败: {update_error}" if update_error else None,
                    }
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"

                except MortonClientError as e:
                    fail_count += 1
                    percent = round((idx + 1) / total * 100)
                    evt = {"index": idx, "success": False, "error": e.message, "progress": f"{idx+1}/{total}", "percent": percent}
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"

                except Exception as e:
                    # ★ 捕获所有异常，避免一道题崩溃导致整个批次中断
                    fail_count += 1
                    percent = round((idx + 1) / total * 100)
                    error_msg = f"{type(e).__name__}: {e}"
                    print(f"[ERROR] 第{idx+1}题异常: {error_msg}")
                    evt = {"index": idx, "success": False, "error": error_msg, "progress": f"{idx+1}/{total}", "percent": percent}
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"

                if idx < total - 1:
                    await asyncio.sleep(0.3)

            # 完成事件
            evt = {"done": True, "total": total, "successCount": success_count, "failCount": fail_count}
            yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"

        except MortonClientError as e:
            evt = {"done": True, "total": total, "successCount": success_count, "failCount": fail_count, "fatalError": e.message}
            yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
        except Exception as e:
            # ★ 兜底：任何未预期的异常都发送 done 事件，避免前端进度条永远转圈
            evt = {"done": True, "total": total, "successCount": success_count, "failCount": fail_count, "fatalError": f"{type(e).__name__}: {e}"}
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


async def parse_blank_excel(file: UploadFile) -> Dict[str, Any]:
    """解析填空题 Excel（精简结构版）。支持空1/空2/... 多空模板。"""
    try:
        content = await file.read()
        wb = load_workbook(io.BytesIO(content), read_only=True)
        ws = wb.active
        header = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        questions: List[Dict[str, Any]] = []

        def get(row, idx, default=""):
            if idx >= 0 and idx < len(row) and row[idx] is not None:
                return row[idx]
            return default

        def find_col(*names):
            for name in names:
                if name in header:
                    return header.index(name)
            return -1

        idx_question_type = find_col("题型")
        idx_name = find_col("名称")
        idx_code = find_col("标准编码", "素材编码", "题号编码", "编码")
        idx_subject = find_col("学科")
        idx_resolution = find_col("分辨率")
        idx_grade = find_col("年级")
        idx_stage = find_col("阶段")
        idx_difficulty = find_col("难度")
        idx_knowledge = find_col("知识点")
        idx_stem_text = find_col("题干文本")
        idx_context_text = find_col("文本")
        idx_stem_audio = find_col("题干音频")

        # ★ 反馈/语音/图片列（与选择题解析对齐；填空题 Excel 若含这些列则读取，缺则留空）
        idx_wrong_fb = find_col("答错通用反馈")
        idx_correct_fb = find_col("答对通用反馈")
        idx_wrong_fb_img = find_col("答错通用反馈图片")
        idx_correct_fb_img = find_col("答对通用反馈图片")
        idx_wrong_fb_audio = find_col("答错通用反馈音频")
        idx_correct_fb_audio = find_col("正反馈音频")
        idx_q_imgs = find_col("题目图片")
        idx_bg = find_col("背景图")
        idx_cond_min = find_col("最少错次")
        idx_cond_max = find_col("最多错次")
        idx_cond_text = find_col("条件反馈文本")
        idx_cond_img = find_col("条件反馈图片")
        idx_cond_audio = find_col("条件反馈音频")

        # ★ 自动提示（小灯泡）：最多 3 个提示，每个 图片/音频 + 变量1~6（列存在才读取；旧 Excel 无此列则 -1 优雅跳过）
        #   第 1 个提示兼容旧单提示 Excel 的无序号表头（自动提示图片 / 自动提示音频 / 自动提示变量N）
        # 第 1 个提示兼容旧单提示 Excel 的无序号表头（自动提示图片 / 自动提示音频 / 自动提示变量N）
        def _auto_col_names(h: int, base: str):
            names = [f"自动提示{h}{base}"]
            if h == 1:
                if base == "图片":
                    names.append("自动提示图片")
                elif base == "音频":
                    names.append("自动提示音频")
                elif base.startswith("变量"):
                    names.append(f"自动提示{base}")
            return tuple(names)

        idx_auto_img = [find_col(*_auto_col_names(h, "图片")) for h in range(1, AUTO_HINT_COUNT + 1)]
        idx_auto_audio = [find_col(*_auto_col_names(h, "音频")) for h in range(1, AUTO_HINT_COUNT + 1)]
        idx_auto_vars = [[find_col(*_auto_col_names(h, f"变量{i}")) for i in range(1, AUTO_HINT_VAR_COUNT + 1)]
                         for h in range(1, AUTO_HINT_COUNT + 1)]

        # 多空按空编号扫描
        blank_groups = []
        for n in range(1, 10):
            prefix = f"空{n}"
            cols = {
                "name": find_col(f"{prefix}名称"),
                "type": find_col(f"{prefix}类型"),
                "answer_type": find_col(f"{prefix}答案类型"),
                "input_style": find_col(f"{prefix}输入框样式"),
                "top": find_col(f"{prefix}是否置顶"),
                "x": find_col(f"{prefix}坐标X"),
                "y": find_col(f"{prefix}坐标Y"),
                "width": find_col(f"{prefix}宽度"),
                "height": find_col(f"{prefix}高度"),
                "answer": find_col(f"{prefix}答案"),
                "aliases": find_col(f"{prefix}答案别名"),
                "keyboard": find_col(f"{prefix}键盘样式"),
                "show_number": find_col(f"{prefix}显示序号"),
                "show_box": find_col(f"{prefix}显示选项底框"),
                "opt1": find_col(f"{prefix}选项1"),
                "opt2": find_col(f"{prefix}选项2"),
                "opt3": find_col(f"{prefix}选项3"),
                "opt4": find_col(f"{prefix}选项4"),
            }
            if cols["name"] >= 0 or cols["answer"] >= 0:
                blank_groups.append((n, cols))

        for row in rows:
            qtype = str(get(row, idx_question_type, "COMPLETION")).strip().upper() if idx_question_type >= 0 else "COMPLETION"
            if qtype and qtype not in ("COMPLETION", "BLANK"):
                continue
            name = str(get(row, idx_name, "")).strip()
            code = str(get(row, idx_code, "")).strip()
            subject_raw = str(get(row, idx_subject, "")).strip()
            resolution = str(get(row, idx_resolution, "1334*750")).strip()
            grade_raw = str(get(row, idx_grade, "")).strip().replace("－", "-").replace("—", "-")
            stage_raw = str(get(row, idx_stage, "")).strip()
            difficulty_raw = str(get(row, idx_difficulty, "0")).strip() or "0"
            knowledge = str(get(row, idx_knowledge, "")).strip()
            stem_text = str(get(row, idx_stem_text, "")).strip()
            context_text = str(get(row, idx_context_text, stem_text)).strip()
            stem_audio = str(get(row, idx_stem_audio, "")).strip()

            # ★ 自动提示（小灯泡）：逐提示读取图片/音频/变量1~6，构建最多 3 条提示（空提示跳过）
            auto_hint_specs = []
            for _h in range(AUTO_HINT_COUNT):
                _ai = idx_auto_img[_h]
                _aa = idx_auto_audio[_h]
                _img = str(get(row, _ai, "")).strip() if _ai >= 0 else ""
                _aud = str(get(row, _aa, "")).strip() if _aa >= 0 else ""
                _vars = [str(get(row, c, "")).strip() if c >= 0 else "" for c in idx_auto_vars[_h]]
                auto_hint_specs.append((_img, _aud, _vars))
            auto_hints = _parse_auto_hint_columns(auto_hint_specs)

            # ★ 反馈/语音/图片（列存在才读取；与选择题解析对齐）
            wrong_fb = str(get(row, idx_wrong_fb, "")).strip()
            correct_fb = str(get(row, idx_correct_fb, "")).strip()
            wrong_fb_img = str(get(row, idx_wrong_fb_img, "")).strip()
            correct_fb_img = str(get(row, idx_correct_fb_img, "")).strip()
            wrong_fb_audio = str(get(row, idx_wrong_fb_audio, "")).strip()
            correct_fb_audio = str(get(row, idx_correct_fb_audio, "")).strip()
            q_imgs_raw = str(get(row, idx_q_imgs, "")).strip()
            bg_url = str(get(row, idx_bg, "")).strip()
            cond_min_raw = str(get(row, idx_cond_min, "")).strip()
            cond_max_raw = str(get(row, idx_cond_max, "")).strip()
            cond_text = str(get(row, idx_cond_text, "")).strip()
            cond_img = str(get(row, idx_cond_img, "")).strip()
            cond_audio = str(get(row, idx_cond_audio, "")).strip()

            if subject_raw.isdigit():
                subject_id = int(subject_raw)
            else:
                subject_id = SUBJECT_MAP.get(subject_raw, 0)
            if grade_raw.isdigit():
                grade_val = int(grade_raw)
            else:
                grade_val = GRADE_MAP.get(grade_raw, 0)
            if stage_raw.isdigit():
                stage_val = int(stage_raw)
            else:
                stage_val = STAGE_MAP.get(stage_raw, 0)
            if difficulty_raw.isdigit():
                difficulty = int(difficulty_raw)
            else:
                difficulty = DIFFICULTY_MAP.get(difficulty_raw, 0)
            difficulty = max(0, min(5, difficulty))

            blanks = []
            content_markup = stem_text
            for n, cols in blank_groups:
                blank_name = str(get(row, cols["name"], f"空{n}")).strip()
                blank_type = _normalize_blank_value(str(get(row, cols["type"], "内容型")), {"内容型": "CONTENT", "选项型": "OPTION"}, "CONTENT")
                answer_type = _normalize_blank_value(str(get(row, cols["answer_type"], "内容")), {"内容": "CONTENT", "选项": "OPTION"}, "CONTENT")
                input_style = _normalize_blank_value(str(get(row, cols["input_style"], "文本输入框")), {"文本输入框": "TEXT", "选项输入框": "OPTION"}, "TEXT")
                top = str(get(row, cols["top"], "0")).strip().lower() in ("1", "true", "yes")
                x = int(get(row, cols["x"], 0) or 0)
                y = int(get(row, cols["y"], 0) or 0)
                width = int(get(row, cols["width"], 0) or 0)
                height = int(get(row, cols["height"], 0) or 0)
                answer = str(get(row, cols["answer"], "")).strip()
                aliases = [a.strip() for a in str(get(row, cols["aliases"], "")).split(";") if a.strip()]
                keyboard = _normalize_blank_value(str(get(row, cols["keyboard"], "横向双线键盘")), {"横向双线键盘": "HORIZONTAL_DOUBLE_LINE"}, "HORIZONTAL_DOUBLE_LINE")
                show_number = str(get(row, cols["show_number"], "是")).strip().lower() not in ("0", "false", "no", "否")
                show_box = str(get(row, cols["show_box"], "否")).strip().lower() in ("1", "true", "yes", "是")
                options = []
                for ci in ("opt1", "opt2", "opt3", "opt4"):
                    v = str(get(row, cols[ci], "")).strip()
                    if v:
                        options.append(v)
                if not answer and not options:
                    continue
                blanks.append({
                    "varName": blank_name or f"空{n}",
                    "blankType": blank_type,
                    "answerType": answer_type,
                    "inputStyle": input_style,
                    "top": top,
                    "x": x,
                    "y": y,
                    "width": width,
                    "height": height,
                    "answer": answer,
                    "aliases": aliases,
                    "keyboardStyle": keyboard,
                    "showNumber": show_number,
                    "showBox": show_box,
                    "options": options,
                })
                if blank_name and "______" in content_markup and len(blanks) == 1:
                    content_markup = content_markup.replace("______", f"{{{{{blank_name}}}}}", 1)

            if not stem_text or not blanks:
                continue

            # ★ 条件反馈：最少/最多错次 + 条件文本。以"负反馈N"为准做 N-N 单点归一，
            #   并在解析阶段即清洗"负反馈N："前缀（与选择题一致）。
            cond_min_wrong = 0
            cond_max_wrong = 0
            if cond_min_raw and cond_min_raw.replace("-", "").isdigit():
                cond_min_wrong = max(0, int(cond_min_raw))
            if cond_max_raw and cond_max_raw.replace("-", "").isdigit():
                cond_max_wrong = max(0, int(cond_max_raw))
            _neg_m = re.search(r'负反馈\s*(\d+)', cond_text)
            if _neg_m:
                _neg_n = int(_neg_m.group(1))
                cond_min_wrong = _neg_n
                cond_max_wrong = _neg_n

            condition_feedbacks = []
            if cond_text.strip():
                condition_feedbacks.append({
                    "minWrong": cond_min_wrong,
                    "maxWrong": cond_max_wrong,
                    "content": _clean_feedback_label(cond_text),
                    "imageUrl": cond_img,
                    "audioUrl": cond_audio,
                })

            # ★ 题目图片：分号分隔 → [{url, width, height}]
            question_images = []
            if q_imgs_raw:
                for url in q_imgs_raw.split(";"):
                    url = url.strip()
                    if url:
                        question_images.append({"url": url, "width": 0, "height": 0})

            questions.append({
                "stem": stem_text,
                "name": name,
                "code": code,
                "subjectId": subject_id,
                "resolution": resolution,
                "grade": grade_val,
                "stage": stage_val,
                "difficulty": difficulty,
                "knowledge": knowledge,
                "questionType": QUESTION_TYPE_COMPLETION,
                "contentMarkup": content_markup,
                "contexts": [context_text] if context_text else [],
                "blanks": blanks,
                "questionImages": question_images,
                "backgroundUrl": bg_url,
                "wrongFeedback": wrong_fb,
                "correctFeedback": correct_fb,
                "wrongFeedbackImg": wrong_fb_img,
                "correctFeedbackImg": correct_fb_img,
                "wrongFeedbackAudio": wrong_fb_audio,
                "correctFeedbackAudio": correct_fb_audio,
                "stemAudio": stem_audio,
                "conditionFeedbacks": condition_feedbacks,
                "autoHints": auto_hints,
            })
        wb.close()
        return {"success": True, "questions": questions, "total": len(questions)}
    except Exception as e:
        return {"success": False, "error": f"填空题解析失败: {e}"}


@router.post("/parse-excel")
async def parse_excel(file: UploadFile = File(...), question_type: str = "CHOICE") -> Dict[str, Any]:
    """解析上传的 Excel 文件，返回解析后的题目数据。"""
    try:
        if question_type.upper() == QUESTION_TYPE_BLANK:
            return await parse_blank_excel(file)
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

        # ★ 题干音频（朗读/配音）：按表头名识别（配音/题干音频/朗读音频 等），与反馈音频独立
        idx_stem_audio = -1
        for _h in ("题干音频", "配音", "朗读音频", "题干配音", "narration"):
            if _h in header:
                idx_stem_audio = header.index(_h)
                break

        # ★ 自动提示（小灯泡）：最多 3 个提示，每个 图片/音频 + 变量1~6（按表头名定位；列缺失则 -1）
        #   第 1 个提示兼容旧单提示 Excel 的无序号表头（自动提示图片 / 自动提示音频 / 自动提示变量N）
        def _find_auto_col(name: str) -> int:
            return header.index(name) if name in header else -1

        def _find_auto_cols(*names: str) -> int:
            """返回首个命中的表头列号，全部未命中返回 -1。"""
            for n in names:
                if n in header:
                    return header.index(n)
            return -1

        def _auto_col_names(h: int, base: str):
            names = [f"自动提示{h}{base}"]
            if h == 1:
                if base == "图片":
                    names.append("自动提示图片")
                elif base == "音频":
                    names.append("自动提示音频")
                elif base.startswith("变量"):
                    names.append(f"自动提示{base}")
            return tuple(names)

        idx_auto_img = [_find_auto_cols(*_auto_col_names(h, "图片")) for h in range(1, AUTO_HINT_COUNT + 1)]
        idx_auto_audio = [_find_auto_cols(*_auto_col_names(h, "音频")) for h in range(1, AUTO_HINT_COUNT + 1)]
        # ★ 标准编码（公司素材编码）：与「名称」互为别名，用于匹配按公司编码命名的素材文件
        idx_code = _find_auto_cols("标准编码", "素材编码", "题号编码", "编码")
        idx_auto_vars = [[_find_auto_cols(*_auto_col_names(h, f"变量{i}")) for i in range(1, AUTO_HINT_VAR_COUNT + 1)]
                         for h in range(1, AUTO_HINT_COUNT + 1)]




        rows = list(ws.iter_rows(min_row=2, values_only=True))
        questions: List[Dict[str, Any]] = []

        for i, row in enumerate(rows):
            if not row or not row[0]:
                continue

            stem = str(row[0] or "").strip()
            contexts_text = str(row[1] or "").strip()
            # ★ 题干音频（朗读/配音）：按表头定位，与反馈音频独立
            stem_audio = (str(row[idx_stem_audio]).strip()
                          if idx_stem_audio >= 0 and idx_stem_audio < len(row) and row[idx_stem_audio] is not None
                          else "")
            # ★ 自动提示（小灯泡）：逐提示读取图片/音频/变量1~6，构建最多 3 条提示（空提示跳过）
            auto_hint_specs = []
            for _h in range(AUTO_HINT_COUNT):
                _ai = idx_auto_img[_h]
                _aa = idx_auto_audio[_h]
                _img = (str(row[_ai]).strip()
                        if _ai >= 0 and _ai < len(row) and row[_ai] is not None else "")
                _aud = (str(row[_aa]).strip()
                        if _aa >= 0 and _aa < len(row) and row[_aa] is not None else "")
                _vars = [
                    str(row[c]).strip() if (c >= 0 and c < len(row) and row[c] is not None) else ""
                    for c in idx_auto_vars[_h]
                ]
                auto_hint_specs.append((_img, _aud, _vars))
            auto_hints = _parse_auto_hint_columns(auto_hint_specs)
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
            code = str(row[idx_code]).strip() if (idx_code >= 0 and idx_code < len(row) and row[idx_code] is not None) else ""

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

            # ★ 单点归一：条件文本为「负反馈N」时，错误区间强制 N-N（单次触发），
            # 与莫顿 variableValue="N,N" 一致。避免 Excel 填 1~3 被误读为区间，
            # 同时让「编辑反馈」界面与莫顿最终生成都显示 N-N。文本优先于 最少/最多错次 列。
            _neg_m = re.search(r'负反馈\s*(\d+)', cond_fb_text)
            if _neg_m:
                _neg_n = int(_neg_m.group(1))
                cond_min_wrong = _neg_n
                cond_max_wrong = _neg_n
                print(f"[COND_NORM] 条件文本'负反馈{_neg_n}' → 归一化区间 {_neg_n}-{_neg_n}")

            # 构建条件反馈列表
            condition_feedbacks = []
            if cond_fb_text.strip():
                # ★ 解析阶段即清洗「负反馈N：」等标签前缀，让编辑界面与莫顿生成一致显示干净文本
                _cond_fb_clean = _clean_feedback_label(cond_fb_text)
                condition_feedbacks.append({
                    "minWrong": cond_min_wrong,
                    "maxWrong": cond_max_wrong,
                    "content": _cond_fb_clean,
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
                "code": code,
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
                "stemAudio": stem_audio,
                "conditionFeedbacks": condition_feedbacks,
                "autoHints": auto_hints,
            })

        wb.close()
        return {"success": True, "questions": questions, "total": len(questions)}

    except Exception as e:
        return {"success": False, "error": f"解析失败: {str(e)}"}