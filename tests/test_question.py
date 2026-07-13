"""题目模块测试 - Excel 解析、智能布局、payload 构建、默认值回退。"""

import io
from unittest.mock import AsyncMock, patch, MagicMock

import pytest
from fastapi.testclient import TestClient

from app.__init__ import create_app
from app.routers.question import (
    _measure_paragraph,
    _calculate_smart_layout,
    _calculate_blank_smart_layout,
    _build_choice_payload as _build_insert_payload,
    _build_completion_payload,
    _apply_defaults,
    QuestionItem,
    ChoiceOption,
    BlankQuestionItem,
    BlankField,
    BatchRequest,
)


class TestMeasureParagraph:
    """文本测量测试（四阶段布局引擎 Phase 1）。"""

    def test_empty_text(self):
        m = _measure_paragraph("", 40, 500)
        assert m["line_count"] == 0
        assert m["total_height"] == 0

    def test_none_text(self):
        m = _measure_paragraph(None, 40, 500)
        assert m["line_count"] == 0

    def test_single_short_line(self):
        m = _measure_paragraph("你好", 40, 500)
        assert m["line_count"] == 1
        assert m["total_height"] > 0

    def test_long_line_wraps(self):
        """长文本应产生多行。"""
        long_text = "这是一段非常非常长的文本需要换行显示" * 3
        m = _measure_paragraph(long_text, 40, 300)  # 窄宽度强制换行
        assert m["line_count"] >= 2

    def test_multiline_with_newlines(self):
        text = "第一行\n第二行\n第三行"
        m = _measure_paragraph(text, 40, 500)
        assert m["line_count"] >= 3


class TestSmartLayout:
    """智能布局计算测试（四阶段引擎）。"""

    def _layout(self, options_count, contexts_count=1, context_texts=None, images=None):
        """辅助方法：调用布局引擎并返回解包结果。"""
        result = _calculate_smart_layout(
            options_count=options_count,
            contexts_count=contexts_count,
            context_texts=context_texts,
            question_images=images,
        )
        return result["option_positions"], result["text_positions"], result["font_size"], result["layout_mode"]

    def test_2_options(self):
        opt_pos, txt_pos, font_size, mode = self._layout(2, 1, ["1+1=?"])
        assert len(opt_pos) == 2
        assert len(txt_pos) == 1
        assert font_size >= 40

    def test_3_options(self):
        opt_pos, txt_pos, font_size, mode = self._layout(3, 1, ["测试文本"])
        assert len(opt_pos) == 3
        assert len(txt_pos) == 1

    def test_4_options(self):
        opt_pos, txt_pos, font_size, mode = self._layout(4, 1, ["四个选项的题目"])
        assert len(opt_pos) == 4

    def test_no_contexts(self):
        opt_pos, txt_pos, font_size, mode = self._layout(4, 0)
        assert len(opt_pos) == 4
        assert len(txt_pos) == 0

    def test_multiple_contexts(self):
        opt_pos, txt_pos, font_size, mode = self._layout(2, 3, ["行1", "行2", "行3"])
        assert len(opt_pos) == 2
        assert len(txt_pos) == 3

    def test_options_in_content_area(self):
        """所有选项必须在绿框区域内 (x=182~1152, y=130~580)。"""
        opt_pos, _, _, mode = self._layout(4, 1, ["测试"])
        for p in opt_pos:
            assert 182 <= p["x"] <= 1152, f"选项 x={p['x']} 超出"
            assert p["x"] + p["width"] <= 1152, f"选项右边界超出"
            assert 130 <= p["y"] <= 580, f"选项 y={p['y']} 超出"

    def test_large_font_for_short_text(self):
        """短文本（≤15字）应使用 50 号字。"""
        _, _, font_size, _ = self._layout(2, 1, ["短"])
        assert font_size == 50

    def test_medium_font_for_medium_text(self):
        """中等文本（16~40字）应使用 45 号字。"""
        text = "这是一段中等长度的测试文字用于验证字号"  # >15 字
        _, _, font_size, _ = self._layout(2, 1, [text])
        assert font_size == 45

    def test_small_font_for_long_text(self):
        """长文本（>40字）应使用 40 号字。"""
        text = "这是一段很长的文本" * 10
        _, _, font_size, _ = self._layout(2, 1, [text])
        assert font_size == 40

    def test_layout_mode_text_only(self):
        """无图片时布局模式应为 TEXT_ONLY。"""
        _, _, _, mode = self._layout(2, 1, ["测试"])
        assert mode == "TEXT_ONLY"

    def test_text_in_content_area(self):
        """文本必须在绿框区域内。"""
        _, txt_pos, _, mode = self._layout(2, 1, ["测试文本"])
        for pos in txt_pos:
            assert pos["x"] >= 182, f"文本 x={pos['x']} 超出左边界"
            assert pos["y"] >= 130, f"文本 y={pos['y']} 超出上边界"

    def test_font_size_scales_with_text_length(self):
        """文本越短字号越大。"""
        _, _, fs_short, _ = self._layout(2, 1, ["短"])
        _, _, fs_long, _ = self._layout(2, 1, ["这是一段比较长的文本内容" * 3])
        assert fs_short >= fs_long, f"短文本字号({fs_short})应≥长文本字号({fs_long})"

    def test_long_text_options_do_not_overlap_text(self):
        """长文本换行后，选项应仍位于文本下方并保留最小间距。"""
        text = "如图，已知∠BOC＝40°，OD 平分∠AOC，∠AOD＝25°，那么∠AOB 的度数是（ ）。"
        opt_pos, txt_pos, font_size, mode = self._layout(4, 1, [text])
        m = _measure_paragraph(text, font_size, 970)
        text_bottom = txt_pos[0]["y"] + m["total_height"]
        option_top = min(pos["y"] for pos in opt_pos)
        assert option_top - text_bottom >= 20


class TestBuildInsertPayload:
    """构建 insert payload 测试。"""

    def test_minimal_question(self):
        item = QuestionItem(
            stem="1+1=?",
            options=[
                ChoiceOption(label="A", content="1"),
                ChoiceOption(label="B", content="2"),
            ],
            answer="B",
            name="测试题",
            subjectId=4,
        )
        payload, question_data = _build_insert_payload(item, 4)
        assert "questionData" in payload
        qd = payload["questionData"]
        assert qd["stem"]["content"] == "1+1=?"
        assert qd["name"] == "测试题"
        assert qd["questionType"] == "CHOICE"
        assert len(qd["questionItem"]["options"]) == 2

    def test_answer_marked_correctly(self):
        item = QuestionItem(
            stem="question",
            options=[
                ChoiceOption(label="A", content="wrong"),
                ChoiceOption(label="B", content="right"),
                ChoiceOption(label="C", content="wrong2"),
            ],
            answer="B",
            name="test",
            subjectId=4,
        )
        payload, _ = _build_insert_payload(item, 4)
        options = payload["questionData"]["questionItem"]["options"]
        assert options[0]["answer"] is False
        assert options[1]["answer"] is True
        assert options[2]["answer"] is False

    def test_with_knowledge(self):
        item = QuestionItem(
            stem="question",
            options=[ChoiceOption(label="A", content="x")],
            answer="A",
            name="test",
            subjectId=4,
            knowledge="加法;减法",
        )
        payload, _ = _build_insert_payload(item, 4)
        assert payload["questionData"]["knowledge"] == ["加法", "减法"]

    def test_fallback_subject_id(self):
        """subjectId 为 0 时，使用传入的全局 subject_id。"""
        item = QuestionItem(
            stem="question",
            options=[ChoiceOption(label="A", content="x")],
            answer="A",
            name="test",
            subjectId=0,
        )
        payload, _ = _build_insert_payload(item, 13)
        assert payload["questionData"]["subjectId"] == 13

    def test_resolution_parsing(self):
        item = QuestionItem(
            stem="question",
            options=[ChoiceOption(label="A", content="x")],
            answer="A",
            name="test",
            subjectId=4,
            resolution="1920*1080",
        )
        payload, _ = _build_insert_payload(item, 4)
        assert payload["questionData"]["autoWidth"] == 1920
        assert payload["questionData"]["autoHeight"] == 1080
    def test_context_short_text_center_align(self):
        item = QuestionItem(
            stem="填空题",
            contexts=["1+1=?"],
            options=[ChoiceOption(label="A", content="1"), ChoiceOption(label="B", content="2")],
            answer="B",
            name="test",
            subjectId=4,
        )
        payload, _ = _build_insert_payload(item, 4)
        ctx = payload["questionData"]["contexts"][0]
        # contexts 不再包含 displayBoxWeight/textAlign（莫顿内部处理）
        assert ctx["coordinate"]["x"] == 202  # 文本 x=202（匹配莫顿实测）

    def test_context_long_text_center_align(self):
        item = QuestionItem(
            stem="填空题",
            contexts=["如图，已知∠BOC＝40°，OD 平分∠AOC，∠AOD＝25°，那么∠AOB 的度数是（ ）。"],
            options=[ChoiceOption(label="A", content="65"), ChoiceOption(label="B", content="50")],
            answer="A",
            name="test",
            subjectId=4,
        )
        payload, _ = _build_insert_payload(item, 4)
        ctx = payload["questionData"]["contexts"][0]
        # contexts 不再包含 displayBoxWeight/textAlign
        assert ctx["coordinate"]["x"] == 202  # 文本 x=202


class TestBlankSmartLayout:
    """填空题智能布局测试。"""

    def test_single_blank_default_resolution(self):
        layout = _calculate_blank_smart_layout("1+1=______", 1)
        assert "text_position" in layout
        assert "blank_positions" in layout
        assert len(layout["blank_positions"]) == 1
        # 空白应在绿框内容区内 (1334*750 → x=182~1152, y=130~580)
        bp = layout["blank_positions"][0]
        assert 182 <= bp["x"] <= 1152
        assert 130 <= bp["y"] <= 580

    def test_double_blank(self):
        layout = _calculate_blank_smart_layout("1+1=______, 2+2=______", 2)
        assert len(layout["blank_positions"]) == 2
        # 两个空白不应重叠
        b0, b1 = layout["blank_positions"]
        assert b0["x"] + b0["width"] <= b1["x"], "两个空白不应重叠"

    def test_long_text_smaller_font(self):
        short = _calculate_blank_smart_layout("短", 1)
        long = _calculate_blank_smart_layout("这是一段很长很长很长很长很长很长很长的内容" * 3, 1)
        assert short["text_position"]["font_size"] >= long["text_position"]["font_size"]

    def test_custom_resolution_scales(self):
        layout_1080p = _calculate_blank_smart_layout("1+1=______", 1, resolution="1920*1080")
        layout_default = _calculate_blank_smart_layout("1+1=______", 1, resolution="1334*750")
        # 1080p 的空白 x 坐标应该更大（画布更宽）
        bp_1080p = layout_1080p["blank_positions"][0]
        bp_default = layout_default["blank_positions"][0]
        assert bp_1080p["x"] > bp_default["x"], "更宽的分辨率应该有更大的 x 坐标"

    def test_text_in_content_area(self):
        layout = _calculate_blank_smart_layout("填空题测试文本内容", 3)
        tp = layout["text_position"]
        assert tp["x"] >= 182, f"文本 x={tp['x']} 超出左边界"
        for bp in layout["blank_positions"]:
            assert 182 <= bp["x"] <= 1152, f"空白 x={bp['x']} 超出"
            assert 130 <= bp["y"] <= 580, f"空白 y={bp['y']} 超出"


class TestBuildCompletionPayload:
    """填空题 payload 构建测试。"""

    def test_minimal_completion(self):
        item = BlankQuestionItem(
            stem="1+1=______",
            name="填空题测试",
            subjectId=4,
            blanks=[BlankField(varName="空1", answer="2")],
        )
        payload, question_data = _build_completion_payload(item, 4)
        assert "questionData" in payload
        qd = payload["questionData"]
        assert qd["questionType"] == "COMPLETION"
        assert qd["name"] == "填空题测试"
        assert len(qd["questionItem"]["completionBlankList"]) == 1

    def test_double_blank(self):
        item = BlankQuestionItem(
            stem="1+1=______, 2+2=______",
            name="双空测试",
            subjectId=4,
            blanks=[
                BlankField(varName="空1", answer="2"),
                BlankField(varName="空2", answer="4"),
            ],
        )
        payload, _ = _build_completion_payload(item, 4)
        blanks = payload["questionData"]["questionItem"]["completionBlankList"]
        assert len(blanks) == 2
        assert blanks[0]["blankAnswer"] == ["2"]
        assert blanks[1]["blankAnswer"] == ["4"]

    def test_blank_with_aliases(self):
        item = BlankQuestionItem(
            stem="填入颜色______",
            name="别名测试",
            subjectId=4,
            blanks=[BlankField(varName="空1", answer="红色", aliases=["红", "red"])],
        )
        payload, _ = _build_completion_payload(item, 4)
        answers = payload["questionData"]["questionItem"]["completionBlankList"][0]["blankAnswer"]
        assert "红色" in answers
        assert "红" in answers
        assert "red" in answers

    def test_resolution_propagates(self):
        item = BlankQuestionItem(
            stem="1+1=______",
            name="分辨率测试",
            subjectId=4,
            resolution="1920*1080",
            blanks=[BlankField(varName="空1", answer="2")],
        )
        payload, _ = _build_completion_payload(item, 4)
        qd = payload["questionData"]
        assert qd["autoWidth"] == 1920
        assert qd["autoHeight"] == 1080
        assert qd["resolution"] == "1920*1080"

    def test_completion_with_knowledge(self):
        item = BlankQuestionItem(
            stem="1+1=______",
            name="知识点测试",
            subjectId=4,
            knowledge="加法;进位",
            blanks=[BlankField(varName="空1", answer="2")],
        )
        payload, _ = _build_completion_payload(item, 4)
        assert payload["questionData"]["knowledge"] == ["加法", "进位"]

    def test_blank_positions_are_different(self):
        """两个空白的坐标不应相同。"""
        item = BlankQuestionItem(
            stem="a=______, b=______",
            name="坐标测试",
            subjectId=4,
            blanks=[
                BlankField(varName="空1", answer="1"),
                BlankField(varName="空2", answer="2"),
            ],
        )
        payload, _ = _build_completion_payload(item, 4)
        blanks = payload["questionData"]["questionItem"]["completionBlankList"]
        assert len(blanks) == 2
        b0, b1 = blanks[0]["coordinate"], blanks[1]["coordinate"]
        assert b0 != b1, "两个空位的坐标应该不同"


class TestBlankExcelParsing:
    """填空题 Excel 解析测试。"""

    def _make_excel_bytes(self, rows: list) -> bytes:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        for row in rows:
            ws.append(row)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()

    def test_parse_single_blank(self, client):
        data = self._make_excel_bytes([
            ["题型", "名称", "学科", "分辨率", "年级", "阶段", "难度", "知识点",
             "题干文本", "文本", "题干音频",
             "空1名称", "空1类型", "空1答案类型", "空1输入框样式", "空1答案",
             "空1答案别名", "空1键盘样式", "空1显示序号", "空1显示选项底框",
             "空1选项1", "空1选项2", "空1选项3", "空1选项4"],
            ["COMPLETION", "填空题1", "数学", "1334*750", "三年级", "春", "3", "加法",
             "1+1等于几？", "1+1=______", "",
             "空1", "内容型", "内容", "文本输入框", "2",
             "", "横向双线键盘", "是", "否",
             "", "", "", ""],
        ])
        response = client.post(
            "/api/question/parse-excel?question_type=BLANK",
            files={"file": ("blank_test.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200
        result = response.json()
        assert result["success"] is True
        assert len(result["questions"]) == 1
        q = result["questions"][0]
        assert q["questionType"] == "COMPLETION"
        assert q["name"] == "填空题1"
        assert q["subjectId"] > 0
        assert q["resolution"] == "1334*750"
        assert len(q["blanks"]) == 1
        assert q["blanks"][0]["answer"] == "2"

    def test_parse_double_blank(self, client):
        data = self._make_excel_bytes([
            ["题型", "名称", "学科", "分辨率", "年级", "阶段", "难度", "知识点",
             "题干文本", "文本",
             "空1名称", "空1类型", "空1答案类型", "空1输入框样式", "空1答案",
             "空1键盘样式", "空1显示序号", "空1显示选项底框",
             "空2名称", "空2类型", "空2答案类型", "空2输入框样式", "空2答案",
             "空2键盘样式", "空2显示序号", "空2显示选项底框"],
            ["COMPLETION", "双空", "数学", "1334*750", "", "", "", "",
             "计算", "1+1=______, 2+2=______",
             "空1", "内容型", "内容", "文本输入框", "2",
             "横向双线键盘", "是", "否",
             "空2", "内容型", "内容", "文本输入框", "4",
             "横向双线键盘", "是", "否"],
        ])
        response = client.post(
            "/api/question/parse-excel?question_type=BLANK",
            files={"file": ("double_blank.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200
        result = response.json()
        assert result["success"] is True
        q = result["questions"][0]
        assert len(q["blanks"]) == 2
        assert q["blanks"][0]["answer"] == "2"
        assert q["blanks"][1]["answer"] == "4"

    def test_parse_blank_with_aliases(self, client):
        data = self._make_excel_bytes([
            ["题型", "名称", "学科", "分辨率", "年级", "阶段", "难度", "知识点",
             "题干文本", "文本",
             "空1名称", "空1类型", "空1答案类型", "空1输入框样式", "空1答案",
             "空1答案别名", "空1键盘样式", "空1显示序号", "空1显示选项底框"],
            ["COMPLETION", "别名测试", "", "1334*750", "", "", "", "",
             "颜色填空", "苹果是______的",
             "空1", "内容型", "内容", "文本输入框", "红色",
             "红;red", "横向双线键盘", "是", "否"],
        ])
        response = client.post(
            "/api/question/parse-excel?question_type=BLANK",
            files={"file": ("alias.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200
        q = response.json()["questions"][0]
        assert len(q["blanks"][0]["aliases"]) == 2
        assert "红" in q["blanks"][0]["aliases"]


class TestBatchRequest:
    """BatchRequest 模型测试。"""

    def test_default_values(self):
        br = BatchRequest(questions=[], cookie="c", token="t")
        assert br.defaultName == "新建选择题"
        assert br.defaultResolution == "1334*750"
        assert br.defaultGrade == 0
        assert br.defaultStage == 0
        assert br.defaultDifficulty == 0

    def test_custom_defaults(self):
        br = BatchRequest(
            questions=[], cookie="c", token="t",
            defaultGrade=3, defaultStage=2, defaultDifficulty=3,
        )
        assert br.defaultGrade == 3
        assert br.defaultStage == 2
        assert br.defaultDifficulty == 3


class TestApplyDefaults:
    """默认值回填逻辑测试。"""

    def test_apply_defaults_to_empty_fields(self):
        item = QuestionItem(
            stem="1+1=?",
            options=[ChoiceOption(label="A", content="1"), ChoiceOption(label="B", content="2")],
            answer="B",
            name="",
            subjectId=0,
        )
        body = BatchRequest(
            questions=[item.model_dump()], cookie="c", token="t", subjectId=4,
            defaultName="默认题名", defaultGrade=3, defaultStage=2, defaultDifficulty=1,
            defaultWrongFeedback="再想想", defaultCorrectFeedback="答对了",
        )

        _apply_defaults(item, body)

        assert item.name == "默认题名"
        assert item.subjectId == 4
        assert item.grade == 3
        assert item.stage == 2
        assert item.difficulty == 1
        assert item.wrongFeedback == "再想想"
        assert item.correctFeedback == "答对了"

    def test_apply_defaults_keeps_excel_values(self):
        item = QuestionItem(
            stem="1+1=?",
            options=[ChoiceOption(label="A", content="1"), ChoiceOption(label="B", content="2")],
            answer="B",
            name="Excel题名",
            subjectId=13,
            grade=5,
            wrongFeedback="Excel反馈",
        )
        body = BatchRequest(
            questions=[item.model_dump()], cookie="c", token="t", subjectId=4,
            defaultName="默认题名", defaultGrade=3, defaultWrongFeedback="默认反馈",
        )

        _apply_defaults(item, body)

        assert item.name == "Excel题名"
        assert item.subjectId == 13
        assert item.grade == 5
        assert item.wrongFeedback == "Excel反馈"


# ============ FastAPI 集成测试 ============

@pytest.fixture
def client():
    """创建 FastAPI TestClient。"""
    app = create_app()
    return TestClient(app)


class TestParseExcelEndpoint:
    """Excel 解析端点测试。"""

    def _make_excel_bytes(self, rows: list) -> bytes:
        """生成测试用 Excel 的 bytes。"""
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        for row in rows:
            ws.append(row)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()

    def test_parse_valid_excel(self, client):
        data = self._make_excel_bytes([
            ["题干", "文本", "A", "B", "C", "D", "E", "F", "G", "H", "答案", "名称",
             "学科", "分辨率", "年级", "阶段", "难度", "知识点", "题目图片", "画布背景",
             "答错反馈", "答对反馈", "答错图片", "答对图片", "答错音频", "答对音频",
             "最少错次", "最多错次", "条件反馈文本", "条件图片URL", "条件音频URL"],
            ["1+1=?", "", "1", "2", "", "", "", "", "", "", "B", "加法题",
             "数学", "1334*750", "三年级", "春", "3", "", "", "",
             "", "", "", "", "", "",
             "", "", "", "", ""],
        ])
        response = client.post(
            "/api/question/parse-excel",
            files={"file": ("test.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200
        result = response.json()
        assert result["success"] is True
        assert len(result["questions"]) == 1
        q = result["questions"][0]
        assert q["stem"] == "1+1=?"
        assert q["answer"] == "B"
        assert q["subjectId"] > 0  # 数学 = 4
        assert q["grade"] > 0  # 三年级 = 3
        assert q["stage"] > 0  # 春 = 2
        assert q["difficulty"] == 3

    def test_parse_with_difficulty_text(self, client):
        """难度填文字（如'中等'）时正确映射。"""
        data = self._make_excel_bytes([
            ["题干", "文本", "A", "B", "C", "D", "E", "F", "G", "H", "答案", "名称",
             "学科", "分辨率", "年级", "阶段", "难度", "知识点", "题目图片", "画布背景",
             "答错反馈", "答对反馈", "答错图片", "答对图片", "答错音频", "答对音频",
             "最少错次", "最多错次", "条件反馈文本", "条件图片URL", "条件音频URL"],
            ["test", "", "A", "B", "", "", "", "", "", "", "A", "test",
             "数学", "1334*750", "三年级", "春", "中等", "", "", "",
             "", "", "", "", "", "",
             "", "", "", "", ""],
        ])
        response = client.post(
            "/api/question/parse-excel",
            files={"file": ("test.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200
        q = response.json()["questions"][0]
        assert q["difficulty"] == 3

    def test_parse_empty_subject_grade(self, client):
        """学科/年级为空时返回 0。"""
        data = self._make_excel_bytes([
            ["题干", "文本", "A", "B", "C", "D", "E", "F", "G", "H", "答案", "名称",
             "学科", "分辨率", "年级", "阶段", "难度", "知识点", "题目图片", "画布背景",
             "答错反馈", "答对反馈", "答错图片", "答对图片", "答错音频", "答对音频",
             "最少错次", "最多错次", "条件反馈文本", "条件图片URL", "条件音频URL"],
            ["test", "", "A", "B", "", "", "", "", "", "", "A", "test",
             "", "1334*750", "", "", "", "", "", "",
             "", "", "", "", "", "",
             "", "", "", "", ""],
        ])
        response = client.post(
            "/api/question/parse-excel",
            files={"file": ("test.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200
        q = response.json()["questions"][0]
        assert q["subjectId"] == 0
        assert q["grade"] == 0
        assert q["stage"] == 0

    def test_parse_no_file(self, client):
        """无文件上传时返回 422。"""
        response = client.post("/api/question/parse-excel")
        assert response.status_code == 422

    def test_parse_multiple_questions(self, client):
        data = self._make_excel_bytes([
            ["题干", "文本", "A", "B", "C", "D", "E", "F", "G", "H", "答案", "名称",
             "学科", "分辨率", "年级", "阶段", "难度", "知识点", "题目图片", "画布背景",
             "答错反馈", "答对反馈", "答错图片", "答对图片", "答错音频", "答对音频",
             "最少错次", "最多错次", "条件反馈文本", "条件图片URL", "条件音频URL"],
            ["Q1", "", "A", "B", "", "", "", "", "", "", "A", "n1",
             "数学", "1334*750", "", "", "", "", "", "",
             "", "", "", "", "", "",
             "", "", "", "", ""],
            ["Q2", "", "C", "D", "", "", "", "", "", "", "D", "n2",
             "英语", "1334*750", "", "", "", "", "", "",
             "", "", "", "", "", "",
             "", "", "", "", ""],
        ])
        response = client.post(
            "/api/question/parse-excel",
            files={"file": ("test.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )
        assert response.status_code == 200
        result = response.json()
        assert result["success"] is True
        assert len(result["questions"]) == 2


class TestPreviewPayloadEndpoint:
    """创建前预览数据端点测试。"""

    def test_preview_payload_builds_question_data_without_morton_create(self, client):
        with patch("app.routers.question.MortonClient") as mock_morton:
            response = client.post("/api/question/preview-payload", json={
                "subjectId": 4,
                "defaultName": "默认题名",
                "defaultGrade": 3,
                "defaultStage": 2,
                "defaultDifficulty": 1,
                "defaultWrongFeedback": "再想想",
                "defaultCorrectImg": "//cdn.example.com/right.png",
                "question": {
                    "stem": "1+1=?",
                    "contexts": ["1+1=?"],
                    "options": [
                        {"label": "A", "content": "1"},
                        {"label": "B", "content": "2"},
                    ],
                    "answer": "B",
                    "name": "",
                    "subjectId": 0,
                },
            })

        assert response.status_code == 200
        data = response.json()
        assert data["success"] is True
        qd = data["questionData"]
        assert qd["name"] == "默认题名"
        assert qd["subjectId"] == 4
        assert qd["grade"] == 3
        assert qd["stage"] == 2
        assert qd["difficulty"] == 1
        assert qd["wrongQuestionFeedback"]["content"] == "再想想"
        assert qd["rightQuestionFeedback"]["webUrl"] == "//cdn.example.com/right.png"
        assert len(qd["questionItem"]["options"]) == 2
        mock_morton.assert_not_called()


class TestPreviewCreateEndpoint:
    """真实预览题创建端点测试。"""

    def test_preview_create_uses_defaults_and_prefix(self, client):
        created = {}

        class FakeMortonClient:
            def __init__(self, cookie, token):
                self.cookie = cookie
                self.token = token
                self.calls = []
                created["client"] = self

            async def post(self, path, **kwargs):
                self.calls.append((path, kwargs))
                if path.endswith("/insert"):
                    qd = kwargs["json"]["questionData"]
                    assert qd["name"].startswith("【预览-可删除】")
                    assert qd["subjectId"] == 4
                    assert qd["grade"] == 3
                    assert qd["stage"] == 2
                    assert qd["difficulty"] == 1
                    return {"code": 200, "data": 12345}
                if path.endswith("/update"):
                    qd = kwargs["json"]["questionData"]
                    assert qd["questionId"] == 12345
                    assert qd["wrongQuestionFeedback"]["content"] == "再想想"
                    return {"code": 200}
                return {"code": 404, "message": "unexpected"}

            async def close(self):
                self.closed = True

        with patch("app.routers.question.MortonClient", FakeMortonClient):
            response = client.post("/api/question/preview-create", json={
                "cookie": "tokenCrm=t",
                "token": "t",
                "subjectId": 4,
                "defaultName": "默认题名",
                "defaultGrade": 3,
                "defaultStage": 2,
                "defaultDifficulty": 1,
                "defaultWrongFeedback": "再想想",
                "question": {
                    "stem": "1+1=?",
                    "contexts": ["1+1=?"],
                    "options": [
                        {"label": "A", "content": "1"},
                        {"label": "B", "content": "2"},
                    ],
                    "answer": "B",
                    "name": "",
                    "subjectId": 0,
                },
            })

        assert response.status_code == 200
        data = response.json()
        assert data["success"] is True
        assert data["previewQuestionId"] == 12345
        assert "手动清理" in data["warning"]
        assert len(created["client"].calls) == 2


class TestMetadataEndpoint:
    """映射表 API 端点测试。"""

    def test_mappings_endpoint(self, client):
        response = client.get("/api/metadata/mappings")
        assert response.status_code == 200
        data = response.json()
        assert "subjects" in data
        assert "grades" in data
        assert "stages" in data
        assert "difficulties" in data
        assert data["subjects"]["数学"] == 4
        assert data["grades"]["高三"] == 16
        assert data["stages"]["春"] == 2
        assert data["difficulties"]["中等"] == 3


class TestAuthEndpoint:
    """认证 API 端点测试。"""

    def _make_expired_jwt(self) -> str:
        import base64, json
        header = base64.urlsafe_b64encode(json.dumps({"alg": "HS256"}).encode()).rstrip(b"=").decode()
        payload = base64.urlsafe_b64encode(json.dumps({"userId": 1, "exp": 1}).encode()).rstrip(b"=").decode()
        return f"{header}.{payload}.sig"

    def _make_valid_jwt(self) -> str:
        import base64, json, time
        header = base64.urlsafe_b64encode(json.dumps({"alg": "HS256"}).encode()).rstrip(b"=").decode()
        payload = base64.urlsafe_b64encode(json.dumps({"userId": 1, "exp": int(time.time()) + 7200}).encode()).rstrip(b"=").decode()
        return f"{header}.{payload}.sig"

    def test_validate_with_token(self, client):
        token = self._make_valid_jwt()
        response = client.post("/api/auth/validate", json={
            "cookie": f"tokenCrm={token}; userIdCrm=42"
        })
        assert response.status_code == 200
        data = response.json()
        assert data["success"] is True
        assert data["token"] == token
        assert data["operatorId"] == 42
        assert data["isExpired"] is False

    def test_validate_no_token(self, client):
        response = client.post("/api/auth/validate", json={
            "cookie": "randomCookie=blah; userIdCrm=1"
        })
        assert response.status_code == 200
        data = response.json()
        assert data["success"] is False
        assert "error" in data

    def test_validate_expired_token(self, client):
        token = self._make_expired_jwt()
        response = client.post("/api/auth/validate", json={
            "cookie": f"tokenCrm={token}; userIdCrm=42"
        })
        assert response.status_code == 200
        data = response.json()
        assert data["isExpired"] is True
