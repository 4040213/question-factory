"""映射表测试 - 确保莫顿系统改编号时能及时发现。"""

import pytest
from config import SUBJECT_MAP, GRADE_MAP, STAGE_MAP, DIFFICULTY_MAP


class TestSubjectMap:
    """学科映射测试。"""

    def test_known_subjects(self):
        assert SUBJECT_MAP["数学"] == 4
        assert SUBJECT_MAP["英语"] == 13
        assert SUBJECT_MAP["新数学"] == 14
        assert SUBJECT_MAP["思维盒子"] == 18

    def test_no_duplicate_ids(self):
        ids = list(SUBJECT_MAP.values())
        assert len(ids) == len(set(ids)), "学科 ID 不能重复"

    def test_all_subjects_have_positive_id(self):
        for name, id_ in SUBJECT_MAP.items():
            assert id_ > 0, f"学科 {name} 的 ID 应为正数"

    def test_subject_name_not_empty(self):
        for name in SUBJECT_MAP.keys():
            assert name.strip(), "学科名称不能为空"


class TestGradeMap:
    """年级映射测试。"""

    def test_known_grades(self):
        assert GRADE_MAP["高三"] == 16
        assert GRADE_MAP["幼儿-小班"] == 9
        assert GRADE_MAP["一年级"] == 1
        assert GRADE_MAP["初三"] == 13
        assert GRADE_MAP["高二"] == 15

    def test_no_duplicate_ids(self):
        ids = list(GRADE_MAP.values())
        assert len(ids) == len(set(ids)), "年级 ID 不能重复"

    def test_grade_ordering(self):
        """验证年级编号的关键边界（小学→初中→高中）。"""
        # 注意：幼儿班编号（7~9）反而比小学高，这是莫顿系统的实际编号
        assert GRADE_MAP["六年级"] < GRADE_MAP["初一"]
        assert GRADE_MAP["初三"] < GRADE_MAP["高一"]
        assert GRADE_MAP["一年级"] < GRADE_MAP["六年级"]

    def test_all_grades_have_positive_id(self):
        for name, id_ in GRADE_MAP.items():
            assert id_ != 0, f"年级 {name} 的 ID 不应为 0"


class TestStageMap:
    """阶段映射测试。"""

    def test_known_stages(self):
        assert STAGE_MAP["春"] == 2
        assert STAGE_MAP["暑"] == 3
        assert STAGE_MAP["寒"] == 5
        assert STAGE_MAP["招生课"] == 1
        assert STAGE_MAP["专题课"] == 6

    def test_no_duplicate_ids(self):
        ids = list(STAGE_MAP.values())
        assert len(ids) == len(set(ids)), "阶段 ID 不能重复"

    def test_all_stages_have_positive_id(self):
        for name, id_ in STAGE_MAP.items():
            assert id_ > 0, f"阶段 {name} 的 ID 应为正数"


class TestDifficultyMap:
    """难度映射测试。"""

    def test_known_difficulties(self):
        assert DIFFICULTY_MAP["简单"] == 1
        assert DIFFICULTY_MAP["较易"] == 2
        assert DIFFICULTY_MAP["中等"] == 3
        assert DIFFICULTY_MAP["较难"] == 4
        assert DIFFICULTY_MAP["困难"] == 5

    def test_alias_consistency(self):
        """别名（易/难）应与全称一致。"""
        assert DIFFICULTY_MAP["易"] == DIFFICULTY_MAP["简单"]
        assert DIFFICULTY_MAP["难"] == DIFFICULTY_MAP["困难"]

    def test_difficulty_range(self):
        """所有难度值应在 1~5 范围内。"""
        for name, val in DIFFICULTY_MAP.items():
            assert 1 <= val <= 5, f"难度 {name} 的值 {val} 超出范围"


class TestMappingConfig:
    """映射配置完整性测试。"""

    def test_all_maps_are_dict(self):
        assert isinstance(SUBJECT_MAP, dict)
        assert isinstance(GRADE_MAP, dict)
        assert isinstance(STAGE_MAP, dict)
        assert isinstance(DIFFICULTY_MAP, dict)

    def test_no_empty_maps(self):
        assert len(SUBJECT_MAP) > 0, "学科映射不能为空"
        assert len(GRADE_MAP) > 0, "年级映射不能为空"
        assert len(STAGE_MAP) > 0, "阶段映射不能为空"
        assert len(DIFFICULTY_MAP) > 0, "难度映射不能为空"
