"""元数据路由：提供前端动态映射表。"""

from fastapi import APIRouter
from typing import Dict, Any

from config import SUBJECT_MAP, GRADE_MAP, STAGE_MAP, DIFFICULTY_MAP

router = APIRouter(prefix="/api/metadata", tags=["元数据"])


@router.get("/mappings")
async def get_mappings() -> Dict[str, Any]:
    """返回所有映射表，供前端动态渲染下拉框和显示名。"""
    return {
        "success": True,
        "subjects": SUBJECT_MAP,
        "grades": GRADE_MAP,
        "stages": STAGE_MAP,
        "difficulties": DIFFICULTY_MAP,
    }
