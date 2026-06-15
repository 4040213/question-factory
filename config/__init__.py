"""配置加载模块。"""

from pathlib import Path
from typing import Dict, Any
import yaml


_CONFIG_PATH = Path(__file__).parent / "mappings.yaml"


def load_mappings() -> Dict[str, Any]:
    """从 YAML 文件加载映射表。"""
    if not _CONFIG_PATH.exists():
        raise FileNotFoundError(f"映射配置文件不存在: {_CONFIG_PATH}")
    
    with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    
    return {
        "subjects": data.get("subjects", {}),
        "grades": data.get("grades", {}),
        "stages": data.get("stages", {}),
        "difficulties": data.get("difficulties", {}),
    }


# 模块级缓存：启动时加载一次
_MAPPINGS = load_mappings()

SUBJECT_MAP: Dict[str, int] = _MAPPINGS["subjects"]
GRADE_MAP: Dict[str, int] = _MAPPINGS["grades"]
STAGE_MAP: Dict[str, int] = _MAPPINGS["stages"]
DIFFICULTY_MAP: Dict[str, int] = _MAPPINGS["difficulties"]
