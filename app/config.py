"""集中配置：莫顿环境域名切换。

支持在「正式环境」与「测试环境」之间**运行时切换**，无需改动代码：
- 预设环境见 ENV_PRESETS
- 运行时调用 set_env("test" / "formal") 切换，选择会持久化到项目根目录 .env
- 通过系统环境变量 MORTON_API 可强制覆盖（优先级最高）

前端通过 /api/morton/env 接口读写当前环境。
"""
import os
import threading
from pathlib import Path


def _load_env_file(path: Path) -> None:
    """极简 .env 解析器（仅标准库，避免引入第三方依赖）。"""
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        # 仅当该环境变量尚未设置时才用 .env 的值（环境变量优先）
        os.environ.setdefault(key, value)


# .env 位于项目根目录（app/ 的上一级）
_ROOT = Path(__file__).resolve().parent.parent
_ENV_FILE = _ROOT / ".env"
_load_env_file(_ENV_FILE)

# 环境预设
ENV_PRESETS = {
    "formal": {"label": "正式环境", "url": "https://tms-mx.xueqiulearning.com"},
    "test":   {"label": "测试环境", "url": "https://tms.testing.xueqiulearning.com"},
}

_env_lock = threading.Lock()
# 初始值：系统环境变量 > .env > 正式环境默认
_current_url = os.getenv("MORTON_API") or ENV_PRESETS["formal"]["url"]


def get_morton_api() -> str:
    """返回当前生效的莫顿 API 域名（请求时调用，支持运行时切换）。"""
    return _current_url


def _key_for_url(url: str) -> str:
    for k, v in ENV_PRESETS.items():
        if v["url"] == url:
            return k
    return "custom"


def get_env_info() -> dict:
    """返回当前环境信息，供前端展示。"""
    url = _current_url
    key = _key_for_url(url)
    preset = ENV_PRESETS.get(key)
    return {
        "key": key,
        "label": preset["label"] if preset else "自定义环境",
        "url": url,
    }


def set_env(key_or_url: str) -> dict:
    """切换环境：接受预设 key（'formal'/'test'）或完整 URL，并持久化到 .env。"""
    global _current_url
    with _env_lock:
        if key_or_url in ENV_PRESETS:
            url = ENV_PRESETS[key_or_url]["url"]
        else:
            url = key_or_url
        _current_url = url
        os.environ["MORTON_API"] = url
        _persist(url)
    return get_env_info()


def _persist(url: str) -> None:
    """把 MORTON_API 写入 .env（保留其他配置项）。"""
    try:
        lines = []
        if _ENV_FILE.exists():
            for raw in _ENV_FILE.read_text(encoding="utf-8").splitlines():
                if raw.strip().startswith("MORTON_API="):
                    continue
                if raw.strip():
                    lines.append(raw)
        lines.append(f"MORTON_API={url}")
        _ENV_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except Exception as e:
        print(f"[ENV] 持久化 .env 失败: {e}")
