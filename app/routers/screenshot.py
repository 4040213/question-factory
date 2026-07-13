"""Playwright 截图预览路由 — 调用 Node.js 子进程完成截图。

Python 的 greenlet 在当前环境无法编译，因此改用 Node.js 版 Playwright。
Python 后端仅负责调度和缓存，实际的浏览器操作由 screenshot.js 完成。
"""

import asyncio
import hashlib
import os
import shutil
import time
from pathlib import Path

from fastapi import APIRouter
from fastapi.responses import Response
from pydantic import BaseModel

router = APIRouter(prefix="/api/screenshot", tags=["截图预览"])

# ========== 配置 ==========
PROJECT_ROOT = Path(__file__).parent.parent.parent
SCREENSHOT_SCRIPT = PROJECT_ROOT / "screenshot.js"

# ★ Node.js 路径：环境变量优先 → 自动检测 PATH 中的 node → 默认路径兜底
import shutil
_NODE_ENV = os.environ.get("NODE_PATH", "")
_NODE_DEFAULT = r"C:\Users\XQ\.workbuddy\binaries\node\versions\22.22.2\node.exe"

if _NODE_ENV and Path(_NODE_ENV).exists():
    NODE_PATH = _NODE_ENV
else:
    # 尝试从 PATH 自动检测
    _detected = shutil.which("node")
    if _detected:
        NODE_PATH = _detected
    else:
        NODE_PATH = _NODE_DEFAULT

# ========== 截图缓存 ==========
_screenshot_cache: dict = {}   # {cache_key: (png_bytes, timestamp)}
CACHE_TTL = 600                # 10 分钟过期

# ========== 并发锁 ==========
_lock = asyncio.Lock()


class ScreenshotRequest(BaseModel):
    questionId: int
    cookie: str
    token: str
    selector: str = ""


@router.post("/capture")
async def capture_screenshot(req: ScreenshotRequest):
    """用 Node.js Playwright 打开莫顿题目页面，截图并返回 PNG。"""

    # 检查截图脚本是否存在
    if not SCREENSHOT_SCRIPT.exists():
        return Response(
            content=b"screenshot.js not found. Run: npm install playwright && npx playwright install chromium",
            status_code=500,
            media_type="text/plain",
        )

    # 检查缓存
    cache_key = hashlib.md5(
        f"{req.questionId}:{req.cookie[:50]}".encode()
    ).hexdigest()
    if cache_key in _screenshot_cache:
        png_bytes, ts = _screenshot_cache[cache_key]
        if time.time() - ts < CACHE_TTL:
            return Response(
                content=png_bytes,
                media_type="image/png",
                headers={"X-Cache": "HIT"},
            )

    # 加锁，防止并发截图导致浏览器资源耗尽
    async with _lock:
        try:
            # 调用 Node.js 截图脚本
            cmd = [
                NODE_PATH,
                str(SCREENSHOT_SCRIPT),
                str(req.questionId),
                req.cookie,
                req.token,
            ]
            if req.selector:
                cmd.append(req.selector)

            # 设置 NODE_PATH 以找到 playwright 模块
            import os
            env = dict(os.environ)
            env["NODE_PATH"] = str(PROJECT_ROOT / "node_modules")

            proc = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                env=env,
            )

            stdout, stderr = await asyncio.wait_for(
                proc.communicate(), timeout=60.0
            )

            if proc.returncode != 0:
                err_msg = stderr.decode(errors='replace').strip()
                return Response(
                    content=f"截图失败: {err_msg}".encode(),
                    status_code=500,
                    media_type="text/plain",
                )

            png_bytes = stdout

            # 简单校验是否为 PNG（PNG 文件头: 89 50 4E 47）
            if len(png_bytes) < 4 or png_bytes[:4] != b'\x89PNG':
                err_msg = stderr.decode(errors='replace').strip()
                return Response(
                    content=f"截图返回数据异常（不是有效 PNG）: {err_msg or png_bytes[:200]}".encode(),
                    status_code=500,
                    media_type="text/plain",
                )

            # 写入缓存
            _screenshot_cache[cache_key] = (png_bytes, time.time())

            # 清理过期缓存
            now = time.time()
            expired = [k for k, (_, ts) in _screenshot_cache.items() if now - ts > CACHE_TTL]
            for k in expired:
                del _screenshot_cache[k]

            return Response(
                content=png_bytes,
                media_type="image/png",
                headers={"X-Cache": "MISS"},
            )

        except asyncio.TimeoutError:
            return Response(
                content="截图超时（60秒）".encode("utf-8"),
                status_code=504,
                media_type="text/plain",
            )
        except Exception as e:
            return Response(
                content=f"截图异常: {str(e)}".encode(),
                status_code=500,
                media_type="text/plain",
            )


@router.get("/health")
async def screenshot_health():
    """检查截图服务是否可用。"""
    checks = {
        "node": NODE_PATH if Path(NODE_PATH).exists() else "NOT_FOUND",
        "screenshot_script": str(SCREENSHOT_SCRIPT) if SCREENSHOT_SCRIPT.exists() else "NOT_FOUND",
        "playwright_module": str(PROJECT_ROOT / "node_modules" / "playwright") if (PROJECT_ROOT / "node_modules" / "playwright").exists() else "NOT_FOUND",
    }
    all_ok = all(v != "NOT_FOUND" for v in checks.values())
    return {
        "status": "ok" if all_ok else "degraded",
        "checks": checks,
    }
