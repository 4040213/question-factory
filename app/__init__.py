"""应用工厂。"""

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse
from pathlib import Path

from .routers import auth, question, metadata, screenshot


def create_app() -> FastAPI:
    app = FastAPI(title="试题批量生成工具", version="1.0.0")

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    app.include_router(auth.router)
    app.include_router(question.router)
    app.include_router(metadata.router)
    app.include_router(screenshot.router)

    static_dir = Path(__file__).parent.parent / "static"
    index_html_path = static_dir / "index.html"

    # 全局禁用缓存的响应头
    NO_CACHE_HEADERS = {
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
        "Expires": "0",
    }

    @app.get("/", response_class=HTMLResponse)
    async def serve_index():
        if not index_html_path.exists():
            return HTMLResponse("<h1>index.html 不存在</h1>", status_code=404)
        content = index_html_path.read_text(encoding="utf-8")
        return HTMLResponse(content, headers=NO_CACHE_HEADERS)

    # 静态文件统一走 /static/ 路径，禁止浏览器缓存
    @app.get("/static/{filename:path}")
    async def serve_static(filename: str):
        file_path = static_dir / filename
        if not file_path.exists():
            return HTMLResponse("Not Found", status_code=404)
        return FileResponse(file_path, headers=NO_CACHE_HEADERS)

    return app