"""莫顿后端 HTTP 客户端封装（精简版）。"""

import asyncio
from typing import Any, Optional

import httpx

from app.config import get_morton_api  # 莫顿 API 域名（支持运行时切换测试/正式环境）
DEFAULT_TIMEOUT = 30.0
MAX_RETRIES = 2   # 最多重试次数
RETRY_DELAY = 1.0  # 重试间隔（秒）


class MortonClientError(Exception):
    def __init__(self, message: str, status_code: int = 502):
        super().__init__(message)
        self.message = message
        self.status_code = status_code


class MortonClient:
    """封装对莫顿后端的 HTTP 请求，支持自动重试。"""

    def __init__(self, cookie: str, token: str):
        self.cookie = cookie
        self.token = token
        self._client = httpx.AsyncClient(
            base_url=get_morton_api(),
            timeout=DEFAULT_TIMEOUT,
            headers={
                "Cookie": cookie,
                "Authorization": token,
                "Content-Type": "application/json",
            },
        )

    async def post(self, path: str, **kwargs: Any) -> Any:
        """发送 POST 请求并返回 JSON，5xx/超时时自动重试。"""
        last_error = None
        for attempt in range(1, MAX_RETRIES + 2):  # 1次原始 + MAX_RETRIES次重试
            try:
                response = await self._client.request("POST", path, **kwargs)
                # 5xx 才重试，4xx 不重试（客户端错误）
                if response.status_code >= 500:
                    last_error = MortonClientError(
                        f"上游服务器错误 ({response.status_code})",
                        status_code=response.status_code,
                    )
                    if attempt <= MAX_RETRIES:
                        print(f"[RETRY] {path} 第{attempt}次5xx，{RETRY_DELAY}s后重试...")
                        await asyncio.sleep(RETRY_DELAY)
                        continue
                    raise last_error

                if response.status_code >= 400:
                    raise MortonClientError(
                        f"上游错误 ({response.status_code})",
                        status_code=response.status_code,
                    )
                return response.json()

            except httpx.TimeoutException:
                last_error = MortonClientError("请求超时", status_code=504)
                if attempt <= MAX_RETRIES:
                    print(f"[RETRY] {path} 第{attempt}次超时，{RETRY_DELAY}s后重试...")
                    await asyncio.sleep(RETRY_DELAY)
                    continue
                raise last_error

            except httpx.RequestError as exc:
                last_error = MortonClientError(f"连接失败: {exc}", status_code=502)
                if attempt <= MAX_RETRIES:
                    print(f"[RETRY] {path} 第{attempt}次连接失败，{RETRY_DELAY}s后重试...")
                    await asyncio.sleep(RETRY_DELAY)
                    continue
                raise last_error

        raise last_error  # 不应到达

    async def close(self) -> None:
        await self._client.aclose()
