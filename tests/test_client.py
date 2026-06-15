"""MortonClient 测试 - 自动重试、错误处理。"""

import pytest
from unittest.mock import AsyncMock, patch, MagicMock

import httpx
from app.client import MortonClient, MortonClientError


class TestMortonClientError:
    """MortonClientError 异常类测试。"""

    def test_default_status_code(self):
        err = MortonClientError("something wrong")
        assert err.status_code == 502
        assert str(err) == "something wrong"

    def test_custom_status_code(self):
        err = MortonClientError("not found", status_code=404)
        assert err.status_code == 404


class TestMortonClientRetry:
    """重试逻辑测试。"""

    @pytest.mark.asyncio
    async def test_success_first_attempt(self):
        """首次请求成功，不重试。"""
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.json.return_value = {"code": 200, "data": 123}

        with patch.object(httpx.AsyncClient, "request", AsyncMock(return_value=mock_resp)):
            client = MortonClient(cookie="test_cookie", token="test_token")
            try:
                result = await client.post("/test/path")
                assert result == {"code": 200, "data": 123}
                # 只调用了一次
                assert mock_resp.json.call_count == 1
            finally:
                await client.close()

    @pytest.mark.asyncio
    async def test_retry_on_500(self):
        """5xx 错误时重试，最终成功。"""
        fail_resp = MagicMock()
        fail_resp.status_code = 500

        success_resp = MagicMock()
        success_resp.status_code = 200
        success_resp.json.return_value = {"code": 200, "data": "ok"}

        mock_request = AsyncMock(side_effect=[fail_resp, fail_resp, success_resp])

        with patch.object(httpx.AsyncClient, "request", mock_request):
            client = MortonClient(cookie="test_cookie", token="test_token")
            try:
                result = await client.post("/test/path")
                assert result == {"code": 200, "data": "ok"}
                # 原始1次 + 2次重试 = 3次
                assert mock_request.call_count == 3
            finally:
                await client.close()

    @pytest.mark.asyncio
    async def test_no_retry_on_400(self):
        """4xx 客户端错误不重试。"""
        fail_resp = MagicMock()
        fail_resp.status_code = 400

        mock_request = AsyncMock(return_value=fail_resp)

        with patch.object(httpx.AsyncClient, "request", mock_request):
            client = MortonClient(cookie="test_cookie", token="test_token")
            try:
                with pytest.raises(MortonClientError) as exc:
                    await client.post("/test/path")
                assert exc.value.status_code == 400
                # 只调用1次，不重试
                assert mock_request.call_count == 1
            finally:
                await client.close()

    @pytest.mark.asyncio
    async def test_max_retries_exceeded(self):
        """所有重试都失败时抛出异常。"""
        fail_resp = MagicMock()
        fail_resp.status_code = 502

        mock_request = AsyncMock(return_value=fail_resp)

        with patch.object(httpx.AsyncClient, "request", mock_request):
            client = MortonClient(cookie="test_cookie", token="test_token")
            try:
                with pytest.raises(MortonClientError) as exc:
                    await client.post("/test/path")
                assert exc.value.status_code == 502
                # 1次原始 + 2次重试 = 3次
                assert mock_request.call_count == 3
            finally:
                await client.close()

    @pytest.mark.asyncio
    async def test_retry_on_timeout(self):
        """超时时重试。"""
        success_resp = MagicMock()
        success_resp.status_code = 200
        success_resp.json.return_value = {"code": 200}

        mock_request = AsyncMock(side_effect=[
            httpx.TimeoutException("timed out"),
            success_resp,
        ])

        with patch.object(httpx.AsyncClient, "request", mock_request):
            client = MortonClient(cookie="test_cookie", token="test_token")
            try:
                result = await client.post("/test/path")
                assert result == {"code": 200}
                assert mock_request.call_count == 2
            finally:
                await client.close()

    @pytest.mark.asyncio
    async def test_retry_on_connection_error(self):
        """连接失败时重试。"""
        success_resp = MagicMock()
        success_resp.status_code = 200
        success_resp.json.return_value = {"code": 200}

        mock_request = AsyncMock(side_effect=[
            httpx.RequestError("connection refused"),
            success_resp,
        ])

        with patch.object(httpx.AsyncClient, "request", mock_request):
            client = MortonClient(cookie="test_cookie", token="test_token")
            try:
                result = await client.post("/test/path")
                assert result == {"code": 200}
                assert mock_request.call_count == 2
            finally:
                await client.close()

    @pytest.mark.asyncio
    async def test_timeout_max_retries_exceeded(self):
        """超时重试全部失败。"""
        mock_request = AsyncMock(side_effect=httpx.TimeoutException("timed out"))

        with patch.object(httpx.AsyncClient, "request", mock_request):
            client = MortonClient(cookie="test_cookie", token="test_token")
            try:
                with pytest.raises(MortonClientError) as exc:
                    await client.post("/test/path")
                assert exc.value.status_code == 504  # 超时映射为 504
                assert mock_request.call_count == 3
            finally:
                await client.close()


class TestMortonClientInit:
    """客户端初始化测试。"""

    def test_headers_are_set(self):
        client = MortonClient(cookie="my_cookie", token="my_token")
        assert client.cookie == "my_cookie"
        assert client.token == "my_token"
        assert client._client.headers["Cookie"] == "my_cookie"
        assert client._client.headers["Authorization"] == "my_token"
