"""认证模块测试 - Cookie 解析、JWT 解码、Token 过期检测。"""

import time

import pytest
from app.routers.auth import (
    _parse_cookie_value,
    _decode_jwt_payload,
    _extract_operator_id,
    check_token_expiry,
)


def _build_jwt(payload: dict) -> str:
    """构造一个伪 JWT token（header.payload.signature）。"""
    import base64
    import json

    header = base64.urlsafe_b64encode(json.dumps({"alg": "HS256"}).encode()).rstrip(b"=").decode()
    payload_b64 = base64.urlsafe_b64encode(json.dumps(payload).encode()).rstrip(b"=").decode()
    return f"{header}.{payload_b64}.fake_signature"


class TestParseCookieValue:
    """Cookie 字符串解析测试。"""

    def test_found(self):
        cookie = "tokenCrm=abc123; userId=42; theme=dark"
        assert _parse_cookie_value(cookie, "tokenCrm") == "abc123"
        assert _parse_cookie_value(cookie, "userId") == "42"
        assert _parse_cookie_value(cookie, "theme") == "dark"

    def test_not_found(self):
        cookie = "tokenCrm=abc123; userId=42"
        assert _parse_cookie_value(cookie, "missing") is None

    def test_empty_cookie(self):
        assert _parse_cookie_value("", "tokenCrm") is None
        assert _parse_cookie_value(None, "tokenCrm") is None

    def test_no_equals_in_pair(self):
        cookie = "tokenCrm=abc123; bad_entry; userId=42"
        assert _parse_cookie_value(cookie, "tokenCrm") == "abc123"
        assert _parse_cookie_value(cookie, "userId") == "42"

    def test_empty_value(self):
        cookie = "tokenCrm=; userId=42"
        assert _parse_cookie_value(cookie, "tokenCrm") is None

    def test_whitespace_handling(self):
        cookie = " tokenCrm = abc123 ; userId = 42 "
        assert _parse_cookie_value(cookie, "tokenCrm") == "abc123"


class TestDecodeJwtPayload:
    """JWT payload 解码测试。"""

    def test_valid_token(self):
        token = _build_jwt({"userId": 42, "exp": 9999999999})
        payload = _decode_jwt_payload(token)
        assert payload is not None
        assert payload["userId"] == 42
        assert payload["exp"] == 9999999999

    def test_invalid_token_format(self):
        assert _decode_jwt_payload("not_a_jwt") is None
        assert _decode_jwt_payload("only.one.part.extra") is None

    def test_no_payload_part(self):
        assert _decode_jwt_payload("headeronly") is None

    def test_corrupted_base64(self):
        token = "header.!!!invalid_base64!!!.sig"
        assert _decode_jwt_payload(token) is None


class TestExtractOperatorId:
    """用户 ID 提取测试。"""

    def test_from_cookie_userIdCrm(self):
        cookie = "tokenCrm=abc; userIdCrm=100"
        assert _extract_operator_id(cookie) == 100

    def test_from_cookie_userId(self):
        cookie = "userId=200"
        assert _extract_operator_id(cookie) == 200

    def test_from_cookie_uid(self):
        cookie = "uid=300"
        assert _extract_operator_id(cookie) == 300

    def test_from_jwt(self):
        token = _build_jwt({"userId": 999})
        cookie = f"tokenCrm={token}"
        assert _extract_operator_id(cookie) == 999

    def test_not_found(self):
        cookie = "randomKey=blah"
        assert _extract_operator_id(cookie) is None

    def test_priority_cookie_over_jwt(self):
        """Cookie 中的 userId 优先于 JWT 中的 userId。"""
        token = _build_jwt({"userId": 888})
        cookie = f"userId=42; tokenCrm={token}"
        assert _extract_operator_id(cookie) == 42


class TestCheckTokenExpiry:
    """Token 过期检测测试。"""

    def test_not_expired(self):
        future = int(time.time()) + 7200  # 2小时后过期
        token = _build_jwt({"exp": future})
        result = check_token_expiry(token)
        assert result["isExpired"] is False
        assert result["willExpireSoon"] is False
        assert result["remainingSeconds"] > 0

    def test_expired(self):
        past = int(time.time()) - 3600  # 1小时前已过期
        token = _build_jwt({"exp": past})
        result = check_token_expiry(token)
        assert result["isExpired"] is True
        assert result["remainingSeconds"] == 0

    def test_will_expire_soon(self):
        soon = int(time.time()) + 300  # 5分钟后过期
        token = _build_jwt({"exp": soon})
        result = check_token_expiry(token)
        assert result["isExpired"] is False
        assert result["willExpireSoon"] is True
        assert result["remainingSeconds"] > 0

    def test_no_exp_claim(self):
        """无 exp 字段时返回空信息。"""
        token = _build_jwt({"userId": 42})
        result = check_token_expiry(token)
        assert result["expiresAt"] is None
        assert result["isExpired"] is False
        assert result["willExpireSoon"] is False
        assert result["remainingSeconds"] is None
