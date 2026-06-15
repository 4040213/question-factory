"""认证路由：Cookie 验证 + Token 过期检测。"""

import base64
import json
import time
from typing import Any, Dict, Optional

from fastapi import APIRouter
from pydantic import BaseModel

router = APIRouter(prefix="/api/auth", tags=["认证"])


class ValidateRequest(BaseModel):
    cookie: str


def _parse_cookie_value(cookie_string: str, key_name: str) -> Optional[str]:
    """从 Cookie 字符串中提取指定 key 的值。"""
    if not cookie_string:
        return None
    for pair in cookie_string.split(";"):
        pair = pair.strip()
        if "=" not in pair:
            continue
        key, _, value = pair.partition("=")
        if key.strip() == key_name:
            return value.strip() or None
    return None


def _decode_jwt_payload(token: str) -> Optional[dict]:
    """解码 JWT payload（无签名校验）。"""
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return None
        payload_b64 = parts[1]
        padding = "=" * (-len(payload_b64) % 4)
        decoded = base64.urlsafe_b64decode(payload_b64 + padding)
        return json.loads(decoded)
    except Exception:
        return None


def _extract_operator_id(cookie_string: str, token: Optional[str] = None) -> Optional[int]:
    """从 Cookie/JWT 中提取用户 ID。"""
    for key in ("userIdCrm", "userId", "uid"):
        value = _parse_cookie_value(cookie_string, key)
        if value:
            try:
                return int(value)
            except (ValueError, TypeError):
                continue

    if token is None:
        token = _parse_cookie_value(cookie_string, "tokenCrm")
    if token:
        payload = _decode_jwt_payload(token)
        if payload:
            for claim in ("userId", "uid", "sub", "id"):
                value = payload.get(claim)
                if value is not None:
                    try:
                        return int(value)
                    except (ValueError, TypeError):
                        pass
    return None


def check_token_expiry(token: str) -> Dict[str, Any]:
    """检查 JWT Token 过期状态，返回过期信息。"""
    payload = _decode_jwt_payload(token)
    now = time.time()
    expires_at = None
    is_expired = False
    will_expire_soon = False
    remaining_seconds = None

    if payload and "exp" in payload:
        expires_at = payload["exp"]
        remaining_seconds = expires_at - now
        is_expired = remaining_seconds <= 0
        will_expire_soon = not is_expired and remaining_seconds < 600  # 10分钟内

    return {
        "expiresAt": expires_at,
        "isExpired": is_expired,
        "willExpireSoon": will_expire_soon,
        "remainingSeconds": max(0, round(remaining_seconds)) if remaining_seconds is not None else None,
    }


@router.post("/validate")
async def validate_cookie(request: ValidateRequest) -> Dict[str, Any]:
    """验证 Cookie，解析 tokenCrm，检测过期状态。"""
    token = _parse_cookie_value(request.cookie, "tokenCrm")
    if token:
        operator_id = _extract_operator_id(request.cookie, token)
        expiry = check_token_expiry(token)
        return {
            "success": True,
            "token": token,
            "operatorId": operator_id,
            **expiry,
        }
    return {
        "success": False,
        "error": "Cookie 中未找到 tokenCrm，请重新获取",
    }
