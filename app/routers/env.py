"""莫顿环境切换接口（正式 / 测试）。"""
from fastapi import APIRouter
from pydantic import BaseModel

from app.config import get_env_info, set_env, ENV_PRESETS

router = APIRouter(prefix="/api/morton", tags=["env"])


class EnvSwitchRequest(BaseModel):
    key: str | None = None
    url: str | None = None


@router.get("/env")
async def get_env():
    return {**get_env_info(), "presets": list(ENV_PRESETS.keys())}


@router.post("/env")
async def post_env(req: EnvSwitchRequest):
    if req.url:
        info = set_env(req.url)
    elif req.key:
        if req.key not in ENV_PRESETS:
            return {"success": False, "error": f"未知环境: {req.key}"}
        info = set_env(req.key)
    else:
        return {"success": False, "error": "缺少 key 或 url"}
    return {"success": True, **info}
