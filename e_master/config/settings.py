"""全局配置：从 .env 加载，禁止 API Key 硬编码。"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

# e_master/config/settings.py -> e_master/config -> e_master -> 项目根目录
PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _env(name: str, default: str = "") -> str:
    value = os.getenv(name, default)
    return value.strip() if isinstance(value, str) else default


@dataclass(frozen=True)
class AppSettings:
    """集中管理环境变量，后续 AI 模块复用。"""

    deepseek_api_key: str = ""
    deepseek_base_url: str = "https://api.deepseek.com"
    deepseek_model: str = "deepseek-chat"
    log_level: str = "INFO"

    @classmethod
    def load(cls, env_path: Path | None = None) -> "AppSettings":
        load_dotenv(env_path or (PROJECT_ROOT / ".env"))
        return cls(
            deepseek_api_key=_env("DEEPSEEK_API_KEY"),
            deepseek_base_url=_env("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
            deepseek_model=_env("MODEL", "deepseek-chat"),
            log_level=_env("E_MASTER_LOG_LEVEL", "INFO"),
        )
