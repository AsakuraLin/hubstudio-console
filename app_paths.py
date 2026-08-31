# -*- coding: utf-8 -*-
"""应用路径：开发模式 vs PyInstaller 打包后。"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def bundle_dir() -> Path:
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def data_dir() -> Path:
    """用户数据目录（config、results、imgs）— 与 exe 同目录。"""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def config_path() -> Path:
    return data_dir() / "config.json"


def templates_dir() -> Path:
    return bundle_dir() / "templates"


def ensure_user_config() -> Path:
    """首次运行：从示例配置复制 config.json 到 exe 旁。"""
    target = config_path()
    if target.exists():
        return target
    for name in ("config.example.json", "config.json"):
        src = bundle_dir() / name
        if src.is_file():
            shutil.copy2(src, target)
            return target
    target.write_text(
        '{\n  "hubstudio_api": "http://127.0.0.1:6873/api/v1",\n'
        '  "hubstudio_exe_path": "",\n  "local_api_key": ""\n}\n',
        encoding="utf-8",
    )
    return target
