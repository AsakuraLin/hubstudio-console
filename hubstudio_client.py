# -*- coding: utf-8 -*-
"""HubStudio 客户端：检测 API、按配置路径启动。"""

from __future__ import annotations

import subprocess
import time
from pathlib import Path

import requests

_HUBSTUDIO_EXE_NAMES = (
    "HubStudio.exe",
    "hubstudio.exe",
    "HubStudio Client.exe",
    "HubStudio客户端.exe",
)


def normalize_hubstudio_exe_path(raw: str) -> str:
    return (raw or "").strip().strip('"').strip("'")


def resolve_hubstudio_exe(path_text: str) -> Path | None:
    """支持填安装目录或 HubStudio.exe 完整路径。"""
    text = normalize_hubstudio_exe_path(path_text)
    if not text:
        return None
    p = Path(text)
    if p.is_file():
        return p
    if p.is_dir():
        for name in _HUBSTUDIO_EXE_NAMES:
            candidate = p / name
            if candidate.is_file():
                return candidate
        for child in p.rglob("HubStudio.exe"):
            if child.is_file():
                return child
    return None


def hubstudio_api_ok(
    api_base: str,
    api_key: str = "",
    timeout_sec: float = 5.0,
) -> bool:
    base = (api_base or "").rstrip("/")
    if not base:
        return False
    url = f"{base}/env/list"
    headers: dict[str, str] = {}
    if api_key:
        headers["Authorization"] = api_key
    try:
        resp = requests.post(
            url,
            json={"current": 1, "size": 1},
            headers=headers,
            timeout=timeout_sec,
        )
        if resp.status_code != 200:
            return False
        data = resp.json()
        return data.get("code") == 0
    except Exception:
        return False


def launch_hubstudio(exe: Path) -> bool:
    try:
        subprocess.Popen(
            [str(exe)],
            cwd=str(exe.parent),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
        )
        return True
    except Exception:
        try:
            subprocess.Popen([str(exe)], cwd=str(exe.parent))
            return True
        except Exception:
            return False


def ensure_hubstudio_ready(
    api_base: str,
    exe_path_text: str = "",
    api_key: str = "",
    wait_timeout_sec: float = 45.0,
) -> tuple[bool, str]:
    """
    确保 HubStudio Local API 可用。
    若未运行且配置了安装路径，则尝试启动 HubStudio。
    """
    if hubstudio_api_ok(api_base, api_key):
        return True, "HubStudio API 已连接"

    exe = resolve_hubstudio_exe(exe_path_text)
    if exe is None:
        return (
            False,
            "HubStudio 未运行。请先在 HubStudio 中开启 Local API，"
            "或在下方填写 HubStudio 安装目录后保存并重试。",
        )

    if not launch_hubstudio(exe):
        return False, f"无法启动 HubStudio: {exe}"

    deadline = time.time() + wait_timeout_sec
    while time.time() < deadline:
        if hubstudio_api_ok(api_base, api_key):
            return True, f"已启动 HubStudio（{exe}）"
        time.sleep(1.2)

    return (
        False,
        "已尝试启动 HubStudio，但 API 仍未响应。"
        "请确认 HubStudio 已登录，且 设置 → Local API 已开启（默认端口 6873）。",
    )
