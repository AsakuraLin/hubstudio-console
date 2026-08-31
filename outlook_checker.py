# -*- coding: utf-8 -*-
"""
HubStudio + Microsoft Outlook 账号可用性检测工具

用法:
  1. 确保 HubStudio 客户端已打开，且 Local API 已启用 (默认 http://127.0.0.1:6873)
  2. 编辑 accounts.csv，填入机子号/邮箱/密码（机子号=环境名称）
  3. pip install -r requirements.txt
  4. python outlook_checker.py              # 自动登录并判断可用性
     python outlook_checker.py --open-only  # 只打开环境并进入登录页

半自动模式: 遇到验证码/MFA/安全验证时，脚本会暂停，你在浏览器里手动完成后按回车继续。
"""

from __future__ import annotations

import argparse
import asyncio
import base64
import csv
import json
import re
import sys
import threading
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Callable

import requests
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeout, sync_playwright

from foxmail_automation import (
    minimize_hubstudio_windows,
    register_automation_browser,
    restore_hubstudio_windows,
    reveal_hubstudio_windows_for_screenshot,
    show_hubstudio_windows_for_debug,
)
from app_paths import data_dir
from recovery_email import (
    find_recovery_email_input,
    is_code_verify_page,
    is_definitely_recovery_flow_page,
    is_identity_verification_page,
    is_ms_auth_page_shell,
    is_phone_verify_page,
    is_recovery_bind_page,
    set_keep_background,
    try_bind_recovery_on_page,
)


class AccountStatus(str, Enum):
    OK = "可用"
    LOGIN_OK = "登入"
    AMZ_BANNED = "AMZ账号被封"
    ACCOUNT_NOT_FOUND = "找不到账户"
    BAD_PASSWORD = "密码错误"
    LOCKED = "账户被锁"
    BAD_CREDENTIALS = "密码错误或账号不存在"  # 兼容旧结果
    NEED_VERIFY = "需要额外验证(MFA/手机/邮箱)"
    NEED_PHONE = "需要电话认证"
    NEED_IDENTITY = "无法绑定辅助邮箱"
    NEED_RECOVERY = "需要绑定辅助邮箱"
    WAIT_CODE = "等待验证码"
    RECOVERY_BOUND = "已绑定辅助邮箱"
    ALREADY_BOUND = "已绑定辅助邮箱"
    RECOVERY_FAILED = "辅助邮箱绑定失败"
    RECOVERY_PASSWORD_BAD = "辅助邮箱密码错误"
    CAPTCHA = "需要人机验证"
    STAY_SIGNED_IN = "等待确认保持登录"
    TIMEOUT = "超时"
    NETWORK_CARD = "网卡"
    OPEN_FAILED = "HubStudio打开失败"
    RESOLVE_FAILED = "找不到机子号对应环境"
    UNKNOWN = "未知状态"


@dataclass
class EnvRef:
    machine_id: str
    container_code: str
    serial_number: str
    container_name: str


@dataclass
class CheckResult:
    machine_id: str
    email: str
    status: AccountStatus
    detail: str
    container_code: str = ""
    final_url: str = ""
    checked_at: str = ""
    screenshot_path: str = ""
    recovery_email: str = ""
    awaiting_code: bool = False


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_accounts(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = []
        for row in reader:
            machine_id = (
                row.get("机子号")
                or row.get("machine_id")
                or row.get("serial_number")
                or row.get("containerName")
                or row.get("环境名称")
                or ""
            ).strip()
            email = (row.get("邮箱") or row.get("email") or "").strip()
            password = (row.get("密码") or row.get("password") or "").strip()
            if machine_id:
                rows.append(
                    {
                        "machine_id": machine_id,
                        "email": email,
                        "password": password,
                    }
                )
        return rows


_RUNTIME: dict[str, Any] = {
    "hubstudio_timeout_sec": 60,
    "hubstudio_retries": 3,
    "browser_start_stagger_sec": 1.5,
    "cdp_connect_timeout_ms": 60000,
    "browser_headless": False,
    "browser_silent_open": True,
    "browser_minimize_on_start": True,
    "keep_background": True,
    "screenshot_enabled": True,
    "screenshot_max_wait_sec": 8.0,
    "screenshot_dir": "",
    "env_cache": None,
    "fast_fill": True,
    "login_form_wait_sec": 6,
    "acquire_page_max_sec": 12,
}


def apply_runtime_config(config: dict[str, Any] | None) -> None:
    """从 config.json 注入运行时参数（批量任务开始前调用）。"""
    if not config:
        return
    _RUNTIME["hubstudio_timeout_sec"] = int(config.get("hubstudio_timeout_sec", 60))
    _RUNTIME["hubstudio_retries"] = int(config.get("hubstudio_retries", 3))
    _RUNTIME["browser_start_stagger_sec"] = float(config.get("browser_start_stagger_sec", 0.5))
    _RUNTIME["cdp_connect_timeout_ms"] = int(config.get("cdp_connect_timeout_ms", 60000))
    browser_cfg = config.get("browser") or {}
    _RUNTIME["browser_headless"] = bool(browser_cfg.get("headless", False))
    _RUNTIME["browser_silent_open"] = bool(browser_cfg.get("silent_open", True))
    _RUNTIME["browser_minimize_on_start"] = bool(browser_cfg.get("minimize_on_start", True))
    _RUNTIME["browser_lightweight"] = bool(browser_cfg.get("lightweight", True))
    _RUNTIME["keep_background"] = bool(browser_cfg.get("keep_background", True))
    set_keep_background(bool(_RUNTIME["keep_background"]))
    shot_cfg = config.get("screenshot") or {}
    _RUNTIME["screenshot_enabled"] = bool(shot_cfg.get("enabled", True))
    _RUNTIME["screenshot_max_wait_sec"] = float(shot_cfg.get("max_wait_sec", 8.0))
    _RUNTIME["fast_fill"] = bool(config.get("fast_fill", True))
    _RUNTIME["login_form_wait_sec"] = int(config.get("login_form_wait_sec", 6))
    _RUNTIME["acquire_page_max_sec"] = int(config.get("acquire_page_max_sec", 12))


def _fast_fill_enabled() -> bool:
    return bool(_RUNTIME.get("fast_fill", True))


def hubstudio_post(
    api_base: str,
    endpoint: str,
    payload: dict[str, Any],
    api_key: str,
    *,
    timeout_sec: int | None = None,
    retries: int | None = None,
) -> dict[str, Any]:
    url = f"{api_base.rstrip('/')}/{endpoint.lstrip('/')}"
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = api_key
        headers["X-API-KEY"] = api_key
    timeout = timeout_sec if timeout_sec is not None else int(_RUNTIME["hubstudio_timeout_sec"])
    max_attempts = retries if retries is not None else int(_RUNTIME["hubstudio_retries"])
    last_exc: Exception | None = None
    for attempt in range(max_attempts):
        try:
            resp = requests.post(url, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            return resp.json()
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_exc = exc
            if attempt + 1 < max_attempts:
                time.sleep(2 * (attempt + 1))
                continue
            raise
    if last_exc:
        raise last_exc
    raise RuntimeError("HubStudio API 请求失败")


def resolve_env(api_base: str, machine_id: str, api_key: str) -> EnvRef | None:
    """机子号优先按环境名称(containerName)查找，其次按备注/序号/环境ID。"""
    cache = _RUNTIME.get("env_cache")
    if isinstance(cache, dict):
        hit = cache.get(machine_id)
        if hit is not None:
            return hit

    name_items: list[dict] = []
    # 1) 环境名称精确匹配（你截图里搜索的就是这个）
    data = hubstudio_post(
        api_base,
        "env/list",
        {"current": 1, "size": 20, "containerName": machine_id},
        api_key,
    )
    name_items = (data.get("data") or {}).get("list") or []
    for item in name_items:
        name = str(item.get("containerName") or "").strip()
        if name == machine_id:
            return EnvRef(
                machine_id=machine_id,
                container_code=str(item["containerCode"]),
                serial_number=str(item.get("serialNumber") or ""),
                container_name=name,
            )

    # 2) 备注精确匹配
    data = hubstudio_post(
        api_base,
        "env/list",
        {"current": 1, "size": 20, "remark": machine_id},
        api_key,
    )
    items = (data.get("data") or {}).get("list") or []
    for item in items:
        remark = str(item.get("remark") or "").strip()
        if remark == machine_id:
            return EnvRef(
                machine_id=machine_id,
                container_code=str(item["containerCode"]),
                serial_number=str(item.get("serialNumber") or ""),
                container_name=str(item.get("containerName") or ""),
            )

    # 3) 纯数字：按序号或环境ID查
    if machine_id.isdigit():
        data = hubstudio_post(
            api_base,
            "env/list",
            {"current": 1, "size": 20, "serialNumbers": [int(machine_id)]},
            api_key,
        )
        items = (data.get("data") or {}).get("list") or []
        if items:
            item = items[0]
            return EnvRef(
                machine_id=machine_id,
                container_code=str(item["containerCode"]),
                serial_number=str(item.get("serialNumber") or ""),
                container_name=str(item.get("containerName") or ""),
            )

        data = hubstudio_post(
            api_base,
            "env/list",
            {"current": 1, "size": 20, "containerCodes": [machine_id]},
            api_key,
        )
        items = (data.get("data") or {}).get("list") or []
        if items:
            item = items[0]
            return EnvRef(
                machine_id=machine_id,
                container_code=str(item["containerCode"]),
                serial_number=str(item.get("serialNumber") or ""),
                container_name=str(item.get("containerName") or ""),
            )

    # 4) 名称模糊：返回唯一结果时采用（复用步骤 1 结果，避免重复 API）
    if len(name_items) == 1:
        item = name_items[0]
        return EnvRef(
            machine_id=machine_id,
            container_code=str(item["containerCode"]),
            serial_number=str(item.get("serialNumber") or ""),
            container_name=str(item.get("containerName") or ""),
        )

    return None


def _env_ref_from_item(item: dict, machine_id: str) -> EnvRef:
    return EnvRef(
        machine_id=machine_id,
        container_code=str(item["containerCode"]),
        serial_number=str(item.get("serialNumber") or ""),
        container_name=str(item.get("containerName") or ""),
    )


def build_env_cache(api_base: str, api_key: str, *, page_size: int = 100) -> dict[str, EnvRef]:
    """批量任务开始前拉取环境列表，减少逐账号 env/list 请求。"""
    index: dict[str, EnvRef] = {}
    page = 1
    total = None
    while page <= 50:
        data = hubstudio_post(
            api_base,
            "env/list",
            {"current": page, "size": page_size},
            api_key,
        )
        payload = data.get("data") or {}
        items = payload.get("list") or []
        if total is None:
            try:
                total = int(payload.get("total") or 0)
            except (TypeError, ValueError):
                total = 0
        if not items:
            break
        for item in items:
            ref = _env_ref_from_item(item, str(item.get("containerName") or "").strip())
            for key in (
                str(item.get("containerName") or "").strip(),
                str(item.get("remark") or "").strip(),
                str(item.get("serialNumber") or "").strip(),
                str(item.get("containerCode") or "").strip(),
            ):
                if key and key not in index:
                    indexed = _env_ref_from_item(item, key)
                    index[key] = indexed
        if total and page * page_size >= total:
            break
        if len(items) < page_size:
            break
        page += 1
    return index


def set_env_cache(cache: dict[str, EnvRef] | None) -> None:
    _RUNTIME["env_cache"] = cache


def clear_env_cache() -> None:
    _RUNTIME["env_cache"] = None


def start_browser(
    api_base: str,
    container_code: str,
    login_url: str,
    api_key: str,
    *,
    force_visible: bool = False,
    open_login_tab: bool = True,
) -> tuple[int | None, str]:
    """
    启动/复用环境并返回 debuggingPort。
    open_login_tab=False：不传 containerTabs，避免把当前问题页冲回登录页（点「打开」用）。
    """
    payload: dict[str, Any] = {
        "containerCode": container_code,
        "skipSystemResourceCheck": True,
        # 不覆盖客户端已有标签，保留绑定辅助邮箱等问题页
        "shouldCloseTabsOnOpen": False,
    }
    if open_login_tab and login_url:
        payload["containerTabs"] = [login_url]
    use_headless = bool(
        _RUNTIME.get("browser_headless")
        or (_silent_open_enabled() and not force_visible)
    )
    if use_headless:
        payload["isHeadless"] = True
    else:
        payload["isHeadless"] = False
    payload["args"] = _build_browser_start_args(
        use_headless=use_headless, force_visible=force_visible
    )
    data = hubstudio_post(api_base, "browser/start", payload, api_key)
    code = data.get("code")
    # -10013 环境正在运行：部分版本仍返回 debuggingPort；否则再用不带 tabs 的方式取端口
    if code == -10013:
        info = data.get("data") or {}
        port = info.get("debuggingPort")
        if port:
            remember_debug_port(container_code, int(port))
            return int(port), ""
        # 再请求一次不带登录页，仅取端口
        if open_login_tab:
            return start_browser(
                api_base,
                container_code,
                login_url,
                api_key,
                force_visible=True,
                open_login_tab=False,
            )
        return None, data.get("msg") or "环境正在运行但未返回调试端口"
    if code != 0:
        return None, data.get("msg") or str(data)
    info = data.get("data") or {}
    port = info.get("debuggingPort")
    if not port:
        return None, "未返回 debuggingPort"
    remember_debug_port(container_code, int(port))
    return int(port), ""


def bring_browser_to_foreground(api_base: str, container_code: str, api_key: str) -> tuple[bool, str]:
    """HubStudio 官方接口：把已打开环境窗口置顶。"""
    try:
        data = hubstudio_post(
            api_base,
            "browser/foreground",
            {"containerCode": container_code},
            api_key,
            timeout_sec=15,
            retries=1,
        )
        if data.get("code") == 0:
            return True, "已调用 browser/foreground 置顶"
        return False, data.get("msg") or str(data)
    except Exception as exc:
        return False, str(exc)


def browser_is_running(api_base: str, container_code: str, api_key: str) -> bool | None:
    """查询环境是否已开启。True/False；查不到返回 None。"""
    try:
        data = hubstudio_post(
            api_base,
            "browser/all-browser-status",
            {"containerCodes": [container_code]},
            api_key,
            timeout_sec=10,
            retries=1,
        )
        if data.get("code") != 0:
            return None
        payload = data.get("data") or {}
        # 兼容多种返回结构
        items = payload if isinstance(payload, list) else (
            payload.get("list")
            or payload.get("statusList")
            or payload.get("browsers")
            or []
        )
        if isinstance(payload, dict) and not items:
            # { containerCode: status }
            st = payload.get(container_code)
            if st is not None:
                return int(st) in (0, 1)  # 0已开启 1开启中
        for it in items:
            if not isinstance(it, dict):
                continue
            code = str(it.get("containerCode") or it.get("container_code") or "")
            if code != container_code:
                continue
            st = it.get("status", it.get("browserStatus", it.get("state")))
            try:
                return int(st) in (0, 1)
            except Exception:
                return bool(st)
    except Exception:
        return None
    return None


def ensure_browser_port(
    api_base: str,
    container_code: str,
    login_url: str,
    api_key: str,
) -> tuple[int | None, str]:
    """对已打开的环境再次 start 也能返回 debuggingPort；失败时自动重试。"""
    retries = int(_RUNTIME["hubstudio_retries"])
    last_err = ""
    tried_visible_fallback = False
    for attempt in range(retries):
        try:
            force_visible = tried_visible_fallback
            if attempt == 0:
                log_step("正在后台启动 HubStudio 环境")
            port, err = start_browser(
                api_base,
                container_code,
                login_url,
                api_key,
                force_visible=force_visible,
            )
            if port is not None:
                remember_debug_port(container_code, port)
                if force_visible or not _effective_headless():
                    minimize_automation_browser_now()
                else:
                    log_step(f"环境已静默启动，调试端口 {port}")
                return port, err
            last_err = err or "未知错误"
            if _silent_open_enabled() and not tried_visible_fallback:
                tried_visible_fallback = True
                log_step("静默无头启动失败，改用最小化窗口重试")
                continue
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_err = str(exc)
        if attempt + 1 < retries:
            time.sleep(2 * (attempt + 1) if _fast_fill_enabled() else 3 * (attempt + 1))
    hint = "（并行过多或 HubStudio 负载高，建议降至并行 3）"
    if "timed out" in last_err.lower() or "timeout" in last_err.lower():
        return None, f"HubStudio 打开浏览器超时{hint}: {last_err}"
    return None, last_err or f"HubStudio 打开浏览器失败{hint}"


_step_ctx = threading.local()
# containerCode -> debuggingPort（查号时登记，点「打开」复用，避免再带登录 URL 冲掉问题页）
_debug_ports_by_container: dict[str, int] = {}
_debug_ports_lock = threading.Lock()


def remember_debug_port(container_code: str, port: int | None) -> None:
    code = (container_code or "").strip()
    if not code or not port:
        return
    with _debug_ports_lock:
        _debug_ports_by_container[code] = int(port)
    register_automation_browser(int(port))


def forget_debug_port(container_code: str) -> None:
    code = (container_code or "").strip()
    if not code:
        return
    with _debug_ports_lock:
        _debug_ports_by_container.pop(code, None)


def get_remembered_debug_port(container_code: str) -> int | None:
    code = (container_code or "").strip()
    if not code:
        return None
    with _debug_ports_lock:
        port = _debug_ports_by_container.get(code)
    if not port:
        return None
    # 端口是否仍可连
    try:
        resp = requests.get(f"http://127.0.0.1:{int(port)}/json/version", timeout=1.5)
        if resp.status_code == 200:
            return int(port)
    except Exception:
        pass
    forget_debug_port(code)
    return None


# 人工验证码等待队列：machine_id -> 会话信息
_pending_codes: dict[str, dict[str, Any]] = {}
_pending_codes_lock = threading.Lock()


def register_pending_recovery_code(
    *,
    machine_id: str,
    container_code: str,
    port: int | None,
    recovery_email: str,
    login_email: str = "",
) -> None:
    mid = (machine_id or "").strip()
    if not mid:
        return
    with _pending_codes_lock:
        _pending_codes[mid] = {
            "machine_id": mid,
            "container_code": (container_code or "").strip(),
            "port": int(port) if port else None,
            "recovery_email": (recovery_email or "").strip(),
            "login_email": (login_email or "").strip(),
            "created_at": time.time(),
        }


def get_pending_recovery_code(machine_id: str) -> dict[str, Any] | None:
    mid = (machine_id or "").strip()
    with _pending_codes_lock:
        row = _pending_codes.get(mid)
        return dict(row) if row else None


def clear_pending_recovery_code(machine_id: str) -> None:
    mid = (machine_id or "").strip()
    with _pending_codes_lock:
        _pending_codes.pop(mid, None)


def apply_manual_recovery_code(
    api_base: str,
    api_key: str,
    login_url: str,
    machine_id: str,
    code: str,
    on_progress: Callable[[str], None] | None = None,
) -> tuple[bool, str, AccountStatus]:
    """
    控制台人工输入验证码后：连回该机子浏览器，填入验证码并提交，
    再自动走 次へ → パスキーキャンセル → はい → Outlook 收件箱。
    """
    from recovery_email import (
        is_code_verify_page,
        submit_verification_code,
    )

    def progress(msg: str) -> None:
        log_step(msg)
        if on_progress:
            try:
                on_progress(msg)
            except Exception:
                pass

    code = (code or "").strip()
    if not code:
        return False, "验证码为空", AccountStatus.WAIT_CODE
    if not re.fullmatch(r"[0-9A-Za-z]{4,12}", code):
        return False, "验证码格式不正确", AccountStatus.WAIT_CODE

    pending = get_pending_recovery_code(machine_id)
    env = resolve_env(api_base, machine_id, api_key)
    container_code = (pending or {}).get("container_code") or (
        env.container_code if env else ""
    )
    if not container_code:
        return False, f"未找到机子号 {machine_id} 的环境", AccountStatus.WAIT_CODE

    recovery_email = (pending or {}).get("recovery_email") or ""

    progress("正在连接浏览器…")
    port = get_remembered_debug_port(container_code)
    if port is None and pending and pending.get("port"):
        try:
            port = int(pending["port"])
        except Exception:
            port = None
    if port is None:
        progress("调试端口丢失，正在重新打开环境…")
        port, err = start_browser(
            api_base,
            container_code,
            login_url,
            api_key,
            force_visible=True,
            open_login_tab=False,
        )
        if port is None:
            return False, err or "无法连接浏览器（环境可能已关闭）", AccountStatus.WAIT_CODE

    remember_debug_port(container_code, port)
    bring_browser_to_foreground(api_base, container_code, api_key)

    playwright = _start_sync_playwright()
    try:
        browser = playwright.chromium.connect_over_cdp(
            f"http://127.0.0.1:{port}",
            timeout=int(_RUNTIME["cdp_connect_timeout_ms"]),
        )
        _cdp_ignore_certificate_errors(browser)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = None
        for candidate in reversed(context.pages):
            try:
                if is_code_verify_page(candidate):
                    page = candidate
                    break
            except Exception:
                continue
        if page is None and context.pages:
            page = context.pages[-1]
        if page is None:
            return False, "浏览器无标签页，无法填验证码", AccountStatus.WAIT_CODE

        try:
            page.bring_to_front()
        except Exception:
            pass

        progress("正在定位验证码页…")
        if not is_code_verify_page(page):
            for _ in range(15):
                if is_code_verify_page(page):
                    break
                safe_wait(page, 300)
        if not is_code_verify_page(page):
            # 可能已经离开验证码页（用户手动点过）：直接尝试收尾登录
            if is_outlook_inbox(page):
                clear_pending_recovery_code(machine_id)
                progress("已在收件箱，正在检查封号邮件…")
                status, detail = evaluate_inbox_account_status(page)
                shot = capture_error_screenshot(page, machine_id, status)
                if shot:
                    detail = f"{detail} | 截图: {shot}"
                return True, detail, status
            progress("未找到验证码页，尝试登录收尾…")
            ok, detail, status = complete_login_after_recovery_code(
                page, timeout_sec=55, on_progress=progress
            )
            if ok:
                clear_pending_recovery_code(machine_id)
                shot = capture_error_screenshot(page, machine_id, status)
                if shot:
                    detail = f"{detail} | 截图: {shot}"
            return ok, detail, status

        hint = f"机子号 {machine_id}"
        if recovery_email:
            hint += f" / 辅助邮箱 {recovery_email}"
        progress(f"正在填写验证码并点击次へ… [{hint}]")
        if not submit_verification_code(page, code, log=log_step):
            return False, "验证码未能写入页面或未能点击次へ", AccountStatus.WAIT_CODE

        safe_wait(page, 450)
        progress("验证码已提交，登录收尾中…")
        ok, detail, status = complete_login_after_recovery_code(
            page, timeout_sec=75, on_progress=progress
        )
        if ok:
            clear_pending_recovery_code(machine_id)
        # 成功或失败都尽量截图；终态只保留最新一张
        shot = capture_error_screenshot(page, machine_id, status)
        detail = re.sub(r"\s*\|\s*截图:\s*\S+", "", detail or "").strip()
        if shot:
            detail = f"{detail} | 截图: {shot}"
        if ok:
            return True, detail, status
        if status == AccountStatus.WAIT_CODE:
            return False, detail, status
        # 超时等非验证码错误：仍清不掉 pending，便于重试或点打开查看
        return False, f"机子号 {machine_id}：{detail}", status
    except Exception as exc:
        return False, f"机子号 {machine_id} 回填验证码异常: {exc}", AccountStatus.WAIT_CODE
    finally:
        try:
            playwright.stop()
        except Exception:
            pass


def set_task_context(
    machine_id: str,
    reporter: Callable[[str, str], None] | None = None,
) -> None:
    _step_ctx.machine_id = machine_id
    _step_ctx.reporter = reporter


def clear_task_context() -> None:
    _step_ctx.machine_id = ""
    _step_ctx.reporter = None


def log_step(msg: str) -> None:
    print(f"    [步骤] {msg}", flush=True)
    reporter = getattr(_step_ctx, "reporter", None)
    machine_id = getattr(_step_ctx, "machine_id", "") or ""
    if reporter and machine_id:
        try:
            reporter(machine_id, msg)
        except Exception:
            pass


def _safe_filename(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\s]+', "_", (text or "unknown").strip())[:80]


def _keep_background() -> bool:
    return bool(_RUNTIME.get("keep_background", True))


def _silent_open_enabled() -> bool:
    """后台静默打开：无界面启动 HubStudio 环境，仅用 CDP 填表。"""
    if _RUNTIME.get("browser_headless"):
        return False
    if not _keep_background():
        return False
    return bool(_RUNTIME.get("browser_silent_open", True))


def _effective_headless() -> bool:
    return bool(_RUNTIME.get("browser_headless") or _silent_open_enabled())


def _build_browser_start_args(*, use_headless: bool, force_visible: bool = False) -> list[str]:
    args: list[str] = []
    if use_headless:
        args.extend(["--headless=new", "--disable-gpu"])
    elif (
        not force_visible
        and _keep_background()
        and _RUNTIME.get("browser_minimize_on_start", True)
    ):
        args.extend(["--start-minimized", "--window-position=-24000,-24000"])
    if _RUNTIME.get("browser_lightweight", True) and not force_visible:
        args.extend(
            [
                "--blink-settings=imagesEnabled=false",
                "--disable-dev-shm-usage",
                "--disable-background-networking",
                "--disable-default-apps",
                "--disable-renderer-backgrounding",
                "--disable-backgrounding-occluded-windows",
            ]
        )
    return args


_MINIMIZE_LOCK = threading.Lock()
_LAST_MINIMIZE_TS = 0.0

# React 兼容的原生 input 赋值（前台模式会 focus）
_JS_SET_NATIVE_VALUE = """(el, v) => {
    if (!el) return false;
    el.focus();
    const proto = window.HTMLInputElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, 'value');
    if (desc && desc.set) desc.set.call(el, v);
    else el.value = v;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    try {
        el.dispatchEvent(new InputEvent('input', {
            bubbles: true, inputType: 'insertText', data: v
        }));
    } catch (e) {}
    return el.value === v;
}"""

# 后台模式：不 focus，避免抢用户前台窗口
_JS_SET_NATIVE_VALUE_NO_FOCUS = """(el, v) => {
    if (!el) return false;
    const proto = window.HTMLInputElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, 'value');
    if (desc && desc.set) desc.set.call(el, v);
    else el.value = v;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    try {
        el.dispatchEvent(new InputEvent('input', {
            bubbles: true, inputType: 'insertText', data: v
        }));
    } catch (e) {}
    return el.value === v;
}"""


def _js_set_native_value_expr() -> str:
    return _JS_SET_NATIVE_VALUE_NO_FOCUS if _keep_background() else _JS_SET_NATIVE_VALUE

_JS_PICK_EMAIL_INPUT = """() => {
    const t = (document.body && document.body.innerText) || '';
    // 辅助邮箱绑定页也有 email 输入框，绝不能当成登录邮箱页
    const recoveryMarkers = [
        'アカウントを保護しましょう', 'アカウントを保護',
        'Help us protect your account', "Let's protect your account",
        'someone@example.com',
        '連絡用メールアドレス', 'セキュリティコードが送信',
    ];
    if (recoveryMarkers.some(p => t.includes(p))) return null;
    const pwdPageMarkers = [
        'パスワードの入力', 'Enter password', '输入密码', '输入你的密码',
        'Microsoft アカウントのパスワード', '请输入你的密码', 'Enter your password'
    ];
    if (pwdPageMarkers.some(p => t.includes(p))) return null;
    for (const sel of ['input[type="password"]', 'input[name="passwd"]', '#i0118', '#passwordEntry']) {
        const el = document.querySelector(sel);
        if (!el) continue;
        const cls = el.className || '';
        if (cls.includes('moveOffScreen')) continue;
        if (el.disabled || el.getAttribute('aria-hidden') === 'true') continue;
        return null;
    }
    // 登录页优先 #i0116 / loginfmt；不要用裸的 type=email（辅助邮箱页会误中）
    const inputs = [...document.querySelectorAll(
        '#i0116, input[name="loginfmt"]'
    )];
    return inputs.find(el => {
        const cls = el.className || '';
        if (cls.includes('moveOffScreen')) return false;
        if (el.disabled || el.getAttribute('aria-hidden') === 'true') return false;
        const ph = (el.getAttribute('placeholder') || '').toLowerCase();
        if (ph.includes('someone@example.com') || ph.includes('example.com')) return false;
        return true;
    }) || null;
}"""

_JS_PICK_PASSWORD_INPUT = """() => {
    const selectors = [
        'input[type="password"]',
        'input[name="passwd"]',
        '#i0118',
        '#passwordEntry',
        'input[autocomplete="current-password"]',
        'input[aria-label*="password" i]',
        'input[aria-label*="パスワード"]',
        'input[placeholder*="password" i]',
        'input[placeholder*="パスワード"]',
    ];
    const seen = new Set();
    const inputs = [];
    function add(el) {
        if (!el || seen.has(el)) return;
        seen.add(el);
        inputs.push(el);
    }
    function walk(node) {
        if (!node) return;
        for (const sel of selectors) {
            try { node.querySelectorAll(sel).forEach(add); } catch (_) {}
        }
        const nodes = node.querySelectorAll ? node.querySelectorAll('*') : [];
        for (const el of nodes) {
            if (el.shadowRoot) walk(el.shadowRoot);
        }
    }
    walk(document);
    const bodyText = ((document.body && document.body.innerText) || '');
    // 仍在「邮箱/电话/Skype」登录首页时，不要把隐藏密码框当成密码页
    const onEmailHome = (
        bodyText.includes('メール、電話、Skype')
        || bodyText.includes('Email, phone, or Skype')
        || bodyText.includes('邮箱、电话')
    ) && !(
        bodyText.includes('パスワードの入力')
        || bodyText.includes('Enter password')
        || bodyText.includes('输入密码')
    );
    if (onEmailHome) return null;

    const scored = inputs
        .filter(el => {
            const cls = el.className || '';
            if (cls.includes('moveOffScreen')) return false;
            if (el.disabled || el.getAttribute('aria-hidden') === 'true') return false;
            try {
                const st = window.getComputedStyle(el);
                if (st.display === 'none' || st.visibility === 'hidden') return false;
            } catch (_) {}
            return true;
        })
        .map(el => {
            const r = el.getBoundingClientRect();
            let score = Math.max(0, r.width) * Math.max(0, r.height);
            if (el.name === 'passwd' || el.id === 'i0118' || el.id === 'passwordEntry') score += 1000;
            if (el.type === 'password') score += 500;
            // 面积为 0 的隐藏框大幅降权（邮箱页常见）
            if (r.width < 2 || r.height < 2) score -= 2000;
            return { el, score };
        })
        .filter(x => x.score > 0)
        .sort((a, b) => b.score - a.score);
    return scored.length ? scored[0].el : null;
}"""


def minimize_browser_windows_debounced(log=log_step) -> None:
    """多开时合并最小化请求，减轻桌面卡顿。"""
    global _LAST_MINIMIZE_TS
    if not _RUNTIME.get("browser_minimize_on_start") or _effective_headless():
        return
    now = time.time()
    with _MINIMIZE_LOCK:
        if now - _LAST_MINIMIZE_TS < 5.0:
            return
        _LAST_MINIMIZE_TS = now
    minimize_hubstudio_windows(log=log)


def minimize_automation_browser_now(log=log_step) -> None:
    """密码页出现后立即最小化 HubStudio，避免页面跳转时窗口弹出抢焦点。"""
    if (
        not _RUNTIME.get("browser_minimize_on_start")
        or _effective_headless()
        or not _keep_background()
    ):
        return
    minimize_hubstudio_windows(log=log)


def read_email_value(page: Page) -> str:
    try:
        val = page.evaluate(
            f"""() => {{
                const pick = {_JS_PICK_EMAIL_INPUT};
                const el = pick();
                return el ? (el.value || '').trim() : '';
            }}"""
        )
        if val is not None:
            return str(val).strip()
    except Exception:
        pass
    loc = find_email_input(page)
    if loc is None:
        return ""
    try:
        return loc.input_value(timeout=600).strip()
    except Exception:
        return ""


def fill_email_via_page_js(page: Page, email: str) -> bool:
    target = email.strip()
    if not target:
        return False
    try:
        return bool(
            page.evaluate(
                f"""(value) => {{
                    const pick = {_JS_PICK_EMAIL_INPUT};
                    const setVal = {_js_set_native_value_expr()};
                    const el = pick();
                    return el ? setVal(el, value) : false;
                }}""",
                target,
            )
        )
    except Exception:
        return False


def fill_email_robust(page: Page, email: str, max_attempts: int = 5) -> bool:
    """邮箱填表：后台 JS 优先，最小化时也能写入。"""
    target = email.strip()
    if not target:
        return False
    bg = _keep_background()
    fast = _fast_fill_enabled()
    if bg:
        pause_ms = 60 if fast else 120
        tries = min(max_attempts, 3)
        for _ in range(tries):
            if fill_email_via_page_js(page, target):
                safe_wait(page, pause_ms)
                if read_email_value(page) == target:
                    return True
            safe_wait(page, pause_ms)
        return read_email_value(page) == target
    ensure_browser_focus_for_input(page)
    js_set = _js_set_native_value_expr()
    for attempt in range(max_attempts):
        if fill_email_via_page_js(page, target):
            safe_wait(page, 200)
            if read_email_value(page) == target:
                return True
        email_loc = find_email_input(page)
        if email_loc is not None:
            if fill_locator(email_loc, target, fast=fast):
                safe_wait(page, 200)
                if read_email_value(page) == target:
                    return True
            try:
                if email_loc.evaluate(js_set, target):
                    safe_wait(page, 200)
                    if read_email_value(page) == target:
                        return True
            except Exception:
                pass
        safe_wait(page, 250 + attempt * 150)
    return read_email_value(page) == target


def read_password_value(page: Page) -> str:
    try:
        val = page.evaluate(
            f"""() => {{
                const pick = {_JS_PICK_PASSWORD_INPUT};
                const el = pick();
                return el ? (el.value || '').trim() : '';
            }}"""
        )
        if val:
            return str(val).strip()
    except Exception:
        pass
    loc = find_password_input(page)
    if loc is None:
        return ""
    try:
        return loc.input_value(timeout=600).strip()
    except Exception:
        return ""


def fill_password_via_page_js(page: Page, password: str) -> bool:
    target = password.strip()
    if not target:
        return False
    try:
        return bool(
            page.evaluate(
                f"""(value) => {{
                    const pick = {_JS_PICK_PASSWORD_INPUT};
                    const setVal = {_js_set_native_value_expr()};
                    const el = pick();
                    return el ? setVal(el, value) : false;
                }}""",
                target,
            )
        )
    except Exception:
        return False


def click_login_submit(page: Page, next_selectors: list[str] | None = None) -> bool:
    next_selectors = next_selectors or [
        '#idSIButton9',
        'input[type="submit"]',
        'button[type="submit"]',
    ]
    if _keep_background():
        try:
            if page.evaluate(
                """() => {
                    const btn = document.querySelector('#idSIButton9')
                        || document.querySelector('input[type="submit"]')
                        || document.querySelector('button[type="submit"]');
                    if (btn) {
                        btn.click();
                        return true;
                    }
                    const form = document.querySelector('form');
                    if (form && form.requestSubmit) {
                        form.requestSubmit();
                        return true;
                    }
                    return false;
                }"""
            ):
                return True
        except Exception:
            pass
    try:
        page.locator('#idSIButton9').first.click(force=True, timeout=1500)
        return True
    except Exception:
        return wait_and_click(page, next_selectors, timeout_ms=800)


def init_screenshot_session(base_dir: Path | None = None) -> Path:
    """为本次批量任务创建截图目录（程序目录/imgs/时间戳）。"""
    root = base_dir or data_dir() / "imgs"
    session = root / datetime.now().strftime("%Y%m%d_%H%M%S")
    session.mkdir(parents=True, exist_ok=True)
    _RUNTIME["screenshot_dir"] = str(session)
    return session


def should_capture_screenshot(status: AccountStatus) -> bool:
    if not _RUNTIME.get("screenshot_enabled", True):
        return False
    # 登入 / 封号 / 已绑辅助邮箱 / 无法绑定 等都要截图
    return True


def _page_ready_for_screenshot(page: Page) -> bool:
    """页面已有可辨认内容（非空白壳），再截图。"""
    try:
        return bool(
            page.evaluate(
                """() => {
                    const norm = (s) => (s || '').replace(/\\s+/g, ' ').trim();
                    const t = norm(document.body && document.body.innerText);
                    const selectors = [
                        '#i0116', '#usernameEntry', 'input[name="loginfmt"]',
                        'input[name="passwd"]', '#passwordEntry', 'input[type="password"]',
                        '#idSIButton9', 'input[type="submit"]', 'button[type="submit"]',
                    ];
                    if (selectors.some(s => document.querySelector(s))) return true;
                    const markers = [
                        'Microsoft', 'Sign in', 'サインイン', 'パスワード', 'password',
                        '邮箱', '账户', 'アカウント', 'ログイン', 'login', '错误', '错误',
                        'locked', 'ロック', 'verify', '確認', 'コード',
                        'Authenticator', '辅助邮箱', '绑定', 'Verify your email',
                        'Check your email', 'メールをご確認', 'サインイン要求',
                        'Send code', 'コードの送信', 'someone@example.com',
                        'ERR_CONNECTION', 'このサイトにアクセスできません',
                        "can't be reached", '无法访问此网站',
                        'Inbox', '受信トレイ', 'Focused', 'Outlook', 'Mail',
                    ];
                    if (markers.some(m => t.includes(m)) && t.length > 35) return true;
                    if (/outlook\\.(live|office)\\.com/i.test(location.href) && t.length > 40) return true;
                    return t.length > 90;
                }"""
            )
        )
    except Exception:
        return False


def wait_for_screenshot_ready(page: Page, timeout_sec: float = 10.0) -> bool:
    """等待登录/错误页渲染完成，避免截到空白壳。"""
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if _page_ready_for_screenshot(page):
            safe_wait(page, 250)
            if _page_ready_for_screenshot(page):
                return True
        try:
            page.wait_for_load_state("domcontentloaded", timeout=800)
        except Exception:
            pass
        time.sleep(0.15)
    return _page_ready_for_screenshot(page)


def _screenshot_is_blank(path: Path, min_bytes: int = 8192) -> bool:
    try:
        return path.stat().st_size < min_bytes
    except OSError:
        return True


def _screenshot_via_cdp(page: Page, path: Path) -> bool:
    """最小化时 Playwright 截图可能空白，CDP 兜底。"""
    try:
        client = page.context.new_cdp_session(page)
        result = client.send(
            "Page.captureScreenshot",
            {"format": "png", "captureBeyondViewport": True, "fromSurface": True},
        )
        path.write_bytes(base64.b64decode(result["data"]))
        client.detach()
        return not _screenshot_is_blank(path)
    except Exception:
        return False


def _take_screenshot_once(page: Page, path: Path, *, prefer_cdp: bool) -> bool:
    if prefer_cdp:
        if _screenshot_via_cdp(page, path) and not _screenshot_is_blank(path):
            return True
    try:
        page.screenshot(
            path=str(path),
            full_page=False,
            animations="disabled",
            timeout=15000,
        )
        if not _screenshot_is_blank(path):
            return True
    except Exception:
        pass
    if not prefer_cdp:
        if _screenshot_via_cdp(page, path) and not _screenshot_is_blank(path):
            return True
    try:
        page.screenshot(path=str(path), full_page=True, animations="disabled", timeout=15000)
        if not _screenshot_is_blank(path):
            return True
    except Exception:
        pass
    try:
        page.locator("body").screenshot(path=str(path), timeout=10000)
        return not _screenshot_is_blank(path)
    except Exception:
        return False


def resolve_page_for_screenshot(page: Page | None, status: AccountStatus) -> Page | None:
    """截图前定位到实际展示异常/验证/邮箱内容的标签页。"""
    if page is None:
        return None
    try:
        context = page.context
    except Exception:
        return page

    prefer_inbox = status in {
        AccountStatus.LOGIN_OK,
        AccountStatus.OK,
        AccountStatus.AMZ_BANNED,
    }
    prefer_bound = status == AccountStatus.ALREADY_BOUND
    try:
        for candidate in reversed(context.pages):
            try:
                if prefer_inbox and is_outlook_inbox(candidate):
                    # 垃圾箱/搜索页不算合格终态视图
                    u = (candidate.url or "").lower()
                    if "junk" in u or "/search" in u:
                        continue
                    return candidate
                if prefer_bound and detect_already_bound_email(candidate):
                    return candidate
                if is_recovery_bind_page(candidate):
                    return candidate
                # 密码页优先于「网卡」误判页，避免截到错页却标网卡
                if is_password_entry_page(candidate) or find_password_input(candidate):
                    return candidate
                if detect_login_errors(candidate):
                    return candidate
            except Exception:
                continue
    except Exception:
        pass
    return page


def capture_error_screenshot(page: Page, machine_id: str, status: AccountStatus) -> str:
    """遇到异常页面时截图，返回相对项目根目录的路径。"""
    if not should_capture_screenshot(status):
        return ""
    page = resolve_page_for_screenshot(page, status) or page
    if status in {
        AccountStatus.LOGIN_OK,
        AccountStatus.OK,
    }:
        try:
            ensure_outlook_inbox_for_screenshot(page)
        except Exception:
            pass
    shot_dir_str = (_RUNTIME.get("screenshot_dir") or "").strip()
    if not shot_dir_str:
        shot_dir = init_screenshot_session()
    else:
        shot_dir = Path(shot_dir_str)
        shot_dir.mkdir(parents=True, exist_ok=True)
    fname = (
        f"{_safe_filename(machine_id)}_{_safe_filename(status.value)}_"
        f"{datetime.now().strftime('%H%M%S')}.png"
    )
    path = shot_dir / fname
    need_reminimize = False
    background = _keep_background()
    prefer_cdp = background or _effective_headless()
    inbox_result = status in {
        AccountStatus.LOGIN_OK,
        AccountStatus.OK,
        AccountStatus.AMZ_BANNED,
    }
    min_bytes = 3500 if (status == AccountStatus.ALREADY_BOUND or inbox_result) else 8192
    relax_ready = status == AccountStatus.ALREADY_BOUND or inbox_result
    max_wait = float(_RUNTIME.get("screenshot_max_wait_sec", 8.0))
    if inbox_result:
        max_wait = min(max_wait, 3.5)
    try:
        if is_microsoft_login_page_loading(page):
            wait_for_login_transition(page, timeout_sec=min(8.0, max_wait))
        wait_for_screenshot_ready(
            page,
            timeout_sec=min(5.0 if inbox_result else (14.0 if relax_ready else 12.0), max_wait + (2.0 if inbox_result else 4.0)),
        )

        if not background and not _effective_headless():
            reveal_hubstudio_windows_for_screenshot()
            need_reminimize = bool(_RUNTIME.get("browser_minimize_on_start"))
            try:
                page.bring_to_front()
            except Exception:
                pass
            safe_wait(page, 300 if inbox_result else 600)
            wait_for_screenshot_ready(page, timeout_sec=2.5 if inbox_result else (5.0 if relax_ready else 4.0))

        captured = False
        attempts = 3 if inbox_result else (6 if relax_ready else 5)
        for attempt in range(attempts):
            if attempt:
                wait_for_screenshot_ready(page, timeout_sec=(1.2 if inbox_result else 2.5) + attempt * 0.5)
                safe_wait(page, (200 if inbox_result else 400) + attempt * 150)
            if _take_screenshot_once(page, path, prefer_cdp=prefer_cdp):
                ready_ok = _page_ready_for_screenshot(page)
                if relax_ready or ready_ok or attempt >= 2:
                    if not _screenshot_is_blank(path, min_bytes=min_bytes):
                        captured = True
                        break
                safe_wait(page, 250 if inbox_result else 500)

        project_root = data_dir()
        try:
            rel = path.relative_to(project_root).as_posix()
        except ValueError:
            rel = path.as_posix()
        if not captured or _screenshot_is_blank(path, min_bytes=min_bytes):
            log_step(f"截图可能未加载完成 [{machine_id}]: {rel}")
        else:
            log_step(f"截图 [{machine_id}]: {rel}")
        return rel if path.exists() and not _screenshot_is_blank(path, min_bytes=min_bytes) else ""
    except Exception as exc:
        log_step(f"截图失败 [{machine_id}]: {exc}")
        return ""
    finally:
        if need_reminimize:
            safe_wait(page, 200)
            minimize_browser_windows_debounced()


def finalize_check_result(
    *,
    machine_id: str,
    email: str,
    status: AccountStatus,
    detail: str,
    container_code: str = "",
    final_url: str = "",
    checked_at: str = "",
    page: Page | None = None,
    recovery_email: str = "",
    awaiting_code: bool = False,
) -> CheckResult:
    awaiting = awaiting_code or status == AccountStatus.WAIT_CODE
    if page is not None and should_capture_screenshot(status):
        return make_check_result(
            machine_id=machine_id,
            email=email,
            status=status,
            detail=detail,
            container_code=container_code,
            final_url=final_url,
            checked_at=checked_at,
            page=page,
            recovery_email=recovery_email,
            awaiting_code=awaiting,
        )
    return CheckResult(
        machine_id=machine_id,
        email=email,
        status=status,
        detail=detail,
        container_code=container_code,
        final_url=final_url,
        checked_at=checked_at,
        recovery_email=recovery_email,
        awaiting_code=awaiting,
    )


def make_check_result(
    *,
    machine_id: str,
    email: str,
    status: AccountStatus,
    detail: str,
    container_code: str = "",
    final_url: str = "",
    checked_at: str = "",
    page: Page | None = None,
    recovery_email: str = "",
    awaiting_code: bool = False,
) -> CheckResult:
    # 最终状态只保留一张终态截图，避免「等待验证码」旧图叠在封号结果上
    detail = re.sub(r"\s*\|\s*截图:\s*\S+", "", detail or "").strip()
    screenshot_path = ""
    if page is not None:
        screenshot_path = capture_error_screenshot(page, machine_id, status)
    if screenshot_path:
        detail = f"{detail} | 截图: {screenshot_path}"
    return CheckResult(
        machine_id=machine_id,
        email=email,
        status=status,
        detail=detail,
        container_code=container_code,
        final_url=final_url,
        checked_at=checked_at,
        screenshot_path=screenshot_path,
        recovery_email=recovery_email,
        awaiting_code=awaiting_code or status == AccountStatus.WAIT_CODE,
    )


def stop_browser(api_base: str, container_code: str, api_key: str) -> tuple[bool, str]:
    payload = {"containerCode": container_code}
    try:
        data = hubstudio_post(api_base, "browser/stop", payload, api_key)
        if data.get("code") == 0:
            return True, "已关闭"
        return False, str(data.get("msg") or data)
    except Exception as exc:
        return False, str(exc)


def stop_all_browsers(
    api_base: str, api_key: str, *, clear_opening: bool = True
) -> tuple[bool, str]:
    """关闭 HubStudio 全部已打开环境（含无头/静默模式）。"""
    payload = {"clearOpening": clear_opening}
    try:
        data = hubstudio_post(api_base, "browser/stop-all", payload, api_key)
        if data.get("code") == 0:
            detail = (data.get("data") or {}).get("err") or "已关闭全部环境"
            return True, str(detail)
        return False, str(data.get("msg") or data)
    except Exception as exc:
        return False, str(exc)


def stop_browsers_by_machine_ids(
    api_base: str, machine_ids: list[str], api_key: str
) -> list[dict[str, str | bool]]:
    rows: list[dict[str, str | bool]] = []
    for machine_id in machine_ids:
        mid = (machine_id or "").strip()
        if not mid:
            continue
        env = resolve_env(api_base, mid, api_key)
        if env is None:
            rows.append({"machine_id": mid, "ok": False, "detail": "未找到 HubStudio 环境"})
            continue
        ok, detail = stop_browser(api_base, env.container_code, api_key)
        rows.append(
            {
                "machine_id": mid,
                "container_code": env.container_code,
                "ok": ok,
                "detail": detail,
            }
        )
    return rows


# 这些状态保持浏览器打开，便于手动绑定或排查
_KEEP_BROWSER_OPEN_STATUSES = frozenset(
    {
        AccountStatus.NEED_RECOVERY,
        AccountStatus.WAIT_CODE,
        AccountStatus.UNKNOWN,
        AccountStatus.TIMEOUT,
        AccountStatus.NETWORK_CARD,
        AccountStatus.RECOVERY_FAILED,
        AccountStatus.CAPTCHA,
        AccountStatus.NEED_VERIFY,
        AccountStatus.NEED_PHONE,
        AccountStatus.NEED_IDENTITY,
        AccountStatus.LOCKED,
        AccountStatus.STAY_SIGNED_IN,
        AccountStatus.OPEN_FAILED,
    }
)

# 登录后不应进入辅助邮箱绑定流程的状态
_NON_RECOVERY_FLOW_STATUSES = frozenset(
    {
        AccountStatus.ACCOUNT_NOT_FOUND,
        AccountStatus.BAD_PASSWORD,
        AccountStatus.LOCKED,
        AccountStatus.BAD_CREDENTIALS,
        AccountStatus.ALREADY_BOUND,
        AccountStatus.NEED_IDENTITY,
        AccountStatus.NEED_PHONE,
        AccountStatus.NEED_VERIFY,
        AccountStatus.NETWORK_CARD,
        AccountStatus.UNKNOWN,
        AccountStatus.CAPTCHA,
        AccountStatus.LOGIN_OK,
        AccountStatus.AMZ_BANNED,
        AccountStatus.OK,
    }
)


def should_close_browser_after_check(status: AccountStatus, *, enabled: bool = True) -> bool:
    """已判定的正常结果可关闭；需要绑定辅助邮箱及异常状况保持打开。"""
    if not enabled:
        return False
    if _effective_headless():
        # 无头/静默模式无法在桌面操作，除需手动绑定的流程外一律关闭
        if status in {
            AccountStatus.NEED_RECOVERY,
            AccountStatus.RECOVERY_FAILED,
            AccountStatus.WAIT_CODE,
        }:
            return False
        return True
    return status not in _KEEP_BROWSER_OPEN_STATUSES


def maybe_stop_browser(
    api_base: str,
    container_code: str,
    api_key: str,
    status: AccountStatus,
    close_enabled: bool,
) -> None:
    if should_close_browser_after_check(status, enabled=close_enabled):
        ok, detail = stop_browser(api_base, container_code, api_key)
        forget_debug_port(container_code)
        if ok:
            log_step(f"已关闭 HubStudio 环境 ({container_code})")
        else:
            log_step(f"关闭环境失败 ({container_code}): {detail}")


def wait_and_click(page: Page, selectors: list[str], timeout_ms: int = 1500) -> bool:
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=timeout_ms)
            loc.click()
            return True
        except PlaywrightTimeout:
            continue
        except Exception:
            continue
    return False


def fill_first(page: Page, selectors: list[str], value: str, timeout_ms: int = 8000) -> bool:
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=timeout_ms)
            loc.fill(value)
            return True
        except PlaywrightTimeout:
            continue
        except Exception:
            continue
    return False


def page_text(page: Page) -> str:
    """读取可见文案；Microsoft 表单常在 iframe，需合并各 frame 文本。"""
    parts: list[str] = []
    try:
        text = page.evaluate("() => (document.body && document.body.innerText) || ''")
        if text:
            parts.append(str(text))
    except Exception:
        try:
            t = page.inner_text("body", timeout=2000)
            if t:
                parts.append(t)
        except Exception:
            pass
    try:
        for frame in page.frames:
            try:
                if frame == page.main_frame:
                    continue
            except Exception:
                continue
            try:
                t = frame.evaluate("() => (document.body && document.body.innerText) || ''")
                if t and str(t).strip():
                    parts.append(str(t))
            except Exception:
                try:
                    t = frame.inner_text("body", timeout=800)
                    if t and str(t).strip():
                        parts.append(str(t))
                except Exception:
                    continue
    except Exception:
        pass
    return "\n".join(parts)


def page_text_lower(page: Page) -> str:
    return page_text(page).lower()


def detect_identity_verification(page: Page) -> tuple[AccountStatus, str] | None:
    """Microsoft 异常活动 / 本人确认页（非辅助邮箱绑定）。"""
    if not is_identity_verification_page(page):
        return None
    return AccountStatus.NEED_IDENTITY, "无法绑定辅助邮箱（Microsoft 检测到异常活动，要求本人确认）"


def detect_network_card_error(page: Page) -> tuple[AccountStatus, str] | None:
    """Chrome/Edge 连接中断页：代理/IP 网卡异常，应更换 IP。"""
    # 已出现 Microsoft 登录表单（含密码页）→ 绝不是网卡错误页
    try:
        if is_password_entry_page(page) or find_password_input(page) is not None:
            return None
    except Exception:
        pass
    try:
        if find_email_input(page) is not None:
            # 有邮箱框且 URL 仍是登录域，不算网卡（半加载也优先继续登录）
            url0 = (page.url or "").lower()
            if any(
                h in url0
                for h in (
                    "login.microsoftonline.com",
                    "login.live.com",
                    "account.live.com",
                    "signup.live.com",
                )
            ):
                return None
    except Exception:
        pass
    try:
        if is_recovery_bind_page(page) or is_code_verify_page(page):
            return None
    except Exception:
        pass

    text = page_text(page)
    text_l = text.lower()
    # 密码/保护账户正文出现时，即便某 iframe 残留错误词也不判网卡
    if any(
        m in text or m in text_l
        for m in (
            "パスワードの入力",
            "enter password",
            "アカウントを保護",
            "protect your account",
            "someone@example.com",
        )
    ):
        return None

    network_markers = [
        "ERR_CONNECTION_CLOSED",
        "ERR_CONNECTION_RESET",
        "ERR_CONNECTION_TIMED_OUT",
        "ERR_CONNECTION_REFUSED",
        "ERR_NAME_NOT_RESOLVED",
        "ERR_TUNNEL_CONNECTION_FAILED",
        "ERR_PROXY_CONNECTION_FAILED",
        "ERR_NETWORK_CHANGED",
        "ERR_INTERNET_DISCONNECTED",
        "ERR_SSL_PROTOCOL_ERROR",
        "このサイトにアクセスできません",
        "により途中で接続が切断されました",
        "接続がタイムアウトしました",
        "接続がリセットされました",
        "This site can’t be reached",
        "This site can't be reached",
        "took too long to respond",
        "connection was reset",
        "Unexpectedly closed the connection",
        "无法访问此网站",
        "连接已重置",
        "连接超时",
        "连接被意外终止",
        "网关超时",
    ]
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    # 真正的 Chrome 错误页 URL，或正文含明确 ERR_/无法访问 标记
    chrome_error_url = "chrome-error://" in url or url.startswith("chrome://")
    marker_hit = any(m.lower() in text_l or m in text for m in network_markers)
    if not marker_hit and not chrome_error_url:
        return None
    if marker_hit:
        return AccountStatus.NETWORK_CARD, "应更换IP"

    # Chrome 错误页结构兜底：标题含无法访问 + live/microsoft 域名
    host_hit = any(
        h in url or h in text_l
        for h in (
            "account.live.com",
            "login.live.com",
            "login.microsoftonline.com",
            "outlook.live.com",
        )
    )
    chrome_fail = any(
        x in text
        for x in (
            "このサイトにアクセスできません",
            "无法访问此网站",
            "This site can't be reached",
            "This site can’t be reached",
        )
    )
    if host_hit and chrome_fail:
        return AccountStatus.NETWORK_CARD, "应更换IP"
    return None


def detect_login_errors(page: Page) -> tuple[AccountStatus, str] | None:
    """识别 Microsoft 登录常见错误（含日文锁定页、Abuse URL）。"""
    identity = detect_identity_verification(page)
    if identity:
        return identity

    try:
        if is_phone_verify_page(page):
            return AccountStatus.NEED_PHONE, "需要电话认证（确认手机号末位，非辅助邮箱流程）"
    except Exception:
        pass

    network = detect_network_card_error(page)
    if network:
        return network

    try:
        url = page.url.lower()
    except Exception:
        url = ""

    # URL 优先：账户锁定页、滥用拦截页
    if any(x in url for x in ["/abuse", "account.live.com/abuse", "account.microsoft.com/abuse"]):
        return AccountStatus.LOCKED, "账户被锁"

    text = page_text(page)

    locked_patterns = [
        "ご使用のアカウントがロックされました",
        "Microsoft サービス規約に違反するアクティビティ",
        "アカウントがロックされました",
        "アカウントがロック",
        "ロックされました",
        "account has been locked",
        "your account has been locked",
        "帐户已被锁定",
        "账号已被锁定",
    ]
    if any(p in text for p in locked_patterns):
        return AccountStatus.LOCKED, "账户被锁"
    for snippet in locked_patterns[:5]:
        try:
            if page.get_by_text(snippet, exact=False).first.is_visible(timeout=400):
                return AccountStatus.LOCKED, "账户被锁"
        except Exception:
            pass

    not_found_patterns = [
        "そのユーザー名のアカウントが見つかりませんでした",
        "That Microsoft account doesn't exist",
        "that microsoft account doesn't exist",
        "We couldn't find an account",
        "couldn't find a microsoft account",
        "找不到此 Microsoft 帐户",
        "找不到此帐户",
    ]
    if any(p in text for p in not_found_patterns):
        return AccountStatus.ACCOUNT_NOT_FOUND, "找不到账户"

    password_patterns = [
        "正しくないアカウントまたはパスワード",
        "パスワードでのサインインは使用できません",
        "Sign-in with password is not available",
        "sign-in with password isn't available",
        "You can't sign in with your password",
        "incorrect password",
        "password is incorrect",
        "Your account or password is incorrect",
        "密码不正确",
        "密码错误",
        "无法使用密码登录",
        "不能使用密码登录",
        "所定の回数を超えました",
    ]
    if any(p in text for p in password_patterns):
        return AccountStatus.BAD_PASSWORD, "密码错误"
    for snippet in password_patterns[:6]:
        try:
            if page.get_by_text(snippet, exact=False).first.is_visible(timeout=400):
                return AccountStatus.BAD_PASSWORD, "密码错误"
        except Exception:
            pass

    return None


def is_signin_request_verify_page(page: Page) -> bool:
    """邮箱提交后出现「向设备发送登录请求 / Authenticator」页，表示已绑定安全验证。"""
    markers = [
        "サインイン要求を取得する",
        "要求をデバイスに送信",
        "要求を送信する",
        "Get a sign-in request",
        "Send a request to your device",
        "Send request",
        "Approve sign in",
        "Approve sign-in",
        "Microsoft Authenticator",
        "その他のサインイン方法",
        "Other sign-in methods",
        "Use your face, fingerprint",
        "We'll send a sign-in request",
        "发送登录请求",
        "向你的设备发送",
    ]
    try:
        found = page.evaluate(
            """(markers) => {
                const t = (document.body && document.body.innerText) || '';
                return markers.some(p => t.includes(p));
            }""",
            markers,
        )
        if found:
            return True
    except Exception:
        pass
    text = page_text(page)
    return any(p in text for p in markers)


def is_already_bound_email_page(page: Page) -> bool:
    """
    输入登录邮箱后出现「确认已绑定辅助邮箱」页：
    例如日文「メールをご確認ください」+ 遮罩邮箱 yk*****@xxx +「コードの送信」。
    这表示辅助邮箱已绑定，不是未知状态，也不是「需要绑定辅助邮箱」。
    """
    text = page_text(page)
    text_lower = text.lower()

    # 强特征：确认已有辅助邮箱（优先于 recovery 绑定页判断）
    strong_jp = (
        "メールをご確認ください" in text
        or "コードの送信" in text
        or "ご自身のものであることを確認" in text
    )
    strong_en = (
        "please check your email" in text_lower
        or "verify your email" in text_lower
        or "send code" in text_lower
        or "confirm it's yours" in text_lower
        or "confirm it is yours" in text_lower
    )
    strong_zh = (
        "请查看你的邮箱" in text
        or "验证你的电子邮件" in text
        or "发送代码" in text
        or "发送验证码" in text and "确认" in text
    )
    masked = bool(
        re.search(r"[A-Za-z0-9._%+-]*\*+[A-Za-z0-9._%+-]*@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    )
    if (strong_jp or strong_en or strong_zh) and (
        masked
        or "ご自身のものであることを確認" in text
        or "confirm it's yours" in text_lower
        or "コードの送信" in text
        or "send code" in text_lower
    ):
        # 排除「去绑定新辅助邮箱」保护账户页
        if "アカウントを保護しましょう" in text or "help us protect your account" in text_lower:
            if "someone@example.com" in text_lower:
                return False
        return True

    # 设备/Authenticator 确认也算已绑安全信息
    if is_signin_request_verify_page(page):
        return True

    # 真正的「需要新绑辅助邮箱」页，不算已绑定
    if is_recovery_bind_page(page):
        return False

    bound_patterns = [
        "メールをご確認ください",
        "Please check your email",
        "Check your email",
        "コードの送信",
        "Send code",
        "にコードを送信します",
        "ご自身のものであることを確認",
        "confirm it's yours",
        "Verify your email",
        "验证你的电子邮件",
        "请查看你的邮箱",
        "既にコードを受け取りましたか",
    ]
    return any(p in text or p.lower() in text_lower for p in bound_patterns)


def detect_already_bound_email(page: Page) -> tuple[AccountStatus, str] | None:
    if is_signin_request_verify_page(page):
        return (
            AccountStatus.ALREADY_BOUND,
            "已绑定辅助邮箱（Microsoft 要求设备/Authenticator 确认登录）",
        )
    if is_already_bound_email_page(page):
        return (
            AccountStatus.ALREADY_BOUND,
            "已绑定辅助邮箱（Microsoft 要求验证已绑定的辅助邮箱）",
        )
    return None


def wait_and_detect_error(page: Page, timeout_sec: float = 3.0) -> tuple[AccountStatus, str] | None:
    """短暂等待并检测是否出现错误页或已绑邮箱页。"""
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        err = detect_login_errors(page)
        if err:
            return err
        bound = detect_already_bound_email(page)
        if bound:
            return bound
        time.sleep(0.2)
    return None


def is_microsoft_login_page_loading(page: Page) -> bool:
    """Microsoft 登录页只渲染邮箱/Logo 壳，密码框或按钮尚未出现。"""
    try:
        # 辅助邮箱绑定表单已出（常在 iframe）→ 不算加载中
        if find_recovery_email_input(page) is not None:
            return False
    except Exception:
        pass
    try:
        blob = page_text(page) or ""
        if any(
            m in blob
            for m in (
                "アカウントを保護しましょう",
                "アカウントを保護",
                "アカウントの保護",
                "help us protect your account",
                "let's protect your account",
                "someone@example.com",
            )
        ):
            return False
    except Exception:
        pass
    try:
        if is_ms_auth_page_shell(page):
            return True
        return bool(
            page.evaluate(
                f"""() => {{
                    const url = (location.href || '').toLowerCase();
                    if (!['login.microsoftonline.com', 'login.live.com', 'account.live.com']
                        .some(h => url.includes(h))) {{
                        return false;
                    }}
                    const pickPwd = {_JS_PICK_PASSWORD_INPUT};
                    if (pickPwd()) return false;
                    const pickEmail = {_JS_PICK_EMAIL_INPUT};
                    if (pickEmail()) return false;
                    const t = ((document.body && document.body.innerText) || '')
                        .replace(/\\s+/g, ' ').trim();
                    const settledMarkers = [
                        'incorrect', '正しくない', '見つかりません', "doesn't exist",
                        'locked', 'ロック', '错误', 'password is', 'パスワードでのサインイン',
                        'verify', '確認', 'コード', 'Authenticator', 'Stay signed in',
                        '保持登录', 'サインイン要求', 'protect your account', 'アカウントを保護',
                        'パスワードの入力', 'Enter password', '输入密码',
                        'ERR_CONNECTION', 'このサイトにアクセスできません',
                        "can't be reached", '无法访问此网站',
                        'someone@example.com',
                    ];
                    if (settledMarkers.some(m => t.toLowerCase().includes(m.toLowerCase()))) {{
                        return false;
                    }}
                    const hasSubmit = !!document.querySelector(
                        '#idSIButton9, input[type="submit"]:not([disabled]), button[type="submit"]:not([disabled])'
                    );
                    const hasAccountBanner = !!document.querySelector(
                        '[data-testid="identityBanner"], #displayName, #identityBanner'
                    );
                    const hasEmailInText = /@\\S+\\.\\S+/.test(t);
                    const hasMicrosoft = /microsoft/i.test(t);
                    const visibleInputs = Array.from(
                        document.querySelectorAll('input:not([type=hidden])')
                    ).filter(el => {{
                        try {{
                            const r = el.getBoundingClientRect();
                            return r.width > 0 && r.height > 0 && el.offsetParent !== null;
                        }} catch (e) {{ return false; }}
                    }});
                    if ((hasAccountBanner || hasEmailInText) && !hasSubmit
                        && visibleInputs.length === 0 && t.length < 320) {{
                        return true;
                    }}
                    if (hasMicrosoft && t.length < 120 && !hasSubmit
                        && visibleInputs.length === 0) return true;
                    return false;
                }}"""
            )
        )
    except Exception:
        return False


def wait_for_login_transition(page: Page, timeout_sec: float = 12.0) -> bool:
    """等待登录页从加载壳过渡到可操作/可判定状态。"""
    deadline = time.time() + timeout_sec
    poll = 0.1 if (_keep_background() or _effective_headless()) else 0.18
    while time.time() < deadline:
        if not is_microsoft_login_page_loading(page):
            return True
        if detect_login_errors(page):
            return True
        if password_field_ready(page, None):
            return True
        url = page.url.lower()
        if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
            return True
        if detect_already_bound_email(page):
            return True
        if is_recovery_bind_page(page):
            return True
        try:
            page.wait_for_load_state("domcontentloaded", timeout=900)
        except Exception:
            pass
        time.sleep(poll)
    return not is_microsoft_login_page_loading(page)


def wait_for_post_login_outcome(
    page: Page,
    timeout_sec: float = 15.0,
    *,
    email: str = "",
) -> tuple[AccountStatus, str]:
    """密码提交后等待页面结果：邮箱/锁定/错误/绑定辅助邮箱等。"""
    log_step("等待登录结果")
    deadline = time.time() + timeout_sec
    bg = _keep_background() or _effective_headless()
    poll = 0.08 if bg else 0.2
    last_log = 0.0
    last_resubmit = 0.0
    loading_grace = 0.0
    max_loading_grace = 18.0 if bg else 12.0
    email_stuck_since: float | None = None
    did_email_refresh = False

    while time.time() < deadline + loading_grace:
        if is_microsoft_login_page_loading(page):
            loading_grace = min(max_loading_grace, loading_grace + 1.2)
            now = time.time()
            if now - last_log > 2.0:
                log_step("登录页仍在加载，继续等待…")
                last_log = now
            time.sleep(poll)
            continue

        # 网卡回到邮箱输入页：停留片刻后刷新两下再继续
        if is_stuck_on_login_email_page(page):
            now = time.time()
            if email_stuck_since is None:
                email_stuck_since = now
            elif (not did_email_refresh) and (now - email_stuck_since >= 2.0):
                recover_network_stuck_email_page(page, email=email, refreshes=2)
                did_email_refresh = True
                email_stuck_since = None
                # 刷新后若仍在邮箱页，交给上层 auto_login 重跑；此处先继续观察
                time.sleep(poll)
                continue
            time.sleep(poll)
            continue
        else:
            email_stuck_since = None

        err = detect_login_errors(page)
        if err:
            return err
        bound = detect_already_bound_email(page)
        if bound:
            return bound

        if is_password_entry_page(page) and is_empty_password_error(page):
            return AccountStatus.UNKNOWN, "未能填入密码（页面提示密码为空）"

        url = page.url.lower()
        if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
            return evaluate_inbox_account_status(page)

        if is_recovery_bind_page(page):
            if is_ms_auth_page_shell(page) or is_microsoft_login_page_loading(page):
                time.sleep(poll)
                continue
            return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"

        text_lower = page_text_lower(page)
        if is_stay_signed_in_page(page) or any(
            x in text_lower for x in ["stay signed in", "保持登录", "keep me signed in"]
        ):
            click_stay_signed_in(page)
            safe_wait(page, 800 if bg else 1500)
            continue

        if is_passkey_setup_page(page):
            dismiss_passkey_setup(page)
            safe_wait(page, 800 if bg else 1200)
            continue

        status, detail = detect_status(page)
        if status not in {AccountStatus.UNKNOWN, AccountStatus.STAY_SIGNED_IN}:
            return status, detail
        if status == AccountStatus.STAY_SIGNED_IN:
            click_stay_signed_in(page)
            safe_wait(page, 800 if bg else 1500)
            continue

        if (
            bg
            and is_password_entry_page(page)
            and read_password_value(page)
            and time.time() - last_resubmit > 2.0
        ):
            click_login_submit(page)
            last_resubmit = time.time()

        now = time.time()
        if now - last_log > 2.0:
            short_url = url.split("?")[0][-48:] if url else "加载中"
            log_step(f"等待登录结果… {short_url}")
            last_log = now

        time.sleep(poll)

    err = detect_login_errors(page)
    if err:
        return err
    bound = detect_already_bound_email(page)
    if bound:
        return bound

    if is_microsoft_login_page_loading(page):
        log_step("登录页仍在加载，额外等待过渡…")
        wait_for_login_transition(page, timeout_sec=16.0 if bg else 12.0)

    url = page.url.lower()
    if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
        return evaluate_inbox_account_status(page)
    if is_recovery_bind_page(page) or _page_looks_like_recovery_bind(page):
        if is_ms_auth_page_shell(page) or is_microsoft_login_page_loading(page):
            # 绑定页壳还在加载：再等一会，不要立刻超时
            wait_for_login_transition(page, timeout_sec=10.0)
            if _page_looks_like_recovery_bind(page) and not is_ms_auth_page_shell(page):
                return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"
            if is_ms_auth_page_shell(page) or is_microsoft_login_page_loading(page):
                return AccountStatus.UNKNOWN, "绑定页仍在加载中"
        return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"

    if is_stuck_on_login_email_page(page) and not did_email_refresh:
        recover_network_stuck_email_page(page, email=email, refreshes=2)
        did_email_refresh = True
        if is_stuck_on_login_email_page(page):
            return AccountStatus.UNKNOWN, "网卡卡在邮箱输入页（已刷新仍未恢复）"

    status, detail = detect_status(page)
    if status != AccountStatus.UNKNOWN:
        return status, detail

    err = detect_login_errors(page)
    if err:
        return err

    if any(x in url for x in ["login.microsoftonline.com", "login.live.com", "account.live.com"]):
        if is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
            return AccountStatus.TIMEOUT, "Microsoft 页面加载超时（半加载壳）"
        if is_stuck_on_login_email_page(page):
            return AccountStatus.UNKNOWN, "网卡卡在邮箱输入页（已刷新仍未恢复）"
        if is_password_entry_page(page):
            text = page_text(page)
            pwd_reject_markers = [
                "パスワードでのサインインは使用できません",
                "正しくないアカウントまたはパスワード",
                "Sign-in with password is not available",
                "incorrect password",
                "password is incorrect",
                "Your account or password is incorrect",
                "密码不正确",
                "密码错误",
            ]
            if any(m in text for m in pwd_reject_markers):
                return AccountStatus.BAD_PASSWORD, "密码错误"
        # 超时前：若已是绑定页，按需绑定而非笼统超时
        if is_definitely_recovery_flow_page(page) or is_recovery_bind_page(page):
            return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱（超时前已出现绑定页）"
        if is_code_verify_page(page):
            return AccountStatus.WAIT_CODE, "等待验证码（超时前已到验证码页）"
        return AccountStatus.TIMEOUT, "Microsoft 验证流程等待超时"

    return AccountStatus.UNKNOWN, "登录结果未确认"


def detect_status(page: Page) -> tuple[AccountStatus, str]:
    url = page.url.lower()

    err = detect_login_errors(page)
    if err:
        return err

    bound = detect_already_bound_email(page)
    if bound:
        return bound

    if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
        return AccountStatus.OK, "已进入邮箱页面"

    if any(
        x in page_text_lower(page) or x in page_text(page)
        for x in [
            "help us protect your account",
            "let's protect your account",
            "アカウントを保護しましょう",
            "アカウントを保護",
            "アカウントの保護",
            "verify your identity",
            "approve sign in request",
            "enter the code",
            "authenticator",
            "验证你的身份",
            "帮助我们保护你的帐户",
            "发送验证码",
            "メールの追加",
        ]
    ):
        if is_phone_verify_page(page):
            return AccountStatus.NEED_PHONE, "需要电话认证（确认手机号末位，非辅助邮箱流程）"
        if is_identity_verification_page(page):
            return AccountStatus.NEED_IDENTITY, "无法绑定辅助邮箱（Microsoft 检测到异常活动，要求本人确认）"
        if is_recovery_bind_page(page) or _page_looks_like_recovery_bind(page):
            if is_ms_auth_page_shell(page) or is_microsoft_login_page_loading(page):
                return AccountStatus.UNKNOWN, "登录页仍在加载中"
            return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"
        return AccountStatus.NEED_VERIFY, "需要 MFA/安全验证"

    text = page_text_lower(page)
    if any(x in text for x in ["captcha", "human", "robot", "人机验证", "我不是机器人"]):
        return AccountStatus.CAPTCHA, "需要人机验证"

    if is_stay_signed_in_page(page) or any(
        x in text for x in ["stay signed in", "保持登录", "keep me signed in"]
    ):
        return AccountStatus.STAY_SIGNED_IN, "等待确认是否保持登录"

    if "login.microsoftonline.com" in url or "login.live.com" in url:
        if is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
            return AccountStatus.UNKNOWN, "登录页仍在加载中"
        return AccountStatus.UNKNOWN, "仍在登录流程中"

    return AccountStatus.UNKNOWN, f"当前 URL: {page.url}"


def click_stay_signed_in(page: Page) -> None:
    wait_and_click(
        page,
        [
            'input[type="submit"][value="Yes"]',
            'input[type="submit"][value="是"]',
            'input[type="submit"][value="はい"]',
            "#idSIButton9",
            'button:has-text("Yes")',
            'button:has-text("是")',
            'button:has-text("はい")',
            'input[value="はい"]',
        ],
        timeout_ms=5000,
    )


def is_passkey_setup_page(page: Page) -> bool:
    text = page_text_lower(page)
    return any(
        x in text
        for x in [
            "パスキーを設定",
            "setting up a passkey",
            "create a passkey",
            "パスキーの設定",
            "セキュリティ ウィンドウが開いています",
            "security window is open",
        ]
    )


def dismiss_passkey_setup(page: Page) -> bool:
    """通行密钥设置页（パスキーを設定しています）点「キャンセル」。"""
    if not is_passkey_setup_page(page):
        return False
    clicked = wait_and_click(
        page,
        [
            'button:has-text("キャンセル")',
            'input[type="button"][value="キャンセル"]',
            'input[value="キャンセル"]',
            'button:has-text("Cancel")',
            'input[type="button"][value="Cancel"]',
            'input[value="Cancel"]',
            'button:has-text("取消")',
            "#idBtn_Back",
            'button[data-testid*="cancel" i]',
        ],
        timeout_ms=2500,
    )
    if clicked:
        log_step("通行密钥页已点 キャンセル")
    return clicked


def is_stay_signed_in_page(page: Page) -> bool:
    text = page_text_lower(page)
    return any(
        x in text
        for x in [
            "stay signed in",
            "keep me signed in",
            "保持登录",
            "保持登入",
            "サインインの状態を維持",
            "サインインしたままにする",
        ]
    )


def is_outlook_inbox(page: Page) -> bool:
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    return "outlook.live.com/mail" in url or "outlook.office.com/mail" in url


AMZ_BAN_SUBJECT_MARKERS = (
    "お客様の Amazon アカウントを一時停止",
    "Amazon アカウントを一時停止いたしました",
    "アカウントを一時停止いたしました",
    "ご利用の Amazon アカウントを一時停止",
    # 列表截断常见前缀
    "お客様の Amazon アカ",
    "Amazon アカウントを一時停止",
    "Amazon アカ",
    # 西语封号主题（含截断）
    "Se ha suspendido su cuenta",
    "Se ha suspendido su cu",
)
AMZ_BAN_BODY_MARKERS = (
    "この措置が講じられた理由",
    "利用規約に違反する形で Amazon アカウントを使用",
    "利用規約に違反",
    "不正確な情報を提供",
    "アカウントを一時停止いたしましたのでご連絡",
    # 西语正文
    "¿A qué se debe esto?",
    "infringiendo nuestra política",
)
# 用户确认：出现即视为 AMZ 封号（封号通知信固定话术）
AMZ_BAN_STRONG_PHRASES = (
    # 日文
    "平素は Amazon をご利用いただき",
    "平素はAmazonをご利用いただき",
    "平素は Amazonをご利用いただき",
    "平素はAmazon をご利用いただき",
    # 西语（amazon.es）
    "Hemos suspendido su cuenta de Amazon",
    "Hemos suspendido su cuenta",
    "Se ha suspendido su cuenta",
    "Se ha suspendido su cu",
)
AMZ_BAN_SENDER_MARKERS = (
    "baa-customer-appeal@amazon.co.jp",
    "baa-customer-appeal@amazon.es",
    "baa-customer-appeal@amazon.",  # 各国站点截断/通用
    "baa-customer-appeal@amazon.c",  # 列表截断
    "baa-customer-appeal",
    "baa-customer-",  # 正文发件人进一步截断
)
# 英文订单处理失败类（用户确认亦属封号相关）
AMZ_BAN_EN_MARKERS = (
    "we had a problem processing",
    "had a problem processing your",
    "problem processing your order",
)
AMZ_BAN_EN_SENDERS = (
    "no-reply@amazon.co.jp",
    "no-reply@amazon.",
    "auto-confirm@amazon",
)


def _normalize_amz_scan_text(text: str) -> str:
    """压缩空白，便于匹配列表截断/换行差异。"""
    if not text:
        return ""
    return re.sub(r"[\s\u3000]+", " ", text).strip()


def _page_has_amz_ban_mail(text: str) -> bool:
    if not text:
        return False
    lower = text.lower()
    norm = _normalize_amz_scan_text(text)
    norm_compact = norm.replace(" ", "")

    # 强特征：日文「平素は…」/ 西语「Hemos|Se ha suspendido…」→ 直接判封号
    for phrase in AMZ_BAN_STRONG_PHRASES:
        if phrase in text or phrase in norm:
            return True
        if phrase.replace(" ", "") in norm_compact:
            return True
        # 西语不区分大小写
        if phrase.lower() in lower or phrase.lower().replace(" ", "") in lower.replace(" ", ""):
            return True

    # 日文封号：发件人（含列表截断）— 最强信号
    if any(s in lower for s in AMZ_BAN_SENDER_MARKERS):
        return True
    # 日文主题
    if any(m in text for m in AMZ_BAN_SUBJECT_MARKERS):
        return True
    # 正文特征（封号通知信固定话术）
    if any(m in text for m in AMZ_BAN_BODY_MARKERS):
        if "amazon" in lower or "一時停止" in text or "アカウント" in text:
            return True
    if "お客様の Amazon" in text and (
        "一時停止" in text or "アカ" in text or "baa-customer" in lower
    ):
        return True
    if "一時停止" in text and "amazon" in lower and (
        "アカウント" in text or "account" in lower or "アカ" in text
    ):
        if (
            "amazon.co.jp" in lower
            or "baa-customer" in lower
            or "amazon アカ" in lower
            or "措置" in text
            or "ご連絡" in text
        ):
            return True
    # 英文：no-reply@amazon + problem processing / orde…
    if any(m in lower for m in AMZ_BAN_EN_MARKERS):
        if any(s in lower for s in AMZ_BAN_EN_SENDERS) or "amazon.co.jp" in lower:
            return True
        if "your amazon.co.jp orde" in lower or "your amazon.co.jp order" in lower:
            return True
    if "your amazon.co.jp orde" in lower and "problem" in lower:
        return True
    return False


def _collect_reading_pane_text(page: Page) -> str:
    """收集 Outlook 阅读窗格正文（选中 Older 区邮件时常仅此处有完整主题/正文）。"""
    try:
        text = page.evaluate(
            """() => {
                const selectors = [
                    '[data-app-section="ReadingPane"]',
                    '[role="region"][aria-label*="Reading" i]',
                    '[role="region"][aria-label*="阅读"]',
                    '[role="region"][aria-label*="読み取り"]',
                    '[aria-label*="Message body" i]',
                    '[aria-label*="邮件正文"]',
                    '.ReadingPaneContents',
                ];
                for (const sel of selectors) {
                    const el = document.querySelector(sel);
                    if (el) {
                        const t = (el.innerText || el.textContent || '').trim();
                        if (t.length > 20) return t;
                    }
                }
                const mains = document.querySelectorAll('[role="main"]');
                if (mains.length > 1) {
                    const t = (mains[mains.length - 1].innerText || '').trim();
                    if (t.length > 20) return t;
                }
                return '';
            }"""
        )
        return str(text or "")
    except Exception:
        return ""


def _extract_inbox_message_row_texts(page: Page) -> list[str]:
    """提取当前可见的收件箱邮件行文本（含 aria-label / title）。"""
    try:
        rows = page.evaluate(
            """() => {
                const out = [];
                const seen = new Set();
                const push = (raw) => {
                    const t = (raw || '').replace(/\\s+/g, ' ').trim();
                    if (t.length < 2 || seen.has(t)) return;
                    seen.add(t);
                    out.push(t);
                };
                document.querySelectorAll(
                    '[role="option"], [role="listitem"], [data-convid], ' +
                    '[role="row"][aria-label], [draggable="true"][aria-label]'
                ).forEach(el => {
                    push(el.innerText || el.textContent || '');
                    push(el.getAttribute('aria-label') || '');
                    push(el.getAttribute('title') || '');
                });
                document.querySelectorAll('[aria-label]').forEach(el => {
                    const a = el.getAttribute('aria-label') || '';
                    if (/baa-customer|一時停止|Amazon アカ|amazon\\.co\\.jp|suspendido/i.test(a)) {
                        push(a);
                    }
                });
                return out;
            }"""
        )
        return [str(r) for r in (rows or []) if r]
    except Exception:
        return []


def _scroll_inbox_message_list(page: Page, delta: int) -> bool:
    """滚动收件箱邮件列表（虚拟列表需逐步滚动才能加载 Older 区）。"""
    try:
        return bool(
            page.evaluate(
                """(delta) => {
                    const pickScroller = () => {
                        const sels = [
                            '[role="listbox"]',
                            '[data-app-section="MessageList"]',
                            '[aria-label*="Message list" i]',
                            '[aria-label*="邮件列表"]',
                            '[aria-label*="メッセージ リスト"]',
                            '[aria-label*="メッセージ"]',
                        ];
                        for (const sel of sels) {
                            for (const el of document.querySelectorAll(sel)) {
                                if (el.scrollHeight > el.clientHeight + 30) return el;
                            }
                        }
                        for (const el of document.querySelectorAll('div')) {
                            const st = getComputedStyle(el);
                            if ((st.overflowY === 'auto' || st.overflowY === 'scroll') &&
                                el.scrollHeight > el.clientHeight + 80 &&
                                el.querySelector('[role="option"], [data-convid], [role="listitem"]')) {
                                return el;
                            }
                        }
                        return null;
                    };
                    const scroller = pickScroller();
                    if (!scroller) {
                        window.scrollBy(0, delta);
                        return true;
                    }
                    const before = scroller.scrollTop;
                    scroller.scrollTop = Math.min(
                        scroller.scrollTop + delta,
                        scroller.scrollHeight
                    );
                    return scroller.scrollTop > before ||
                        scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 8;
                }""",
                delta,
            )
        )
    except Exception:
        return False


def _scroll_inbox_to_top(page: Page) -> None:
    try:
        page.evaluate(
            """() => {
                for (const sel of [
                    '[role="listbox"]',
                    '[data-app-section="MessageList"]',
                    '[aria-label*="Message list" i]',
                ]) {
                    for (const el of document.querySelectorAll(sel)) {
                        try { el.scrollTop = 0; } catch (e) {}
                    }
                }
            }"""
        )
    except Exception:
        pass


def _inbox_has_older_section(page: Page) -> bool:
    try:
        return bool(
            page.evaluate(
                """() => {
                    const markers = new Set([
                        'Older', '古い', '古いメール', '古いアイテム', '更早', '较旧', '更早的邮件',
                    ]);
                    for (const el of document.querySelectorAll(
                        'span, div, button, [role="heading"], [role="separator"], [role="group"]'
                    )) {
                        const t = (el.innerText || '').trim();
                        if (markers.has(t)) return true;
                    }
                    return false;
                }"""
            )
        )
    except Exception:
        return False


def _scroll_older_section_into_view(page: Page) -> bool:
    """将 Older 分组标题滚入视口，触发其下虚拟列表邮件加载。"""
    try:
        return bool(
            page.evaluate(
                """() => {
                    const markers = new Set([
                        'Older', '古い', '古いメール', '古いアイテム', '更早', '较旧', '更早的邮件',
                    ]);
                    for (const el of document.querySelectorAll(
                        'span, div, button, [role="heading"], [role="separator"], [role="group"], h3, h4'
                    )) {
                        const t = (el.innerText || '').trim();
                        if (!markers.has(t)) continue;
                        try {
                            el.scrollIntoView({ block: 'center', behavior: 'instant' });
                        } catch (e) {
                            el.scrollIntoView();
                        }
                        return true;
                    }
                    return false;
                }"""
            )
        )
    except Exception:
        return False


def _try_open_suspicious_amz_row(page: Page) -> bool:
    """点击疑似封号邮件行，加载阅读窗格后再判。"""
    try:
        opened = page.evaluate(
            """() => {
                const isSuspicious = (t) => {
                    const s = (t || '').replace(/\\s+/g, ' ').trim();
                    if (!s) return false;
                    const lower = s.toLowerCase();
                    if (/baa-customer/i.test(lower)) return true;
                    if (/一時停止/.test(s) && /amazon|アカ/i.test(s)) return true;
                    if (/お客様の\\s*Amazon/.test(s)) return true;
                    if (/suspendido\\s+su\\s+cuenta/i.test(lower)) return true;
                    return false;
                };
                const rows = document.querySelectorAll(
                    '[role="option"], [role="listitem"], [data-convid], [draggable="true"][aria-label]'
                );
                for (const row of rows) {
                    const blob = [
                        row.innerText || '',
                        row.getAttribute('aria-label') || '',
                        row.getAttribute('title') || '',
                    ].join(' ');
                    if (!isSuspicious(blob)) continue;
                    try {
                        row.scrollIntoView({ block: 'center', behavior: 'instant' });
                    } catch (e) {
                        row.scrollIntoView();
                    }
                    row.click();
                    return true;
                }
                return false;
            }"""
        )
        if opened:
            safe_wait(page, 700)
            pane = _collect_reading_pane_text(page)
            if pane and _page_has_amz_ban_mail(pane):
                log_step("发现 AMZ 封号邮件（点击疑似行后阅读窗格）")
                return True
        return False
    except Exception:
        return False


def _scan_inbox_rows_for_amz_ban(page: Page) -> bool:
    """
    逐步滚动 Focused/Other 收件箱列表，遍历 Older 分组下的邮件行。
    封号信常在 Older 区且需虚拟列表滚入视口才出现在 DOM。
    """
    log_step("遍历收件箱邮件列表（含 Older 分组）")
    _scroll_inbox_to_top(page)
    safe_wait(page, 280)

    step_px = 320
    max_rounds = 16
    seen_rows: set[str] = set()
    stagnant = 0
    found_older = False

    for round_i in range(max_rounds):
        pane = _collect_reading_pane_text(page)
        if pane and _page_has_amz_ban_mail(pane):
            log_step("发现 AMZ 封号邮件（阅读窗格）")
            return True

        new_in_round = 0
        for row_text in _extract_inbox_message_row_texts(page):
            if row_text in seen_rows:
                continue
            seen_rows.add(row_text)
            new_in_round += 1
            if _page_has_amz_ban_mail(row_text):
                log_step(f"发现 AMZ 封号邮件（列表行）: {row_text[:96]}")
                return True

        if _amz_ban_visible_on_page(page):
            log_step("发现 AMZ 封号邮件（可见节点）")
            return True

        found_older = found_older or _inbox_has_older_section(page)
        if found_older and round_i in {1, 3, 6}:
            _scroll_older_section_into_view(page)
            safe_wait(page, 320)
            if _try_open_suspicious_amz_row(page):
                return True
            pane = _collect_reading_pane_text(page)
            if pane and _page_has_amz_ban_mail(pane):
                log_step("发现 AMZ 封号邮件（Older 区阅读窗格）")
                return True

        if new_in_round == 0:
            stagnant += 1
        else:
            stagnant = 0

        scrolled = _scroll_inbox_message_list(page, step_px)
        at_bottom = bool(
            page.evaluate(
                """() => {
                    const sels = ['[role="listbox"]', '[data-app-section="MessageList"]'];
                    for (const sel of sels) {
                        for (const el of document.querySelectorAll(sel)) {
                            if (el.scrollHeight > el.clientHeight + 30) {
                                return el.scrollTop + el.clientHeight >= el.scrollHeight - 10;
                            }
                        }
                    }
                    return false;
                }"""
            )
        )
        safe_wait(page, 220 if round_i < 8 else 300)

        if not scrolled and stagnant >= 2:
            break
        if at_bottom and stagnant >= 1:
            break
        if at_bottom and found_older and round_i >= 3:
            break

    if found_older:
        log_step("已扫过 Older 分组区域")
    if _try_open_suspicious_amz_row(page):
        return True
    return False


def _collect_outlook_mail_scan_text(page: Page) -> str:
    """尽量收集收件箱列表+正文文本（含虚拟列表节点）。"""
    chunks: list[str] = []
    try:
        chunks.append(page_text(page) or "")
    except Exception:
        pass
    try:
        pane = _collect_reading_pane_text(page)
        if pane:
            chunks.append(pane)
    except Exception:
        pass
    try:
        extra = page.evaluate(
            """() => {
                const parts = [];
                const push = (el) => {
                    if (!el) return;
                    const t = (el.innerText || el.textContent || '').trim();
                    if (t && t.length > 2) parts.push(t);
                };
                document.querySelectorAll(
                    '[role="option"], [role="listitem"], [role="row"], [data-convid], ' +
                    '[aria-label*="amazon" i], [aria-label*="Amazon"], [aria-label*="baa-customer"], ' +
                    '[aria-label*="一時停止"], [title*="一時停止"], [title*="Amazon"]'
                ).forEach(push);
                document.querySelectorAll('[aria-label]').forEach(el => {
                    const a = el.getAttribute('aria-label') || '';
                    if (/baa-customer|一時停止|Amazon アカ|amazon\\.co\\.jp/i.test(a)) {
                        parts.push(a);
                    }
                });
                return parts.join('\\n');
            }"""
        )
        if extra:
            chunks.append(str(extra))
    except Exception:
        pass
    return "\n".join(chunks)


def _amz_ban_visible_on_page(page: Page) -> bool:
    """列表/正文可见节点快速探测（适配截断发件人/主题）。"""
    probes = (
        "baa-customer-appeal",
        "baa-customer-",
        "一時停止いたしました",
        "お客様の Amazon アカ",
        "お客様の Amazon",
        "ご利用の Amazon アカウントを一時停止",
        "この措置が講じられた理由",
        "平素は Amazon をご利用いただき",
        "Hemos suspendido su cuenta de Amazon",
        "Hemos suspendido su cuenta",
        "Se ha suspendido su cuenta",
        "Se ha suspendido su cu",
        "We had a problem processing",
        "no-reply@amazon.co.jp",
    )
    for label in probes:
        try:
            locs = page.get_by_text(label, exact=False)
            count = min(locs.count(), 8)
            for i in range(count):
                loc = locs.nth(i)
                try:
                    if not loc.is_visible(timeout=120):
                        continue
                except Exception:
                    continue
                weak = label in {
                    "お客様の Amazon",
                    "We had a problem processing",
                }
                # 「平素は…」/「Hemos|Se ha suspendido…」用户确认：出现即封号
                if label.startswith("平素は") or "suspendido su cuenta" in label.lower() or label.startswith("Se ha suspendido"):
                    return True
                if weak:
                    try:
                        nearby = (
                            loc.evaluate(
                                "el => (el.closest('[role=option], [role=listitem], [role=row], [data-convid], li, div') || el).innerText"
                            )
                            or ""
                        )
                    except Exception:
                        nearby = ""
                    if _page_has_amz_ban_mail(nearby):
                        return True
                    if "一時停止" in nearby or "baa-customer" in nearby.lower():
                        return True
                    continue
                return True
        except Exception:
            continue
    try:
        hit = page.evaluate(
            """() => {
                const nodes = document.querySelectorAll('[aria-label], [title]');
                for (const el of nodes) {
                    const a = (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '');
                    if (/baa-customer/i.test(a)) return true;
                    if (/平素は\\s*Amazon\\s*をご利用いただき/.test(a)) return true;
                    if (/Hemos\\s+suspendido\\s+su\\s+cuenta/i.test(a)) return true;
                    if (/Se\\s+ha\\s+suspendido\\s+su\\s+cu/i.test(a)) return true;
                    if (/一時停止/.test(a) && /Amazon|amazon|アカ/.test(a)) return true;
                    if (/お客様の\\s*Amazon/.test(a)) return true;
                }
                return false;
            }"""
        )
        if hit:
            return True
    except Exception:
        pass
    return False


def _try_outlook_search_amz_ban(page: Page, query: str, wait_ms: int = 1600) -> bool:
    """在 Outlook 搜索框执行查询，并等待结果。"""
    search = page.locator(
        '#topSearchInput, input[aria-label*="Search" i], '
        'input[placeholder*="Search" i], input[aria-label*="検索"], '
        'input[placeholder*="検索"], input[aria-label*="搜索"], '
        'input[type="search"], [role="searchbox"]'
    ).first
    if not search.is_visible(timeout=1200):
        return False
    search.click(timeout=800)
    search.fill("", timeout=800)
    search.fill(query, timeout=1500)
    search.press("Enter")
    safe_wait(page, wait_ms)
    return True


def _try_click_outlook_label(page: Page, labels: tuple[str, ...], *, wait_ms: int = 900) -> str | None:
    """点击 Outlook 侧栏/列表上的文件夹或筛选标签，成功返回点中的文案。"""
    for label in labels:
        try:
            for exact in (True, False):
                loc = page.get_by_text(label, exact=exact).first
                try:
                    if not loc.is_visible(timeout=350):
                        continue
                except Exception:
                    continue
                try:
                    clicked = loc.evaluate(
                        """el => {
                            const nav = el.closest(
                                'nav, [role="tree"], [role="treeitem"], [role="tab"], ' +
                                '[role="button"], button, a, [data-folder-name], [aria-label]'
                            );
                            (nav || el).click();
                            return true;
                        }"""
                    )
                    if not clicked:
                        loc.click(timeout=800)
                except Exception:
                    try:
                        loc.click(timeout=800)
                    except Exception:
                        continue
                safe_wait(page, wait_ms)
                return label
        except Exception:
            continue
    for label in labels:
        try:
            loc = page.locator(
                f'[aria-label="{label}"], [aria-label*="{label}"], '
                f'[title="{label}"], [title*="{label}"], '
                f'[data-folder-name="{label}"]'
            ).first
            if loc.is_visible(timeout=300):
                loc.click(timeout=800)
                safe_wait(page, wait_ms)
                return label
        except Exception:
            continue
    return None


def _scan_outlook_older_and_other(page: Page, quick_hit) -> bool:
    """
    封号信常落在 Other / Older（及日文「その他」「古い」）。
    逐个切入并扫描，命中即返回 True。
    """
    # Focused 收件箱内 Older 分组：先滚动遍历当前列表
    if _scan_inbox_rows_for_amz_ban(page):
        return True
    if quick_hit():
        log_step("发现 AMZ 封号邮件（列表遍历后当前页）")
        return True

    other_labels = (
        "Other",
        "その他",
        "その他の受信トレイ",
        "其他",
    )
    hit = _try_click_outlook_label(page, other_labels, wait_ms=1000)
    if hit:
        if _scan_inbox_rows_for_amz_ban(page):
            log_step(f"发现 AMZ 封号邮件（{hit} 列表遍历）")
            return True
        if quick_hit():
            log_step(f"发现 AMZ 封号邮件（文件夹: {hit}）")
            return True
        older_in = _try_click_outlook_label(
            page,
            ("Older", "古い", "古いメール", "较旧", "更早"),
            wait_ms=1100,
        )
        if older_in:
            if _scan_inbox_rows_for_amz_ban(page) or quick_hit():
                log_step(f"发现 AMZ 封号邮件（{hit} → {older_in}）")
                return True

    older_labels = (
        "Older",
        "古い",
        "古いメール",
        "古いアイテム",
        "较旧",
        "更早的邮件",
        "更早",
    )
    hit = _try_click_outlook_label(page, older_labels, wait_ms=1100)
    if hit:
        if _scan_inbox_rows_for_amz_ban(page) or quick_hit():
            log_step(f"发现 AMZ 封号邮件（文件夹: {hit}）")
            return True

    try:
        page.evaluate(
            """() => {
                const lists = document.querySelectorAll(
                    '[role="listbox"], [role="list"], [data-app-section="MessageList"], ' +
                    '.customScrollBar, [class*="scroll"]'
                );
                for (const el of lists) {
                    try { el.scrollTop = el.scrollHeight; } catch (e) {}
                }
                window.scrollTo(0, document.body.scrollHeight);
            }"""
        )
        safe_wait(page, 600)
        if _scan_inbox_rows_for_amz_ban(page):
            log_step("发现 AMZ 封号邮件（滚动后列表遍历）")
            return True
        hit = _try_click_outlook_label(page, older_labels, wait_ms=1100)
        if hit and ( _scan_inbox_rows_for_amz_ban(page) or quick_hit()):
            log_step(f"发现 AMZ 封号邮件（滚动后 {hit}）")
            return True
    except Exception:
        pass

    try:
        cur = (page.url or "").split("?")[0]
        if "outlook.live.com/mail" in cur or "outlook.office.com/mail" in cur:
            for dest in (
                "https://outlook.live.com/mail/0/inbox?view=other",
                "https://outlook.live.com/mail/0/junkemail",
            ):
                page.goto(dest, wait_until="domcontentloaded", timeout=12000)
                safe_wait(page, 1200)
                if _scan_inbox_rows_for_amz_ban(page) or quick_hit():
                    log_step(f"发现 AMZ 封号邮件（URL: {dest.split('/')[-1]}）")
                    return True
                older_in = _try_click_outlook_label(
                    page, ("Older", "古い", "古いメール"), wait_ms=1000
                )
                if older_in and (_scan_inbox_rows_for_amz_ban(page) or quick_hit()):
                    log_step(f"发现 AMZ 封号邮件（URL + {older_in}）")
                    return True
    except Exception:
        pass

    return False


def check_amz_account_banned_mail(page: Page, timeout_sec: float = 10.0) -> bool:
    """
    登录进 Outlook 后检查是否有亚马逊账号暂停/封号相关邮件。
    优先搜索（Older 区最快），再短扫列表；整体控制在约 10 秒内。
    """
    log_step("检查收件箱是否有 AMZ 封号邮件")
    timeout_sec = max(8.0, float(timeout_sec))
    deadline = time.time() + timeout_sec

    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    if "outlook.live.com/mail" not in url and "outlook.office.com/mail" not in url:
        try:
            page.goto(
                "https://outlook.live.com/mail/0/inbox",
                wait_until="domcontentloaded",
                timeout=10000,
            )
        except Exception:
            pass

    try:
        page.wait_for_selector(
            '[role="option"], [role="listitem"], [data-convid], [aria-label*="邮件"], [aria-label*="Message"]',
            timeout=3000,
        )
    except Exception:
        pass
    safe_wait(page, 350)

    def _quick_hit() -> bool:
        text = _collect_outlook_mail_scan_text(page)
        if _page_has_amz_ban_mail(text):
            return True
        return _amz_ban_visible_on_page(page)

    if _quick_hit():
        log_step("发现 AMZ 账号暂停/封号邮件（当前页）")
        return True

    # 优先搜索：Older/Other 里的封号信用搜索最快
    from urllib.parse import quote

    if time.time() < deadline:
        try:
            dest = "https://outlook.live.com/mail/0/search?q=" + quote("baa-customer-appeal")
            page.goto(dest, wait_until="domcontentloaded", timeout=10000)
            safe_wait(page, 900)
            if _quick_hit():
                log_step("发现 AMZ 封号邮件（搜索URL: baa-customer-appeal）")
                return True
        except Exception as exc:
            log_step(f"封号搜索URL失败: {exc}")

    if time.time() < deadline:
        try:
            if _try_outlook_search_amz_ban(page, "baa-customer-appeal", wait_ms=1000):
                if _quick_hit():
                    log_step("发现 AMZ 封号邮件（搜索框）")
                    return True
        except Exception:
            pass

    # 搜索未中再短扫列表（含 Older），避免拖太久
    if time.time() < deadline:
        try:
            # 回到收件箱再扫
            page.goto(
                "https://outlook.live.com/mail/0/inbox",
                wait_until="domcontentloaded",
                timeout=10000,
            )
            safe_wait(page, 400)
            if _scan_outlook_older_and_other(page, _quick_hit):
                return True
        except Exception as exc:
            log_step(f"Other/Older 扫描异常: {exc}")

    found = _quick_hit()
    log_step("AMZ 封号邮件: " + ("有" if found else "无"))
    if not found:
        try:
            ensure_outlook_inbox_for_screenshot(page)
        except Exception:
            pass
    return found


def evaluate_inbox_account_status(page: Page) -> tuple[AccountStatus, str]:
    """进入邮箱后：有封号信 → AMZ账号被封；否则 → 登入。"""
    banned = check_amz_account_banned_mail(page)
    if banned:
        try:
            for label in (
                "お客様の Amazon アカウントを一時停止",
                "お客様の Amazon アカ",
                "一時停止いたしました",
                "baa-customer-appeal",
                "baa-customer-",
                "ご利用の Amazon アカウントを一時停止",
                "この措置が講じられた理由",
                "Hemos suspendido su cuenta de Amazon",
                "Hemos suspendido su cuenta",
                "Se ha suspendido su cuenta",
                "Se ha suspendido su cu",
                "We had a problem processing",
                "Your Amazon.co.jp orde",
                "no-reply@amazon.co.jp",
                "baa-customer-appeal@amazon.es",
            ):
                loc = page.get_by_text(label, exact=False).first
                if loc.is_visible(timeout=500):
                    loc.click(timeout=1000)
                    safe_wait(page, 500)
                    break
        except Exception:
            pass
        # 封号结果也尽量停在可读邮件页；若已在搜索/垃圾箱则保留当前视图
        return (
            AccountStatus.AMZ_BANNED,
            "收件箱/Other/Older 发现 Amazon 账号暂停/封号邮件（baa-customer-appeal / 一時停止）",
        )
    # 扫描会跳到搜索/垃圾箱；登入截图前必须回到收件箱
    ensure_outlook_inbox_for_screenshot(page)
    return AccountStatus.LOGIN_OK, "登入成功，未发现 AMZ 封号邮件"


def ensure_outlook_inbox_for_screenshot(page: Page) -> None:
    """把 Outlook 从搜索/垃圾箱/Other 拉回收件箱，避免终态截到 Junk Email。"""
    try:
        url = (page.url or "").lower()
    except Exception:
        return
    if "outlook.live.com/mail" not in url and "outlook.office.com/mail" not in url:
        return
    need_inbox = (
        "junk" in url
        or "/search" in url
        or "view=other" in url
        or "/deleted" in url
        or "/inbox" not in url
    )
    if not need_inbox:
        return
    try:
        log_step("截图前回到 Outlook 收件箱")
        page.goto(
            "https://outlook.live.com/mail/0/inbox",
            wait_until="domcontentloaded",
            timeout=12000,
        )
        safe_wait(page, 900)
    except Exception as exc:
        log_step(f"回收件箱失败: {exc}")


def complete_login_after_recovery_code(
    page: Page,
    *,
    timeout_sec: float = 75.0,
    on_progress: Callable[[str], None] | None = None,
) -> tuple[bool, str, AccountStatus]:
    """
    验证码已填并点「次へ」之后：
    パスキー页 → キャンセル → 保持登录页 → はい → 进入 Outlook 收件箱 → 查 AMZ 封号信。
    """
    def progress(msg: str) -> None:
        log_step(msg)
        if on_progress:
            try:
                on_progress(msg)
            except Exception:
                pass

    deadline = time.time() + max(45.0, float(timeout_sec))
    last_log = 0.0
    last_step = ""
    tried_outlook_goto = False

    def set_step(msg: str) -> None:
        nonlocal last_step, last_log
        if msg != last_step or time.time() - last_log > 2.0:
            progress(msg)
            last_step = msg
            last_log = time.time()

    set_step("登录收尾中：等待页面跳转…")

    while time.time() < deadline:
        # 多标签：优先切到 Outlook / 登录流页面
        try:
            context = page.context
            for candidate in reversed(context.pages):
                try:
                    u = (candidate.url or "").lower()
                except Exception:
                    continue
                if "outlook.live.com/mail" in u or "outlook.office.com/mail" in u:
                    page = candidate
                    break
        except Exception:
            pass

        if is_outlook_inbox(page):
            set_step("已进入收件箱，正在检查 AMZ 封号邮件…")
            status, detail = evaluate_inbox_account_status(page)
            return True, detail, status

        if is_code_verify_page(page):
            text = page_text(page)
            wrong_markers = [
                "コードが正しくありません",
                "That code didn't work",
                "code didn't work",
                "Enter the code again",
                "もう一度入力",
                "验证码不正确",
                "コードが無効",
            ]
            if any(m.lower() in text.lower() if m.isascii() else m in text for m in wrong_markers):
                return False, "验证码错误，仍在验证码页", AccountStatus.WAIT_CODE
            set_step("仍在验证码页，等待跳转…")
            safe_wait(page, 400)
            if is_code_verify_page(page) and time.time() > deadline - 8:
                return False, "验证码已提交但仍在验证码页（可能码错误）", AccountStatus.WAIT_CODE
            continue

        if is_passkey_setup_page(page):
            set_step("检测到通行密钥页，正在点取消…")
            dismiss_passkey_setup(page)
            safe_wait(page, 500)
            continue

        if is_stay_signed_in_page(page):
            set_step("保持登录页：点击 はい…")
            click_stay_signed_in(page)
            safe_wait(page, 800)
            continue

        text = page_text(page)
        text_lower = text.lower()

        # 偶发：点了次へ后出现通行密钥相关页，优先取消
        if "パスキーを設定" in text or "setting up a passkey" in text_lower or "passkey" in text_lower:
            set_step("通行密钥相关页：尝试取消…")
            dismiss_passkey_setup(page)
            safe_wait(page, 450)
            continue

        # 账户选择页
        if any(
            x in text_lower
            for x in (
                "pick an account",
                "select an account",
                "アカウントの選択",
                "选择帐户",
                "选择账户",
            )
        ):
            set_step("账户选择页：尝试继续…")
            click_account_picker_other(page)
            # 若列表里有账号，点第一个
            try:
                tiles = page.locator('#tilesHolder [role="listitem"], #tilesHolder [data-test-id], .tile')
                if tiles.count() > 0:
                    tiles.first.click(timeout=1500)
                    safe_wait(page, 800)
            except Exception:
                pass
            safe_wait(page, 500)
            continue

        # 同意 / 权限页
        if any(x in text_lower for x in ("accept", "同意", "許可", "permissions requested")):
            if click_login_submit(
                page,
                [
                    '#idSIButton9',
                    'input[type="submit"]',
                    'button[type="submit"]',
                    'button:has-text("Accept")',
                    'button:has-text("はい")',
                    'button:has-text("同意")',
                ],
            ):
                set_step("权限/同意页：已点击继续…")
                safe_wait(page, 700)
                continue

        # 半加载 / 空白登录壳：再等一会儿
        if is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
            set_step("Microsoft 页面加载中…")
            safe_wait(page, 700)
            continue

        # 验证码后卡在 oauth / login 域名：尝试直达 Outlook
        try:
            url = (page.url or "").lower()
        except Exception:
            url = ""
        if (
            not tried_outlook_goto
            and time.time() > deadline - (timeout_sec * 0.35)
            and any(
                x in url
                for x in (
                    "login.microsoftonline.com",
                    "login.live.com",
                    "account.live.com",
                    "oauth2",
                )
            )
        ):
            tried_outlook_goto = True
            set_step("登录页停滞，尝试直接打开 Outlook…")
            try:
                page.goto(
                    "https://outlook.live.com/mail/0/",
                    wait_until="domcontentloaded",
                    timeout=20000,
                )
                safe_wait(page, 1500)
            except Exception as exc:
                set_step(f"打开 Outlook 失败，继续等待… ({str(exc)[:60]})")
            continue

        if "outlook.live.com" in url and "/mail" not in url:
            set_step("已到 Outlook，正在进入收件箱…")
            try:
                page.goto(
                    "https://outlook.live.com/mail/0/",
                    wait_until="domcontentloaded",
                    timeout=15000,
                )
                safe_wait(page, 1200)
            except Exception:
                pass
            continue

        now = time.time()
        if now - last_log > 2.2:
            try:
                short = (page.url or "").split("?")[0][-56:]
            except Exception:
                short = ""
            set_step(f"登录收尾中… {short}")
        safe_wait(page, 280)

    if is_outlook_inbox(page):
        set_step("已进入收件箱，正在检查 AMZ 封号邮件…")
        status, detail = evaluate_inbox_account_status(page)
        return True, detail, status
    if is_code_verify_page(page):
        return False, "验证码可能错误，仍停在验证码页", AccountStatus.WAIT_CODE
    if is_stay_signed_in_page(page):
        set_step("超时前最后尝试：保持登录 はい…")
        click_stay_signed_in(page)
        safe_wait(page, 1500)
        if is_outlook_inbox(page):
            status, detail = evaluate_inbox_account_status(page)
            return True, detail, status
    try:
        stuck_url = (page.url or "").split("?")[0][-80:]
    except Exception:
        stuck_url = ""
    return (
        False,
        f"验证码已提交，但未进入邮箱（超时）{(' | ' + stuck_url) if stuck_url else ''}。"
        f"可点打开，粘贴右侧 Outlook 登录链接重试",
        AccountStatus.TIMEOUT,
    )


def safe_wait(page: Page, ms: int) -> None:
    try:
        page.wait_for_timeout(ms)
    except Exception:
        pass


def get_login_page(context, login_url: str, timeout_sec: int = 12) -> Page:
    """快速定位登录页，邮箱框出现即返回。"""
    deadline = time.time() + timeout_sec

    while time.time() < deadline:
        for candidate in reversed(context.pages):
            try:
                url = candidate.url.lower()
            except Exception:
                continue
            if not any(
                x in url
                for x in ["login.microsoftonline.com", "login.live.com", "account.live.com"]
            ):
                continue
            if find_email_input(candidate) or find_password_input(candidate):
                return candidate
            click_account_picker_other(candidate)
        time.sleep(0.12)

    page = context.pages[-1] if context.pages else context.new_page()
    try:
        cur = page.url.lower()
    except Exception:
        cur = ""
    if not any(x in cur for x in ["login.microsoftonline.com", "login.live.com"]):
        try:
            safe_page_goto(page, login_url, timeout_ms=25000, retries=3)
        except Exception as exc:
            if not _is_transient_nav_error(exc):
                raise
            log_step(f"get_login_page 打开失败（将继续等表单）: {str(exc)[:120]}")
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if find_email_input(page) or find_password_input(page):
            return page
        click_account_picker_other(page)
        time.sleep(0.12)
    return page


def get_active_flow_page(context, login_url: str, timeout_sec: int = 12) -> Page:
    """定位当前登录/验证/绑定流程页面，避免误跳回登录表单。"""
    deadline = time.time() + timeout_sec
    flow_hosts = [
        "login.microsoftonline.com",
        "login.live.com",
        "account.live.com",
        "outlook.live.com",
        "outlook.office.com",
    ]

    while time.time() < deadline:
        for candidate in reversed(context.pages):
            try:
                url = candidate.url.lower()
            except Exception:
                continue
            if not any(h in url for h in flow_hosts):
                continue
            if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
                return candidate
            if is_recovery_bind_page(candidate):
                return candidate
            try:
                if find_recovery_email_input(candidate) is not None:
                    return candidate
            except Exception:
                pass
            text = ""
            try:
                text = page_text(candidate)
            except Exception:
                try:
                    text = candidate.inner_text("body")
                except Exception:
                    pass
            if any(
                p in text
                for p in ["アカウントを保護", "protect your account", "someone@example.com"]
            ):
                return candidate
            if is_password_entry_page(candidate):
                return candidate
            if find_email_input(candidate) or find_password_input(candidate):
                return candidate
            click_account_picker_other(candidate)
        time.sleep(0.12)

    for candidate in reversed(context.pages):
        try:
            if is_password_entry_page(candidate):
                return candidate
        except Exception:
            continue

    return get_login_page(context, login_url, timeout_sec=8)


def resolve_flow_page_for_password(
    page: Page, context, login_url: str, fast: bool
) -> Page:
    """已在密码页时保持当前标签，避免 get_active_flow_page 误跳回邮箱页。"""
    if is_password_entry_page(page) or find_password_input(page) is not None:
        return page
    try:
        if page.evaluate(f"() => !!({_JS_PICK_PASSWORD_INPUT})()"):
            return page
    except Exception:
        pass
    for candidate in reversed(context.pages):
        try:
            if is_password_entry_page(candidate) or find_password_input(candidate):
                return candidate
        except Exception:
            continue
    return get_active_flow_page(context, login_url, timeout_sec=6 if fast else 10)


def _start_sync_playwright():
    """启动 Playwright Sync；异常时保证可清理，避免线程池复用时残留 asyncio loop。"""
    # 若本线程误留了 running loop（上次异常未 stop），先清掉再启
    try:
        loop = asyncio.get_event_loop_policy().get_event_loop()
        if loop is not None and loop.is_running():
            # 不能在 running loop 里直接 sync_playwright；尝试摘掉错误标记
            try:
                asyncio._set_running_loop(None)  # type: ignore[attr-defined]
            except Exception:
                pass
    except Exception:
        pass
    return sync_playwright().start()


def _cdp_ignore_certificate_errors(browser) -> None:
    """HubStudio 环境常有 HTTPS 中间人证书，忽略以免 page.goto 报 ERR_CERT_AUTHORITY_INVALID。"""
    try:
        session = browser.new_browser_cdp_session()
        session.send("Security.setIgnoreCertificateErrors", {"ignore": True})
    except Exception:
        pass


_TRANSIENT_NAV_MARKERS = (
    "ERR_CONNECTION_CLOSED",
    "ERR_CONNECTION_RESET",
    "ERR_CONNECTION_TIMED_OUT",
    "ERR_CONNECTION_REFUSED",
    "ERR_NAME_NOT_RESOLVED",
    "ERR_TUNNEL_CONNECTION_FAILED",
    "ERR_PROXY_CONNECTION_FAILED",
    "ERR_NETWORK_CHANGED",
    "ERR_INTERNET_DISCONNECTED",
    "ERR_SSL_PROTOCOL_ERROR",
    "ERR_TIMED_OUT",
    "ERR_EMPTY_RESPONSE",
    "ERR_HTTP2",
    "net::ERR_",
    "NS_ERROR_NET",
    "Timeout",
    "timeout",
    "Navigation failed",
)


def _is_transient_nav_error(exc: BaseException) -> bool:
    msg = str(exc) or ""
    return any(m in msg for m in _TRANSIENT_NAV_MARKERS)


def _login_goto_fallbacks(login_url: str) -> list[str]:
    """OAuth 长链经常 ERR_CONNECTION_CLOSED，准备更短的登录入口作重试。"""
    urls: list[str] = []
    primary = (login_url or "").strip()
    if primary:
        urls.append(primary)
    for alt in (
        "https://login.live.com/",
        "https://login.microsoftonline.com/",
        "https://outlook.live.com/mail/0/",
    ):
        if alt not in urls:
            urls.append(alt)
    return urls


def safe_page_goto(
    page: Page,
    url: str,
    *,
    timeout_ms: int = 25000,
    retries: int = 3,
    wait_until: str = "domcontentloaded",
) -> bool:
    """
    带重试的 page.goto。遇到 ERR_CONNECTION_CLOSED 等网卡/代理抖动时：
    等待 → 换 wait_until → 换备用登录 URL，避免直接「脚本异常」。
    """
    targets = _login_goto_fallbacks(url)
    last_exc: Exception | None = None
    wait_modes = (wait_until, "commit", "load")

    for attempt in range(max(1, retries)):
        target = targets[min(attempt, len(targets) - 1)]
        mode = wait_modes[min(attempt, len(wait_modes) - 1)]
        try:
            log_step(f"打开登录页 ({attempt + 1}/{retries}): {target[:72]}… [{mode}]")
            page.goto(target, wait_until=mode, timeout=timeout_ms)
            safe_wait(page, 400)
            # 落在 Chrome 网卡错误页则继续重试
            if detect_network_card_error(page):
                log_step("页面显示连接中断，准备重试…")
                time.sleep(0.8 + 0.4 * attempt)
                continue
            return True
        except Exception as exc:
            last_exc = exc
            msg = str(exc)
            log_step(f"打开登录页失败: {msg[:160]}")
            if "ERR_CERT" in msg or "CERTIFICATE" in msg.upper():
                time.sleep(0.5)
                try:
                    page.goto(target, wait_until="domcontentloaded", timeout=timeout_ms)
                    return True
                except Exception as exc2:
                    last_exc = exc2
            if not _is_transient_nav_error(exc):
                break
            # 先 reload 当前错误页，再换 URL
            try:
                page.reload(wait_until="domcontentloaded", timeout=min(15000, timeout_ms))
                if not detect_network_card_error(page) and (
                    find_email_input(page) or find_password_input(page)
                ):
                    return True
            except Exception:
                pass
            time.sleep(0.9 + 0.5 * attempt)

    if last_exc is not None:
        raise last_exc
    return False


def acquire_login_page(port: int, login_url: str, max_wait_sec: float | None = None):
    """连接 CDP 并拿到可填表的登录页（必须检测到邮箱/密码框才返回）。"""
    register_automation_browser(port)
    log_step("正在连接浏览器 CDP")
    if max_wait_sec is None:
        max_wait_sec = float(_RUNTIME.get("acquire_page_max_sec", 12))
    poll = 0.08 if _fast_fill_enabled() else 0.15
    form_wait = int(_RUNTIME.get("login_form_wait_sec", 6))
    playwright = _start_sync_playwright()
    cdp_url = f"http://127.0.0.1:{port}"
    cdp_timeout = int(_RUNTIME["cdp_connect_timeout_ms"])
    browser = None
    last_exc: Exception | None = None
    try:
        for attempt in range(int(_RUNTIME["hubstudio_retries"])):
            try:
                browser = playwright.chromium.connect_over_cdp(cdp_url, timeout=cdp_timeout)
                break
            except Exception as exc:
                last_exc = exc
                if attempt + 1 < int(_RUNTIME["hubstudio_retries"]):
                    time.sleep(2 * (attempt + 1))
        if browser is None:
            raise RuntimeError(f"CDP 连接失败 (port={port}): {last_exc}")

        _cdp_ignore_certificate_errors(browser)

        deadline = time.time() + max_wait_sec
        page: Page | None = None
        context = browser.contexts[0] if browser.contexts else browser.new_context()

        while time.time() < deadline:
            for candidate in reversed(context.pages):
                try:
                    url = candidate.url.lower()
                except Exception:
                    continue
                # 已是网卡错误页：不要当可用登录页
                try:
                    if detect_network_card_error(candidate):
                        continue
                except Exception:
                    pass
                if not any(
                    x in url
                    for x in [
                        "login.microsoftonline.com",
                        "login.live.com",
                        "account.live.com",
                        "account.microsoft.com",
                        "signup.live.com",
                    ]
                ):
                    continue
                # 已在辅助邮箱页也可直接返回，交给 auto_login 识别
                try:
                    if is_recovery_bind_page(candidate) or is_code_verify_page(candidate):
                        page = candidate
                        break
                except Exception:
                    pass
                click_account_picker_other(candidate)
                if find_email_input(candidate) or find_password_input(candidate):
                    page = candidate
                    break
            if page is not None:
                break
            time.sleep(poll)

        if page is None:
            page = context.new_page()
            try:
                safe_page_goto(page, login_url, timeout_ms=25000, retries=3)
            except Exception as goto_exc:
                # 证书问题再忽略一次后重试
                msg = str(goto_exc)
                if "ERR_CERT" in msg or "CERTIFICATE" in msg.upper():
                    _cdp_ignore_certificate_errors(browser)
                    safe_page_goto(page, login_url, timeout_ms=25000, retries=2)
                elif _is_transient_nav_error(goto_exc):
                    # 交给上层标成「网卡」，不要冒成未捕获脚本异常细节过长
                    raise RuntimeError(
                        f"ERR_CONNECTION / 网卡打开登录页失败: {msg[:220]}"
                    ) from goto_exc
                else:
                    raise

        # 打开后若落在错误页，再刷一轮
        try:
            if detect_network_card_error(page):
                log_step("登录页为连接错误页，自动刷新重试")
                safe_page_goto(page, login_url, timeout_ms=25000, retries=3)
        except Exception:
            pass

        click_account_picker_other(page)
        log_step("等待登录页表单加载")
        # 辅助邮箱页不需要等登录表单
        if not (is_recovery_bind_page(page) or is_code_verify_page(page)):
            ensure_login_form_ready(page, login_url=login_url, timeout_sec=form_wait)

        if (
            _RUNTIME.get("browser_minimize_on_start")
            and not _effective_headless()
        ):
            time.sleep(0.08 if _fast_fill_enabled() else 0.2)
            minimize_browser_windows_debounced()

        return playwright, browser, page
    except Exception:
        try:
            playwright.stop()
        except Exception:
            pass
        raise


def ensure_browser_focus_for_input(page: Page) -> None:
    """非后台模式才短暂恢复窗口；后台模式全程不抢焦点。"""
    if _effective_headless() or _keep_background():
        return
    try:
        restore_hubstudio_windows()
        page.bring_to_front()
    except Exception:
        pass
    safe_wait(page, 300)


def is_interactable(loc, fast: bool = False) -> bool:
    probe_ms = 120 if fast else 300
    try:
        if loc.count() == 0:
            return False
        if not loc.is_visible(timeout=probe_ms):
            return False
        if loc.get_attribute("aria-hidden") == "true":
            return False
        cls = loc.get_attribute("class") or ""
        if "moveOffScreen" in cls:
            return False
        return loc.is_enabled(timeout=probe_ms)
    except Exception:
        return False


def find_email_input(page: Page):
    if is_password_entry_page(page) or password_field_ready(page, None):
        return None
    try:
        if is_recovery_bind_page(page) or is_code_verify_page(page):
            return None
    except Exception:
        pass
    # 后台模式：DOM 查找，不依赖 is_visible
    if _keep_background():
        try:
            marked = page.evaluate(
                f"""() => {{
                    document.querySelectorAll('[data-email-target]').forEach(el => {{
                        el.removeAttribute('data-email-target');
                    }});
                    const pick = {_JS_PICK_EMAIL_INPUT};
                    const el = pick();
                    if (el) {{ el.setAttribute('data-email-target', '1'); return true; }}
                    return false;
                }}"""
            )
            if marked:
                loc = page.locator('[data-email-target="1"]').first
                if loc.count():
                    return loc
        except Exception:
            pass
        for sel in ['#i0116', 'input[name="loginfmt"]']:
            loc = page.locator(sel).first
            if not loc.count():
                continue
            cls = loc.get_attribute("class") or ""
            if "moveOffScreen" in cls:
                continue
            try:
                if loc.is_enabled(timeout=400):
                    return loc
            except Exception:
                continue
        return None

    for sel in ['#i0116', 'input[name="loginfmt"]']:
        loc = page.locator(sel).first
        if loc.count() and is_interactable(loc):
            return loc
    for sel in ['#i0116', 'input[name="loginfmt"]']:
        loc = page.locator(sel).first
        if loc.count():
            try:
                if loc.is_enabled(timeout=400):
                    return loc
            except Exception:
                continue
    return None


def find_password_input(page: Page):
    # 后台模式：DOM 查找，不依赖 is_visible（最小化时 visibility 常为 false）
    if _keep_background():
        try:
            marked = page.evaluate(
                f"""() => {{
                    document.querySelectorAll('[data-pw-target]').forEach(el => {{
                        el.removeAttribute('data-pw-target');
                    }});
                    const pick = {_JS_PICK_PASSWORD_INPUT};
                    const el = pick();
                    if (el) {{ el.setAttribute('data-pw-target', '1'); return true; }}
                    return false;
                }}"""
            )
            if marked:
                loc = page.locator('[data-pw-target="1"]').first
                if loc.count():
                    return loc
        except Exception:
            pass
        for sel in [
            'input[name="passwd"]',
            '#i0118',
            '#passwordEntry',
            'input[type="password"]',
            'input[autocomplete="current-password"]',
        ]:
            loc = page.locator(sel).first
            if not loc.count():
                continue
            cls = loc.get_attribute("class") or ""
            if "moveOffScreen" in cls:
                continue
            try:
                if loc.is_enabled(timeout=400):
                    return loc
            except Exception:
                pass
            # 后台/最小化时 is_enabled 可能误判，attached 即可 JS 填值
            return loc
        return None

    selectors = ['#i0118', 'input[name="passwd"]', 'input[type="password"]']
    for sel in selectors:
        loc = page.locator(sel).first
        if loc.count() and is_interactable(loc, fast=False):
            cls = loc.get_attribute("class") or ""
            if "moveOffScreen" not in cls:
                return loc
    for loc in page.locator('input[type="password"]').all():
        if not is_interactable(loc, fast=False):
            continue
        cls = loc.get_attribute("class") or ""
        if "moveOffScreen" in cls:
            continue
        return loc
    # JS 兜底：标记页面上真正可见的密码框
    try:
        marked = page.evaluate(
            """() => {
                document.querySelectorAll('[data-pw-target]').forEach(el => {
                    el.removeAttribute('data-pw-target');
                });
                const inputs = [...document.querySelectorAll(
                    'input[type="password"], input[name="passwd"]'
                )];
                const el = inputs.find(node => {
                    const cls = node.className || '';
                    if (cls.includes('moveOffScreen')) return false;
                    const r = node.getBoundingClientRect();
                    const st = window.getComputedStyle(node);
                    return r.width > 8 && r.height > 8
                        && st.visibility !== 'hidden'
                        && st.display !== 'none'
                        && !node.disabled;
                });
                if (el) {
                    el.setAttribute('data-pw-target', '1');
                    return true;
                }
                return false;
            }"""
        )
        if marked:
            loc = page.locator('[data-pw-target="1"]').first
            if loc.count():
                return loc
    except Exception:
        pass
    # 最后兜底：attached 的真实 passwd 框
    for sel in ['#i0118', 'input[name="passwd"]']:
        loc = page.locator(sel).first
        if loc.count():
            cls = loc.get_attribute("class") or ""
            if "moveOffScreen" not in cls:
                try:
                    if loc.is_enabled(timeout=500):
                        return loc
                except Exception:
                    continue
    return None


def password_field_value(page: Page) -> str:
    return read_password_value(page)


def _login_email_home_markers(text: str) -> bool:
    """Microsoft 登录「填邮箱」首页（尚未到密码页）。"""
    if not text:
        return False
    if any(
        m in text
        for m in (
            "パスワードの入力",
            "Enter password",
            "Enter your password",
            "输入密码",
            "输入你的密码",
        )
    ):
        return False
    return any(
        m in text
        for m in (
            "メール、電話、Skype",
            "Email, phone, or Skype",
            "邮箱、电话或 Skype",
            "邮箱、电话或Skype",
        )
    )


def is_empty_password_error(page: Page) -> bool:
    text = page_text(page)
    return any(
        p in text
        for p in [
            "パスワードを入力してください",
            "Microsoft アカウントのパスワードを入力",
            "Please enter your password",
            "Enter your password",
            "请输入你的密码",
            "输入你的密码",
        ]
    )


def is_password_entry_page(page: Page) -> bool:
    markers = [
        "パスワードの入力",
        "Enter password",
        "输入密码",
        "输入你的密码",
        "Microsoft アカウントのパスワード",
        "Enter your password",
    ]
    try:
        text = page_text(page)
    except Exception:
        text = ""
    # 仍在邮箱首页：绝不是密码页（DOM 里常藏着隐藏 password 框）
    if _login_email_home_markers(text):
        return False
    if any(p in text for p in markers):
        return True
    try:
        found = page.evaluate(
            f"""(markers) => {{
                const t = (document.body && document.body.innerText) || '';
                if (
                    (t.includes('メール、電話、Skype') || t.includes('Email, phone, or Skype'))
                    && !(t.includes('パスワードの入力') || t.includes('Enter password'))
                ) return false;
                if (markers.some(p => t.includes(p))) return true;
                const pick = {_JS_PICK_PASSWORD_INPUT};
                return !!pick();
            }}""",
            markers,
        )
        if found:
            return True
    except Exception:
        pass
    return False


def is_on_email_entry_page(page: Page) -> bool:
    # 辅助邮箱绑定页有邮箱输入框，绝不能当成登录邮箱页
    try:
        if is_recovery_bind_page(page) or is_code_verify_page(page):
            return False
    except Exception:
        pass
    try:
        text = page_text(page)
    except Exception:
        text = ""
    if _login_email_home_markers(text):
        return True
    if any(
        p in text
        for p in (
            "パスワードの入力",
            "Enter password",
            "Enter your password",
            "输入密码",
        )
    ):
        return False
    if find_email_input(page) is not None and find_password_input(page) is None:
        return True
    if find_email_input(page) is not None and not is_password_entry_page(page):
        return True
    try:
        return bool(
            page.evaluate(
                f"""() => {{
                    const t = (document.body && document.body.innerText) || '';
                    if (
                        (t.includes('メール、電話、Skype') || t.includes('Email, phone, or Skype'))
                        && !(t.includes('パスワードの入力') || t.includes('Enter password'))
                    ) return true;
                    const pick = {_JS_PICK_EMAIL_INPUT};
                    const pickPwd = {_JS_PICK_PASSWORD_INPUT};
                    return !!pick() && !pickPwd();
                }}"""
            )
        )
    except Exception:
        return False


def wait_password_page_ready(page: Page, timeout_sec: float = 6.0) -> bool:
    """等待密码页渲染完成（多开时 DOM 可能较慢）。"""
    if is_microsoft_login_page_loading(page):
        wait_for_login_transition(page, timeout_sec=min(timeout_sec + 4.0, 14.0))
    quick_dismiss_passkey(page)
    if not _keep_background():
        dismiss_passkey_or_other(page)
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if is_microsoft_login_page_loading(page):
            time.sleep(0.12)
            continue
        if find_password_input(page) is not None or is_password_entry_page(page):
            return True
        time.sleep(0.12)
    return find_password_input(page) is not None or is_password_entry_page(page)


def fill_password_via_keyboard(page: Page, password: str) -> bool:
    if _keep_background():
        return False
    target = password.strip()
    if not target:
        return False
    pwd_loc = find_password_input(page)
    if pwd_loc is None:
        return False
    try:
        pwd_loc.click(timeout=2000)
        page.keyboard.press("Control+a")
        page.keyboard.press("Backspace")
        page.keyboard.type(target, delay=20)
        safe_wait(page, 300)
        return password_field_value(page) == target
    except Exception:
        return False


def fill_password_robust(page: Page, password: str, max_attempts: int = 6) -> bool:
    """密码填表：后台 JS 优先，多开时不抢焦点。"""
    target = password.strip()
    if not target:
        return False
    bg = _keep_background()
    fast = _fast_fill_enabled()
    if bg:
        pause_ms = 60 if fast else 120
        tries = min(max_attempts, 3)
        for _ in range(tries):
            if fill_password_via_page_js(page, target):
                safe_wait(page, pause_ms)
                if read_password_value(page) == target:
                    return True
            safe_wait(page, pause_ms)
        return read_password_value(page) == target
    if not bg:
        ensure_browser_focus_for_input(page)
    quick_dismiss_passkey(page)
    dismiss_passkey_or_other(page)
    js_set = _js_set_native_value_expr()
    for attempt in range(max_attempts):
        if fill_password_via_page_js(page, target):
            safe_wait(page, 250)
            if read_password_value(page) == target:
                return True
        if not wait_password_page_ready(page, timeout_sec=3.5):
            safe_wait(page, 300)
            continue
        pwd_loc = find_password_input(page)
        if pwd_loc is not None:
            try:
                if pwd_loc.evaluate(js_set, target):
                    safe_wait(page, 250)
                    if read_password_value(page) == target:
                        return True
            except Exception:
                pass
            if fill_locator(pwd_loc, target, fast=fast):
                safe_wait(page, 250)
                if read_password_value(page) == target:
                    return True
        if fill_password_via_keyboard(page, target):
            return True
        safe_wait(page, 300 + attempt * 150)
    return read_password_value(page) == target


def fill_locator(loc, value: str, fast: bool = False) -> bool:
    fill_timeout = 800 if fast else 2500
    verify_timeout = 300 if fast else 800
    js_fill = _js_set_native_value_expr()
    if _keep_background():
        try:
            if loc.evaluate(js_fill, value.strip()):
                return True
        except Exception:
            pass
    try:
        if not fast and not _keep_background():
            loc.scroll_into_view_if_needed(timeout=1000)
    except Exception:
        pass
    try:
        loc.click(timeout=500 if fast else 1200)
        loc.fill(value, timeout=fill_timeout)
        if loc.input_value(timeout=verify_timeout).strip() == value.strip():
            return True
        loc.press_sequential(value, delay=12)
        if loc.input_value(timeout=verify_timeout).strip() == value.strip():
            return True
        loc.evaluate(js_fill, value)
        if loc.input_value(timeout=verify_timeout).strip() == value.strip():
            return True
        loc.fill(value, timeout=fill_timeout)
        return loc.input_value(timeout=verify_timeout).strip() == value.strip()
    except Exception:
        try:
            if loc.evaluate(js_fill, value.strip()):
                return True
            return loc.input_value(timeout=verify_timeout).strip() == value.strip()
        except Exception:
            return False


def click_email_account_tile(page: Page, email: str, fast: bool = False) -> None:
    """邮箱提交后，可能需要点选账号磁贴。"""
    if password_field_ready(page, None):
        return
    local = email.split("@")[0]
    tile_timeout = 400 if fast else 800
    wait_and_click(
        page,
        [
            f'div[role="button"]:has-text("{email}")',
            f'div.table:has-text("{email}")',
            f'.table-row:has-text("{email}")',
            f'div[data-test-id="account"]:has-text("{email}")',
            f'#tilesHolder :has-text("{email}")',
            f'div[role="listitem"]:has-text("{local}")',
        ],
        timeout_ms=tile_timeout,
    )
    try:
        page.get_by_text(email, exact=True).first.click(timeout=tile_timeout)
    except Exception:
        try:
            page.get_by_text(email, exact=False).first.click(timeout=tile_timeout)
        except Exception:
            pass
    if not fast:
        safe_wait(page, 400)


def quick_dismiss_passkey(page: Page) -> bool:
    """快速跳过通行密钥/使用密码链接，单次短超时。"""
    if _keep_background():
        try:
            return bool(
                page.evaluate(
                    """() => {
                        const el = document.querySelector('#idA_PWD_SwitchToPassword')
                            || document.querySelector('a#idA_PWD_SwitchToPassword');
                        if (el) { el.click(); return true; }
                        return false;
                    }"""
                )
            )
        except Exception:
            return False
    return wait_and_click(
        page,
        [
            '#idA_PWD_SwitchToPassword',
            'a#idA_PWD_SwitchToPassword',
            'a:has-text("Use your password")',
            'a:has-text("使用密码")',
            'a:has-text("パスワードを使用")',
        ],
        timeout_ms=250,
    )


def proceed_after_email(page: Page, email: str, fast: bool = False) -> bool:
    """邮箱提交后，处理账号选择/通行密钥，直到密码框出现。"""
    if password_field_ready(page, None):
        return True
    quick_dismiss_passkey(page)
    click_email_account_tile(page, email, fast=fast)
    if password_field_ready(page, None):
        return True
    if fast:
        quick_dismiss_passkey(page)
    else:
        dismiss_passkey_or_other(page)
    return wait_password_ready(page, timeout_sec=5 if fast else 8)


def wait_after_email_submit(
    page: Page, email: str, timeout_sec: float = 5.0
) -> tuple[bool, tuple[AccountStatus, str] | None]:
    """邮箱提交后并行等待：密码框出现立即继续，同时检测锁定/错误/辅助邮箱页。"""
    deadline = time.time() + timeout_sec
    dismissed_passkey = False
    clicked_tile = False
    poll = 0.04 if _keep_background() else 0.08
    loading_grace = 0.0
    max_loading_grace = 14.0 if _keep_background() else 10.0

    while time.time() < deadline + loading_grace:
        if is_microsoft_login_page_loading(page):
            loading_grace = min(max_loading_grace, loading_grace + 1.0)
            time.sleep(poll)
            continue
        if password_field_ready(page, None):
            minimize_automation_browser_now()
            return True, None
        if is_recovery_bind_page(page) or is_code_verify_page(page):
            return False, (AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱")
        err = detect_login_errors(page)
        if err:
            return False, err
        bound = detect_already_bound_email(page)
        if bound:
            return False, bound
        identity = detect_identity_verification(page)
        if identity:
            return False, identity

        if not dismissed_passkey:
            if quick_dismiss_passkey(page):
                dismissed_passkey = True
                time.sleep(0.03)
                continue
        if not clicked_tile:
            click_email_account_tile(page, email, fast=True)
            clicked_tile = True
            time.sleep(0.03)
            continue

        time.sleep(poll)

    if password_field_ready(page, None):
        minimize_automation_browser_now()
        return True, None
    if is_recovery_bind_page(page) or is_code_verify_page(page):
        return False, (AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱")
    err = detect_login_errors(page)
    if err:
        return False, err
    bound = detect_already_bound_email(page)
    if bound:
        return False, bound
    err = wait_and_detect_error(page, timeout_sec=2.5)
    if err:
        return False, err
    if is_microsoft_login_page_loading(page):
        wait_for_login_transition(page, timeout_sec=12.0 if _keep_background() else 8.0)
        if password_field_ready(page, None):
            minimize_automation_browser_now()
            return True, None
        if is_recovery_bind_page(page) or is_code_verify_page(page):
            return False, (AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱")
    return False, None


def password_field_ready(page: Page, _selectors: list[str] | None = None) -> bool:
    if find_password_input(page) is not None:
        return True
    if is_password_entry_page(page):
        return True
    try:
        return bool(page.evaluate(f"() => !!({_JS_PICK_PASSWORD_INPUT})()"))
    except Exception:
        return False


def wait_login_form(page: Page, timeout_sec: int = 30) -> bool:
    """等待 Microsoft 登录表单渲染完成。"""
    deadline = time.time() + timeout_sec
    poll = 0.08 if _fast_fill_enabled() else 0.12
    while time.time() < deadline:
        if find_email_input(page) or find_password_input(page):
            return True
        if _keep_background():
            try:
                if page.evaluate(
                    f"""() => {{
                        const pickEmail = {_JS_PICK_EMAIL_INPUT};
                        const pickPwd = {_JS_PICK_PASSWORD_INPUT};
                        return !!(pickEmail() || pickPwd());
                    }}"""
                ):
                    return True
            except Exception:
                pass
        click_account_picker_other(page)
        time.sleep(poll)
    return False


def is_stuck_on_login_email_page(page: Page) -> bool:
    """卡在 Microsoft 登录邮箱首页（网卡半加载常见）。"""
    try:
        if is_recovery_bind_page(page) or is_code_verify_page(page):
            return False
    except Exception:
        pass
    try:
        if is_password_entry_page(page):
            return False
    except Exception:
        pass
    try:
        if is_on_email_entry_page(page):
            return True
    except Exception:
        pass
    try:
        text = page_text(page) or ""
        if _login_email_home_markers(text):
            return True
        url = (page.url or "").lower()
        on_ms = any(
            x in url
            for x in (
                "login.live.com",
                "login.microsoftonline.com",
                "account.live.com",
                "signup.live.com",
            )
        )
        if on_ms and ("サインイン" in text or "Sign in" in text or "登录" in text):
            if find_email_input(page) is not None and find_password_input(page) is None:
                return True
    except Exception:
        pass
    return False


def is_login_page_laggy(page: Page) -> bool:
    """
    登录页卡顿/半加载（图三：Microsoft Logo 裂成 Mic）。
    静默模式若关闭图片加载，不把「全图裂图」当成卡顿。
    """
    try:
        return bool(
            page.evaluate(
                """() => {
                    const url = (location.href || '').toLowerCase();
                    if (!['login.live.com','login.microsoftonline.com','account.live.com','signup.live.com']
                        .some(h => url.includes(h))) return false;
                    const t = ((document.body && document.body.innerText) || '')
                        .replace(/\\s+/g, ' ').trim();
                    const hasSignIn = /サインイン|Sign in|登录/.test(t);
                    if (!hasSignIn && !/メール、電話|Email, phone/.test(t)) return false;
                    if (document.readyState !== 'complete') return true;
                    const imgs = Array.from(document.images || []);
                    const broken = imgs.filter(img => !img.complete || img.naturalWidth === 0);
                    // 静默开了 imagesEnabled=false 时几乎全裂，不能当网卡
                    const likelyImagesOff = imgs.length >= 2 && broken.length === imgs.length;
                    if (!likelyImagesOff) {
                        for (const img of imgs) {
                            const hint = ((img.alt || '') + ' ' + (img.src || '') + ' ' + (img.className || '')).toLowerCase();
                            if (/microsoft|logo|mslogo|bannerlogo|img-logo/.test(hint)) {
                                if (!img.complete || img.naturalWidth === 0) return true;
                            }
                        }
                        // 页面上只见残缺 "Mic" 且正文极短
                        if (hasSignIn && t.length < 100 && /\\bMic\\b/.test(t) && !/Microsoft/i.test(t)) {
                            return true;
                        }
                    }
                    const emailEl = document.querySelector(
                        '#i0116, input[name="loginfmt"], #usernameEntry, input[type="email"]'
                    );
                    if (emailEl && (emailEl.disabled || emailEl.getAttribute('aria-disabled') === 'true')) {
                        return true;
                    }
                    return false;
                }"""
            )
        )
    except Exception:
        return False


def recover_network_stuck_email_page(
    page: Page,
    *,
    email: str = "",
    refreshes: int = 2,
    fill_email: bool = True,
) -> bool:
    """
    网卡卡在「メール、電話、Skype」邮箱输入页时刷新（默认两下），
    并尽量重新带出登录表单；fill_email=True 时预填邮箱。
    """
    if not (
        is_stuck_on_login_email_page(page)
        or is_login_page_laggy(page)
        or is_on_email_entry_page(page)
    ):
        return False
    n = max(1, int(refreshes))
    for i in range(n):
        log_step(f"登录邮箱页卡顿，刷新 ({i + 1}/{n})")
        try:
            page.reload(wait_until="domcontentloaded", timeout=20000)
        except Exception as exc:
            log_step(f"刷新失败，尝试重新打开当前页: {exc}")
            try:
                cur = page.url or ""
                if cur:
                    page.goto(cur, wait_until="domcontentloaded", timeout=20000)
            except Exception:
                pass
        safe_wait(page, 700)
        try:
            click_account_picker_other(page)
        except Exception:
            pass
        wait_login_form(page, timeout_sec=8)
        if fill_email and (email or "").strip():
            try:
                fill_email_robust(page, email)
            except Exception:
                pass
    return bool(find_email_input(page) or is_stuck_on_login_email_page(page))


def ensure_login_form_ready(
    page: Page,
    login_url: str = "",
    timeout_sec: int | None = None,
    max_refreshes: int = 2,
) -> bool:
    """等待登录表单；未加载时刷新页面并重试。"""
    if timeout_sec is None:
        timeout_sec = int(_RUNTIME.get("login_form_wait_sec", 6))
    fast = _fast_fill_enabled()
    refresh_wait = max(4, timeout_sec)
    attach_state = "attached" if _keep_background() else "visible"

    if wait_login_form(page, timeout_sec=timeout_sec):
        return True

    for attempt in range(max_refreshes):
        log_step(f"登录表单未出现，刷新页面 ({attempt + 1}/{max_refreshes})")
        fallback_url = login_url
        if not fallback_url:
            try:
                current = page.url or ""
                if any(
                    x in current.lower()
                    for x in [
                        "login.microsoftonline.com",
                        "login.live.com",
                        "account.live.com",
                    ]
                ):
                    fallback_url = current
            except Exception:
                pass
        try:
            page.reload(wait_until="domcontentloaded", timeout=15000 if fast else 20000)
        except Exception:
            if fallback_url:
                try:
                    page.goto(
                        fallback_url,
                        wait_until="domcontentloaded",
                        timeout=15000 if fast else 20000,
                    )
                except Exception:
                    pass
        safe_wait(page, 300 if fast else 600)
        click_account_picker_other(page)
        if wait_login_form(page, timeout_sec=refresh_wait):
            return True

    try:
        page.wait_for_selector(
            '#i0116, input[name="loginfmt"], input[name="passwd"], input[type="password"]',
            state=attach_state,
            timeout=4000 if fast else 8000,
        )
        return bool(find_email_input(page) or find_password_input(page))
    except Exception:
        return False


def wait_password_ready(page: Page, timeout_sec: int = 8) -> bool:
    """等待真实密码框或「使用密码」链接出现；中途若出现错误页则立即停止。"""
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        # 已到密码页就成功；勿被网卡误判打断
        if password_field_ready(page, None) or is_password_entry_page(page):
            return True
        err = detect_login_errors(page)
        if err and err[0] != AccountStatus.NETWORK_CARD:
            return False
        if quick_dismiss_passkey(page):
            time.sleep(0.08)
            continue
        time.sleep(0.06)
    return password_field_ready(page, None) or is_password_entry_page(page)


def click_account_picker_other(page: Page) -> None:
    """账号选择页：点「使用其他账号」。"""
    wait_and_click(
        page,
        [
            "#otherTileText",
            'div[data-test-id="otherTile"]',
            'a:has-text("Use another account")',
            'div:has-text("Use another account")',
            'a:has-text("使用其他帐户")',
            'a:has-text("使用另一个帐户")',
            'div:has-text("使用另一个帐户")',
            'a:has-text("別のアカウントを使用")',
            'div:has-text("別のアカウントを使用")',
        ],
        timeout_ms=1000,
    )


def dismiss_passkey_or_other(page: Page) -> None:
    """跳过通行密钥/其他登录方式，尽量走密码登录。"""
    wait_and_click(
        page,
        [
            'a:has-text("Use your password")',
            'span:has-text("Use your password")',
            'a:has-text("使用密码")',
            'a:has-text("パスワードを使用")',
            '#idA_PWD_SwitchToPassword',
            'a#idA_PWD_SwitchToPassword',
            'button:has-text("Other ways to sign in")',
            'a:has-text("其他方式登录")',
            'a:has-text("別の方法でサインイン")',
        ],
        timeout_ms=1000,
    )


def is_visible(page: Page, selector: str, timeout_ms: int = 800) -> bool:
    try:
        return page.locator(selector).first.is_visible(timeout=timeout_ms)
    except Exception:
        return False


def auto_login(
    page: Page,
    email: str,
    password: str,
    *,
    _stuck_refreshed: bool = False,
) -> tuple[AccountStatus, str]:
    fast = _fast_fill_enabled()
    bg = _keep_background()
    retry_ms = 50 if (fast and bg) else (200 if fast else 600)
    after_next_ms = 80 if (fast and bg) else (150 if fast else 350)
    email_submit_wait = 10.0 if bg else (5.0 if fast else 9.0)
    err_detect_wait = 1.0 if (fast and bg) else (1.5 if fast else 3.0)

    def _retry_after_email_stuck(detail: str) -> tuple[AccountStatus, str]:
        if _stuck_refreshed or not is_stuck_on_login_email_page(page):
            return AccountStatus.UNKNOWN, detail
        recover_network_stuck_email_page(page, email=email, refreshes=2, fill_email=True)
        return auto_login(page, email, password, _stuck_refreshed=True)

    # 已在辅助邮箱绑定页：不要再当登录邮箱页去填
    try:
        if is_phone_verify_page(page):
            return AccountStatus.NEED_PHONE, "需要电话认证（确认手机号末位，非辅助邮箱流程）"
        if is_definitely_recovery_flow_page(page):
            return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"
        bound = detect_already_bound_email(page)
        if bound:
            return bound
        identity = detect_identity_verification(page)
        if identity:
            return identity
    except Exception:
        pass

    # 图三：登录页卡顿（Logo 裂图等）→ 先刷新 1～2 次再填邮箱，绝不提前领辅助邮箱
    try:
        if (not _stuck_refreshed) and (
            is_login_page_laggy(page)
            or is_microsoft_login_page_loading(page)
            or (
                is_on_email_entry_page(page)
                and is_stuck_on_login_email_page(page)
                and is_login_page_laggy(page)
            )
        ):
            log_step("检测到登录页卡顿，先刷新再填邮箱")
            recover_network_stuck_email_page(
                page, email="", refreshes=2, fill_email=False
            )
    except Exception:
        pass

    click_account_picker_other(page)
    if not find_email_input(page) and not find_password_input(page):
        attach_state = "attached" if _keep_background() else "visible"
        try:
            page.wait_for_selector(
                '#i0116, input[name="loginfmt"]',
                state=attach_state,
                timeout=6000 if fast else 10000,
            )
        except Exception:
            if is_definitely_recovery_flow_page(page):
                return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"
            if not ensure_login_form_ready(
                page,
                login_url="",
                timeout_sec=6 if fast else 10,
            ):
                if (not _stuck_refreshed) and (
                    is_stuck_on_login_email_page(page) or is_login_page_laggy(page)
                ):
                    recover_network_stuck_email_page(
                        page, email="", refreshes=2, fill_email=False
                    )
                    return auto_login(page, email, password, _stuck_refreshed=True)
                return AccountStatus.UNKNOWN, "登录页表单未加载"

    password_selectors = ['#i0118', 'input[name="passwd"]', 'input[type="password"]']
    next_selectors = [
        '#idSIButton9',
        'input[type="submit"]',
        'button[type="submit"]',
        'button:has-text("次へ")',
        'input[value="次へ"]',
    ]

    # 仍在「メール、電話、Skype」首页时必须先填邮箱，绝不可因隐藏密码框跳去填密码
    try:
        page_blob = page_text(page)
    except Exception:
        page_blob = ""
    on_email_home = is_on_email_entry_page(page) or _login_email_home_markers(page_blob)
    on_pwd_page = is_password_entry_page(page)

    if on_email_home and not on_pwd_page:
        log_step("填写邮箱")
        if not fill_email_robust(page, email):
            if is_on_email_entry_page(page) or _login_email_home_markers(page_text(page)):
                return _retry_after_email_stuck("未能填入邮箱")
        if read_email_value(page) != email.strip():
            # 仍在邮箱页空着 → 邮箱没填上
            if is_on_email_entry_page(page) or _login_email_home_markers(page_text(page)):
                return _retry_after_email_stuck("未能填入邮箱")
        if not (password_field_ready(page, password_selectors) and is_password_entry_page(page)):
            minimize_automation_browser_now()
            log_step("点击下一步")
            click_login_submit(page, next_selectors)
            if not bg:
                try:
                    page.keyboard.press("Enter")
                except Exception:
                    pass
            safe_wait(page, after_next_ms)
            if not fast and not bg:
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=5000)
                except Exception:
                    pass
            ready, err = wait_after_email_submit(page, email, timeout_sec=email_submit_wait)
            if err:
                return err
            if not ready and not proceed_after_email(page, email, fast=fast):
                bound = detect_already_bound_email(page)
                if bound:
                    return bound
                err = detect_login_errors(page) or wait_and_detect_error(page, timeout_sec=err_detect_wait)
                if err:
                    return err
                bound = detect_already_bound_email(page)
                if bound:
                    return bound
                if wait_password_page_ready(page, timeout_sec=8.0 if bg else 6.0):
                    pass
                elif password_field_ready(page, password_selectors) and is_password_entry_page(page):
                    pass
                elif is_microsoft_login_page_loading(page):
                    wait_for_login_transition(page, timeout_sec=14.0 if bg else 10.0)
                    bound = detect_already_bound_email(page)
                    if bound:
                        return bound
                    if wait_password_page_ready(page, timeout_sec=6.0):
                        pass
                    elif not (
                        password_field_ready(page, password_selectors)
                        and is_password_entry_page(page)
                    ):
                        bound = detect_already_bound_email(page)
                        if bound:
                            return bound
                        if is_on_email_entry_page(page) or _login_email_home_markers(page_text(page)):
                            return _retry_after_email_stuck("未能填入邮箱或点击下一步无效")
                        return _retry_after_email_stuck("点击下一步后未进入密码页")
                elif is_on_email_entry_page(page) or _login_email_home_markers(page_text(page)):
                    bound = detect_already_bound_email(page)
                    if bound:
                        return bound
                    if is_recovery_bind_page(page):
                        return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"
                    return _retry_after_email_stuck("未能填入邮箱或点击下一步无效")
                else:
                    bound = detect_already_bound_email(page)
                    if bound:
                        return bound
                    if is_recovery_bind_page(page):
                        return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱"
                    if wait_password_page_ready(page, timeout_sec=5.0) or is_password_entry_page(page):
                        pass
                    elif is_on_email_entry_page(page) or _login_email_home_markers(page_text(page)):
                        return _retry_after_email_stuck("未能填入邮箱或点击下一步无效")
                    else:
                        return _retry_after_email_stuck("点击下一步后未进入密码页")
    elif not (password_field_ready(page, password_selectors) and is_password_entry_page(page)):
        if is_password_entry_page(page) or wait_password_page_ready(
            page, timeout_sec=5.0 if (fast and bg) else (6.0 if fast else 10.0)
        ):
            pass
        else:
            err = detect_login_errors(page)
            if err:
                return err
            if is_on_email_entry_page(page) or _login_email_home_markers(page_text(page)):
                return _retry_after_email_stuck("未能填入邮箱")
            return AccountStatus.UNKNOWN, "未找到邮箱或密码输入框"

    minimize_automation_browser_now()

    def _still_on_email_home() -> bool:
        try:
            return is_on_email_entry_page(page) or _login_email_home_markers(page_text(page))
        except Exception:
            return False

    if not password_field_ready(page, password_selectors) or not is_password_entry_page(page):
        context = page.context
        page = resolve_flow_page_for_password(page, context, page.url, fast)
    if _still_on_email_home():
        return _retry_after_email_stuck("未能填入邮箱或未能进入密码页（仍在邮箱输入页）")
    if not bg:
        ensure_browser_focus_for_input(page)
        quick_dismiss_passkey(page)
        dismiss_passkey_or_other(page)

    on_password_page = password_field_ready(page, password_selectors) and is_password_entry_page(page)
    if not on_password_page and not bg:
        quick_dismiss_passkey(page)
        pwd_loc = find_password_input(page)
        if pwd_loc is None:
            safe_wait(page, retry_ms)
            pwd_loc = find_password_input(page)
        if pwd_loc is None:
            wait_password_page_ready(page, timeout_sec=6.0 if fast else 10.0)
        on_password_page = password_field_ready(page, password_selectors) and is_password_entry_page(page)
    if not on_password_page:
        if is_microsoft_login_page_loading(page):
            wait_for_login_transition(page, timeout_sec=14.0 if bg else 10.0)
            on_password_page = password_field_ready(page, password_selectors) and is_password_entry_page(page)
        err = detect_login_errors(page)
        if err:
            return err
        if _still_on_email_home():
            return _retry_after_email_stuck("未能填入邮箱或未能进入密码页（仍在邮箱输入页）")
        if not on_password_page:
            return AccountStatus.UNKNOWN, "未找到可交互的密码框"

    log_step("填写密码")
    pwd_attempts = 2 if bg else 5
    if not fill_password_robust(page, password, max_attempts=pwd_attempts):
        if _still_on_email_home():
            return _retry_after_email_stuck("未能填入邮箱或未能进入密码页（仍在邮箱输入页）")
        return AccountStatus.UNKNOWN, "未能填入密码"

    if read_password_value(page) != password.strip():
        if not fill_password_robust(page, password, max_attempts=2 if bg else 3):
            if _still_on_email_home():
                return _retry_after_email_stuck("未能填入邮箱或未能进入密码页（仍在邮箱输入页）")
            return AccountStatus.UNKNOWN, "未能填入密码"

    if read_password_value(page) != password.strip():
        if _still_on_email_home():
            return _retry_after_email_stuck("未能填入邮箱或未能进入密码页（仍在邮箱输入页）")
        return AccountStatus.UNKNOWN, "未能填入密码"

    log_step("点击登录")
    click_login_submit(page, next_selectors)
    safe_wait(page, 200 if bg else (350 if fast else 500))

    if is_empty_password_error(page) or (
        is_password_entry_page(page) and read_password_value(page) != password.strip()
    ):
        log_step("密码未写入，重试填表")
        if fill_password_robust(page, password, max_attempts=2 if bg else 3):
            click_login_submit(page, next_selectors)
            safe_wait(page, 200 if bg else (350 if fast else 500))

    if is_password_entry_page(page) and read_password_value(page) != password.strip():
        if _still_on_email_home():
            return _retry_after_email_stuck("未能填入邮箱或未能进入密码页（仍在邮箱输入页）")
        return AccountStatus.UNKNOWN, "未能填入密码"

    status, detail = wait_for_post_login_outcome(
        page,
        timeout_sec=28 if bg else (16 if fast else 20),
        email=email,
    )
    if (
        (not _stuck_refreshed)
        and status == AccountStatus.UNKNOWN
        and is_stuck_on_login_email_page(page)
    ):
        recover_network_stuck_email_page(page, email=email, refreshes=2)
        return auto_login(page, email, password, _stuck_refreshed=True)
    return status, detail


def resolve_recovery_credentials(
    machine_id: str,
    recovery_email: str,
    recovery_password: str,
    recovery_config: dict | None,
) -> tuple[str, str]:
    """已有辅助邮箱则直接用；否则从任务池 claim_recovery 领取一条。"""
    email = (recovery_email or "").strip()
    password = recovery_password or ""
    if email:
        return email, password
    claim_fn = (recovery_config or {}).get("claim_recovery")
    if not callable(claim_fn):
        return "", ""
    try:
        claimed = claim_fn(machine_id)
    except Exception as exc:
        log_step(f"领取辅助邮箱失败: {exc}")
        return "", ""
    if not claimed:
        return "", ""
    if isinstance(claimed, (tuple, list)) and len(claimed) >= 1:
        return (claimed[0] or "").strip(), (claimed[1] if len(claimed) > 1 else "") or ""
    return "", ""


def release_recovery_if_needed(
    machine_id: str,
    recovery_config: dict | None,
    *,
    reason: str = "",
) -> None:
    release_fn = (recovery_config or {}).get("release_recovery")
    if not callable(release_fn):
        return
    try:
        release_fn(machine_id, reason)
    except TypeError:
        try:
            release_fn(machine_id)
        except Exception as exc:
            log_step(f"归还辅助邮箱失败: {exc}")
    except Exception as exc:
        log_step(f"归还辅助邮箱失败: {exc}")


def _page_looks_like_recovery_bind(page: Page) -> bool:
    """宽松识别「可填辅助邮箱」页（超时前最后兜底用）。"""
    # 优先：跨 frame 找 someone@example.com 输入框（最可靠）
    try:
        if find_recovery_email_input(page) is not None:
            return True
    except Exception:
        pass
    try:
        if is_definitely_recovery_flow_page(page) or is_recovery_bind_page(page):
            return True
    except Exception:
        pass
    try:
        text = page_text(page) or ""
        text_l = text.lower()
    except Exception:
        return False
    markers = (
        "アカウントを保護しましょう",
        "アカウントを保護",
        "アカウントの保護",
        "help us protect your account",
        "let's protect your account",
        "someone@example.com",
        "メールの追加",
        "add email",
    )
    if any(m in text or m in text_l for m in markers):
        return True
    try:
        for frame in getattr(page, "frames", []) or [page]:
            try:
                if frame.locator('input[placeholder="someone@example.com"]').first.is_visible(
                    timeout=300
                ):
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def wait_for_auth_page_settled(
    page: Page,
    context,
    login_url: str,
    *,
    timeout_sec: float = 50.0,
) -> tuple[AccountStatus, str, Page]:
    """
    登录提交后等待 Microsoft 页从半加载壳过渡到可判定/可绑定状态。
    避免页面还没出来就判「未知状态」或「需要绑定但未领辅助邮箱」。
    """
    timeout_sec = max(40.0, float(timeout_sec))
    deadline = time.time() + timeout_sec
    grace_deadline = deadline + 18.0  # 超时后再宽限一会，专等绑定页慢加载
    last_log = 0.0
    did_refresh = False

    while time.time() < grace_deadline:
        in_grace = time.time() >= deadline
        try:
            page = get_active_flow_page(context, login_url, timeout_sec=4)
        except Exception:
            pass

        # 绑定页优先：iframe 里表单已出时，主框仍可能被误判为半加载壳
        if _page_looks_like_recovery_bind(page):
            return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱", page
        try:
            if find_recovery_email_input(page) is not None:
                return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱", page
        except Exception:
            pass

        if is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
            now = time.time()
            if now - last_log > 2.0:
                log_step(
                    "Microsoft 页面半加载，继续等待…"
                    + ("（宽限期）" if in_grace else "")
                )
                last_log = now
            remain = max(0.5, grace_deadline - time.time())
            wait_for_login_transition(page, timeout_sec=min(8.0, remain))
            if (not did_refresh) and (not in_grace) and (deadline - time.time() < timeout_sec * 0.4):
                try:
                    page.reload(wait_until="domcontentloaded", timeout=15000)
                    did_refresh = True
                    log_step("半加载过久，已刷新页面")
                except Exception:
                    pass
            safe_wait(page, 400)
            continue

        err = detect_login_errors(page)
        if err:
            return err[0], err[1], page

        try:
            if is_phone_verify_page(page):
                return (
                    AccountStatus.NEED_PHONE,
                    "需要电话认证（确认手机号末位，非辅助邮箱流程）",
                    page,
                )
        except Exception:
            pass

        identity = detect_identity_verification(page)
        if identity:
            return identity[0], identity[1], page

        bound = detect_already_bound_email(page)
        if bound:
            return bound[0], bound[1], page

        url = (page.url or "").lower()
        if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
            s, d = evaluate_inbox_account_status(page)
            return s, d, page

        # 绑定页：一旦可识别立即返回（不要再空等到超时）
        if _page_looks_like_recovery_bind(page):
            return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱", page

        if is_code_verify_page(page):
            return AccountStatus.WAIT_CODE, "等待验证码", page

        status, detail = detect_status(page)
        if status == AccountStatus.STAY_SIGNED_IN:
            click_stay_signed_in(page)
            safe_wait(page, 800)
            continue
        if status == AccountStatus.NEED_RECOVERY:
            return status, detail, page
        if status not in {AccountStatus.UNKNOWN}:
            return status, detail, page

        if (not in_grace) and (is_stuck_on_login_email_page(page) or is_login_page_laggy(page)):
            recover_network_stuck_email_page(page, email="", refreshes=1, fill_email=False)
            continue

        now = time.time()
        if now - last_log > 2.5:
            log_step("等待登录后续页面…" + ("（宽限期等绑定页）" if in_grace else ""))
            last_log = now
        safe_wait(page, 400)

    try:
        page = get_active_flow_page(context, login_url, timeout_sec=4)
    except Exception:
        pass
    try:
        if is_phone_verify_page(page):
            return (
                AccountStatus.NEED_PHONE,
                "需要电话认证（确认手机号末位，非辅助邮箱流程）",
                page,
            )
    except Exception:
        pass
    if _page_looks_like_recovery_bind(page) or (
        find_recovery_email_input(page) is not None
    ):
        return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱（超时前已出现绑定页）", page
    if is_code_verify_page(page):
        return AccountStatus.WAIT_CODE, "等待验证码（超时前已到验证码页）", page
    if is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
        # 半加载结束前再多等几秒，避免刚出绑定页就被判超时
        log_step("半加载壳：额外等待绑定页出现…")
        extra_end = time.time() + 12.0
        while time.time() < extra_end:
            safe_wait(page, 500)
            try:
                page = get_active_flow_page(context, login_url, timeout_sec=2)
            except Exception:
                pass
            if _page_looks_like_recovery_bind(page) or (
                find_recovery_email_input(page) is not None
            ):
                return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱", page
            if not (is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page)):
                break
        if _page_looks_like_recovery_bind(page) or (
            find_recovery_email_input(page) is not None
        ):
            return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱", page
        return AccountStatus.TIMEOUT, "Microsoft 页面加载超时（半加载壳）", page
    status, detail = detect_status(page)
    if status == AccountStatus.NEED_RECOVERY or _page_looks_like_recovery_bind(page):
        return AccountStatus.NEED_RECOVERY, "需要绑定辅助邮箱", page
    if status == AccountStatus.UNKNOWN and any(
        x in (page.url or "").lower()
        for x in ("login.microsoftonline.com", "login.live.com", "account.live.com")
    ):
        return AccountStatus.TIMEOUT, "Microsoft 验证流程等待超时", page
    return status, detail, page


def wait_for_final_status(
    page: Page,
    timeout_sec: int,
    manual_mfa_timeout_sec: int,
    recovery_email: str = "",
    recovery_password: str = "",
    recovery_config: dict | None = None,
    machine_id: str = "",
) -> tuple[AccountStatus, str]:
    start = time.time()
    manual_prompted = False
    recovery_config = recovery_config or {}
    resolved_email = (recovery_email or "").strip()
    resolved_password = recovery_password or ""

    while time.time() - start < timeout_sec:
        status, detail = detect_status(page)

        if status == AccountStatus.OK:
            if is_outlook_inbox(page):
                return evaluate_inbox_account_status(page)
            return status, detail

        if status == AccountStatus.STAY_SIGNED_IN:
            click_stay_signed_in(page)
            safe_wait(page, 2000)
            continue

        if status in {
            AccountStatus.ACCOUNT_NOT_FOUND,
            AccountStatus.BAD_PASSWORD,
            AccountStatus.BAD_CREDENTIALS,
            AccountStatus.LOCKED,
            AccountStatus.NEED_IDENTITY,
            AccountStatus.NEED_PHONE,
            AccountStatus.ALREADY_BOUND,
            AccountStatus.NETWORK_CARD,
        }:
            return status, detail

        bound = detect_already_bound_email(page)
        if bound:
            return bound

        try:
            if is_phone_verify_page(page):
                return (
                    AccountStatus.NEED_PHONE,
                    "需要电话认证（确认手机号末位，非辅助邮箱流程）",
                )
        except Exception:
            pass

        on_recovery = (
            not is_identity_verification_page(page)
            and not is_phone_verify_page(page)
            and is_definitely_recovery_flow_page(page)
        )
        if on_recovery:
            if not resolved_email:
                resolved_email, resolved_password = resolve_recovery_credentials(
                    machine_id, resolved_email, resolved_password, recovery_config
                )
            if not resolved_email:
                return (
                    AccountStatus.NEED_RECOVERY,
                    "需要绑定辅助邮箱，但辅助邮箱池已空（请在左侧补充）",
                )
            log_step(f"绑定辅助邮箱: {resolved_email}")
            ok, bind_detail = try_bind_recovery_on_page(
                page,
                resolved_email,
                resolved_password,
                recovery_config.get("recovery_imap") or {},
                recovery_config,
                log=log_step,
            ) or (False, "绑定失败")
            manual = bool(recovery_config.get("recovery_manual_code", True))
            if ok and manual:
                return AccountStatus.WAIT_CODE, bind_detail
            if ok:
                return AccountStatus.RECOVERY_BOUND, bind_detail
            release_recovery_if_needed(
                machine_id, recovery_config, reason="绑定失败归还"
            )
            if "电话认证" in (bind_detail or ""):
                return AccountStatus.NEED_PHONE, bind_detail
            return AccountStatus.RECOVERY_FAILED, bind_detail

        # 状态写了需要绑定但绑定页可能仍在加载：继续等，不要立刻结束
        if status == AccountStatus.NEED_RECOVERY and not is_definitely_recovery_flow_page(
            page
        ):
            if is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
                wait_for_login_transition(page, timeout_sec=6)
                safe_wait(page, 500)
                continue
            if is_recovery_bind_page(page):
                safe_wait(page, 800)
                continue
            if is_stuck_on_login_email_page(page):
                recover_network_stuck_email_page(page, email="", refreshes=2, fill_email=False)
                continue
            # 给绑定页再留一点观察时间（总超时由外层 timeout_sec 控制）
            if time.time() - start < timeout_sec * 0.85:
                safe_wait(page, 800)
                continue
            return status, detail + "（绑定页未完全加载，未领取辅助邮箱）"

        if status in {AccountStatus.NEED_VERIFY, AccountStatus.CAPTCHA, AccountStatus.UNKNOWN}:
            if is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
                wait_for_login_transition(page, timeout_sec=6)
                safe_wait(page, 500)
                continue
            if status == AccountStatus.UNKNOWN and "加载" in (detail or ""):
                safe_wait(page, 800)
                continue
            if not manual_prompted and status in {
                AccountStatus.NEED_VERIFY,
                AccountStatus.CAPTCHA,
            }:
                manual_prompted = True
                if not sys.stdin.isatty():
                    # 非交互环境无法等待回车，直接返回当前状态
                    return status, detail + "（非交互模式，未等待人工验证）"
                print(f"    [!] 检测到 {status.value}，请在浏览器中手动完成，完成后按回车...")
                try:
                    input()
                except KeyboardInterrupt:
                    raise
                extra_start = time.time()
                while time.time() - extra_start < manual_mfa_timeout_sec:
                    status2, detail2 = detect_status(page)
                    if status2 == AccountStatus.STAY_SIGNED_IN:
                        click_stay_signed_in(page)
                        safe_wait(page, 2000)
                        continue
                    if status2 == AccountStatus.OK:
                        return status2, detail2
                    if status2 in {
                        AccountStatus.ACCOUNT_NOT_FOUND,
                        AccountStatus.BAD_PASSWORD,
                        AccountStatus.BAD_CREDENTIALS,
                        AccountStatus.LOCKED,
                    }:
                        return status2, detail2
                    page.wait_for_timeout(2000)
                return AccountStatus.TIMEOUT, "手动验证后仍未进入邮箱"

        safe_wait(page, 1500)

    return AccountStatus.TIMEOUT, "等待登录结果超时"


def focus_browser_for_debug(port: int, login_url: str) -> tuple[bool, str]:
    """通过 CDP 将真实浏览器窗口置前；优先停在当前问题页，不强制跳回登录首页。"""
    register_automation_browser(port)
    playwright = _start_sync_playwright()
    parts: list[str] = []
    ok = False
    try:
        browser = playwright.chromium.connect_over_cdp(
            f"http://127.0.0.1:{port}",
            timeout=int(_RUNTIME["cdp_connect_timeout_ms"]),
        )
        _cdp_ignore_certificate_errors(browser)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page: Page | None = None
        url_markers = (
            "login.microsoftonline.com",
            "login.live.com",
            "account.live.com",
            "account.microsoft.com",
            "outlook.live.com",
            "outlook.office.com",
            "signup.live.com",
        )
        # 优先当前 Microsoft 流程页（绑定辅助邮箱 / 验证码 / 已绑确认等）
        # 先按 URL，再按页面内容识别辅助邮箱页
        for candidate in reversed(context.pages):
            try:
                url = (candidate.url or "").lower()
            except Exception:
                continue
            if any(m in url for m in url_markers):
                page = candidate
                break
        if page is None:
            for candidate in reversed(context.pages):
                try:
                    if is_recovery_bind_page(candidate) or is_code_verify_page(candidate):
                        page = candidate
                        break
                except Exception:
                    continue
        if page is None and context.pages:
            page = context.pages[-1]
        if page is None:
            # 只有完全没有标签时才新建；且仅当调用方传了 login_url
            if login_url:
                page = context.new_page()
                try:
                    safe_page_goto(page, login_url, timeout_ms=25000, retries=3)
                    parts.append("无标签，已打开登录页")
                except Exception as exc:
                    parts.append(f"无标签打开登录页失败: {str(exc)[:100]}")
                    shown = show_hubstudio_windows_for_debug(debug_port=port, log=log_step)
                    return shown > 0, " | ".join(parts + ([f"Win32显示{shown}"] if shown else []))
            else:
                parts.append("无标签且禁止跳转登录页")
                shown = show_hubstudio_windows_for_debug(debug_port=port, log=log_step)
                return shown > 0, " | ".join(parts + ([f"Win32显示{shown}" ] if shown else ["无可用标签"]))
        else:
            # 已有会话页：不要 goto 登录首页，否则用户还要再登一次
            parts.append("已定位到当前页面（未重新登录）")
            try:
                parts.append((page.url or "")[:100])
            except Exception:
                pass

        try:
            page.bring_to_front()
        except Exception:
            pass

        try:
            browser_cdp = browser.new_browser_cdp_session()
            targets = browser_cdp.send("Target.getTargets")
            page_url = (page.url or "").lower()
            target_id = None
            for info in targets.get("targetInfos", []):
                if info.get("type") != "page":
                    continue
                turl = (info.get("url") or "").lower()
                if turl == page_url or any(m in turl for m in url_markers):
                    target_id = info.get("targetId")
                    break
            if not target_id:
                for info in targets.get("targetInfos", []):
                    if info.get("type") == "page":
                        target_id = info.get("targetId")
                        break
            if target_id:
                win = browser_cdp.send(
                    "Browser.getWindowForTarget", {"targetId": target_id}
                )
                window_id = win.get("windowId")
                if window_id is not None:
                    browser_cdp.send(
                        "Browser.setWindowBounds",
                        {
                            "windowId": window_id,
                            "bounds": {
                                "left": 80,
                                "top": 60,
                                "width": 1280,
                                "height": 860,
                                "windowState": "normal",
                            },
                        },
                    )
                    ok = True
                    parts.append("CDP 已恢复浏览器窗口")
        except Exception as exc:
            parts.append(f"CDP 置窗: {exc}")

        shown = show_hubstudio_windows_for_debug(debug_port=port, log=log_step)
        if shown:
            ok = True
            parts.append(f"Win32 已显示 {shown} 个浏览器窗口")

        title = ""
        try:
            title = page.title()[:80]
        except Exception:
            title = page.url[:80] if page.url else ""
        if title:
            parts.append(title)
        try:
            parts.append((page.url or "")[:120])
        except Exception:
            pass
        if not ok:
            ok = shown > 0
        return ok, " | ".join(parts) if parts else (page.url or "已连接浏览器")
    except Exception as exc:
        return False, f"连接浏览器失败: {exc}"
    finally:
        try:
            playwright.stop()
        except Exception:
            pass


def reveal_problem_page(
    port: int | None,
    page: Page | None = None,
    *,
    status: AccountStatus | None = None,
) -> None:
    """查号遇到需人工看的状态时：把浏览器亮到当前问题页（不重新登录）。"""
    if port is None:
        return
    keep = status is None or status in _KEEP_BROWSER_OPEN_STATUSES
    if not keep:
        return
    try:
        if page is not None:
            try:
                page.bring_to_front()
            except Exception:
                pass
        ok, detail = focus_browser_for_debug(port, "")
        if ok:
            log_step(f"已显示问题页面（无需再手动登录）: {detail}")
        else:
            show_hubstudio_windows_for_debug(debug_port=port, log=log_step)
            log_step(f"尝试显示浏览器窗口: {detail}")
    except Exception as exc:
        log_step(f"显示问题页失败: {exc}")


def open_browser_for_debug(
    api_base: str,
    machine_id: str,
    login_url: str,
    api_key: str,
) -> tuple[bool, str, str]:
    """
    点「打开」：
    1) 用 force_visible 复用已开环境（不带登录 URL，避免冲掉问题页）
    2) HubStudio foreground + CDP 恢复窗口位置 + Win32 置顶
    只有环境已关掉时才新开（此时才不可避免落到登录页）
    """
    try:
        return _open_browser_for_debug_impl(api_base, machine_id, login_url, api_key)
    except Exception as exc:
        # 环境可能已弹出，勿把异常冒成非 JSON / 未捕获 500
        env_code = ""
        try:
            env = resolve_env(api_base, machine_id, api_key)
            env_code = env.container_code if env else ""
        except Exception:
            pass
        shown = 0
        try:
            shown = show_hubstudio_windows_for_debug(log=log_step)
        except Exception:
            pass
        if shown:
            return True, env_code, f"窗口已显示（过程有异常可忽略）: {exc}"
        return False, env_code, f"打开异常（若窗口已弹出可忽略）: {exc}"


def _open_browser_for_debug_impl(
    api_base: str,
    machine_id: str,
    login_url: str,
    api_key: str,
) -> tuple[bool, str, str]:
    env = resolve_env(api_base, machine_id, api_key)
    if env is None:
        return False, "", f"未找到机子号 {machine_id} 对应 HubStudio 环境"

    parts: list[str] = []

    # 0) 优先用已记住的调试端口（避免 start 接口卡住/无端口时「打开」无响应）
    remembered = get_remembered_debug_port(env.container_code)
    if remembered:
        try:
            ok0, detail0 = focus_browser_for_debug(remembered, "")
            if ok0:
                show_hubstudio_windows_for_debug(debug_port=remembered, log=log_step)
                bring_browser_to_foreground(api_base, env.container_code, api_key)
                return (
                    True,
                    env.container_code,
                    f"已用已记住端口 {remembered} 置顶 | {detail0}",
                )
            parts.append(f"记住端口 {remembered} 暂不可用: {detail0}")
        except Exception as exc:
            parts.append(f"记住端口连接失败: {exc}")

    # 1) 始终尝试可见复用（不带登录页）；静默/屏外窗口也能拿到端口再拉回前台
    port, err = start_browser(
        api_base,
        env.container_code,
        login_url,
        api_key,
        force_visible=True,
        open_login_tab=False,
    )
    if port is None:
        port = get_remembered_debug_port(env.container_code)
    if port is None:
        # 再试一次 start（部分 HubStudio 版本第一次 -10013 无端口）
        time.sleep(0.6)
        port, err2 = start_browser(
            api_base,
            env.container_code,
            login_url,
            api_key,
            force_visible=True,
            open_login_tab=False,
        )
        if err2 and not err:
            err = err2
    if port is None:
        running = browser_is_running(api_base, env.container_code, api_key)
        if running is False:
            port, err = start_browser(
                api_base,
                env.container_code,
                login_url,
                api_key,
                force_visible=True,
                open_login_tab=True,
            )
            if port is None:
                return False, env.container_code, err or "打开环境失败"
            parts.append("环境已关闭，已重新打开（需重新登录）")
        else:
            fg_ok, fg_detail = bring_browser_to_foreground(
                api_base, env.container_code, api_key
            )
            if fg_ok:
                parts.append(fg_detail)
            shown = show_hubstudio_windows_for_debug(log=log_step)
            if shown == 0:
                # 放宽：不限定 debug_port，再扫一次 Chrome 窗
                shown = show_hubstudio_windows_for_debug(debug_port=None, log=log_step)
            if fg_ok or shown:
                parts.append(f"已尝试置顶（未拿到调试端口） shown={shown}")
                return True, env.container_code, " | ".join(parts)
            return False, env.container_code, err or "无法获取调试端口，请确认环境仍在运行"
    else:
        parts.append("已复用运行中环境（未跳转登录页）")

    register_automation_browser(port)
    remember_debug_port(env.container_code, port)

    # 2) 官方置顶
    fg_ok, fg_detail = bring_browser_to_foreground(
        api_base, env.container_code, api_key
    )
    if fg_ok:
        parts.append(fg_detail)

    # 3) CDP 恢复窗口 bounds + Win32 把屏外/最小化窗口拉回（静默 -24000 位置）
    ok, detail = focus_browser_for_debug(port, "")
    if ok:
        parts.append(detail)
    else:
        parts.append(detail)
        # CDP 失败再试一次
        time.sleep(0.5)
        ok2, detail2 = focus_browser_for_debug(port, "")
        if ok2:
            ok = True
            parts.append(f"重试成功: {detail2}")

    shown = show_hubstudio_windows_for_debug(debug_port=port, log=log_step)
    if shown == 0:
        shown = show_hubstudio_windows_for_debug(debug_port=None, log=log_step)
    if shown:
        parts.append(f"Win32 已显示 {shown} 个浏览器窗口")

    if ok or shown or fg_ok:
        return True, env.container_code, " | ".join(parts)

    return False, env.container_code, " | ".join(parts) or "打开失败"


def open_only_one(
    api_base: str,
    api_key: str,
    login_url: str,
    env: EnvRef,
) -> CheckResult:
    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    port, err = ensure_browser_port(api_base, env.container_code, login_url, api_key)
    if port is None:
        return CheckResult(
            machine_id=env.machine_id,
            email="",
            status=AccountStatus.OPEN_FAILED,
            detail=err,
            container_code=env.container_code,
            checked_at=checked_at,
        )
    return CheckResult(
        machine_id=env.machine_id,
        email="",
        status=AccountStatus.OK,
        detail=(
            f"已后台静默打开环境，调试端口 {port}"
            + ("（无头模式）" if _effective_headless() else "（已最小化）")
        ),
        container_code=env.container_code,
        checked_at=checked_at,
    )


def check_one_account(
    api_base: str,
    api_key: str,
    login_url: str,
    env: EnvRef,
    email: str,
    password: str,
    wait_timeout_sec: int,
    manual_mfa_timeout_sec: int,
    close_browser: bool,
    recovery_email: str = "",
    recovery_password: str = "",
    recovery_config: dict | None = None,
) -> CheckResult:
    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result: CheckResult | None = None
    port, err = ensure_browser_port(api_base, env.container_code, login_url, api_key)
    if port is None:
        result = CheckResult(
            machine_id=env.machine_id,
            email=email,
            status=AccountStatus.OPEN_FAILED,
            detail=err,
            container_code=env.container_code,
            checked_at=checked_at,
        )
        maybe_stop_browser(api_base, env.container_code, api_key, result.status, close_browser)
        return result

    time.sleep(0.1 if _fast_fill_enabled() else 0.3)

    try:
        playwright, browser, page = acquire_login_page(port, login_url)
        try:
            status, login_detail = auto_login(page, email, password)
            if status != AccountStatus.OK:
                result = finalize_check_result(
                    machine_id=env.machine_id,
                    email=email,
                    status=status,
                    detail=login_detail,
                    container_code=env.container_code,
                    final_url=page.url,
                    checked_at=checked_at,
                    page=page,
                )
            else:
                # 登录过程可能打开新标签，重新定位登录页
                context = browser.contexts[0]
                page = get_active_flow_page(context, login_url, timeout_sec=8)
                status, detail = wait_for_final_status(
                    page,
                    wait_timeout_sec,
                    manual_mfa_timeout_sec,
                    recovery_email=recovery_email,
                    recovery_password=recovery_password,
                    recovery_config=recovery_config or {},
                    machine_id=env.machine_id,
                )
                try:
                    final_url = page.url
                except Exception:
                    final_url = ""

                used_recovery = (recovery_email or "").strip()
                # 仅回填已领取/已提供的；禁止仅凭状态再 claim（会打乱池顺序）
                if status in {
                    AccountStatus.WAIT_CODE,
                    AccountStatus.RECOVERY_BOUND,
                } and not used_recovery:
                    claimed_email, _ = resolve_recovery_credentials(
                        env.machine_id,
                        used_recovery,
                        recovery_password,
                        recovery_config or {},
                    )
                    if claimed_email:
                        used_recovery = claimed_email
                elif status == AccountStatus.RECOVERY_FAILED:
                    # 绑定失败时 wait_for_final 已归还；结果里不要再领一条
                    pass
                if status == AccountStatus.WAIT_CODE and used_recovery:
                    register_pending_recovery_code(
                        machine_id=env.machine_id,
                        container_code=env.container_code,
                        port=port,
                        recovery_email=used_recovery,
                        login_email=email,
                    )
                    remember_debug_port(env.container_code, port)

                result = finalize_check_result(
                    machine_id=env.machine_id,
                    email=email,
                    status=status,
                    detail=detail,
                    container_code=env.container_code,
                    final_url=final_url,
                    checked_at=checked_at,
                    page=page,
                    recovery_email=used_recovery,
                    awaiting_code=status == AccountStatus.WAIT_CODE,
                )
        finally:
            playwright.stop()
    except Exception as exc:
        msg = str(exc)
        if _is_transient_nav_error(exc) or "ERR_CONNECTION" in msg or "网卡打开登录页" in msg:
            result = CheckResult(
                machine_id=env.machine_id,
                email=email,
                status=AccountStatus.NETWORK_CARD,
                detail=f"应更换IP（打开登录页连接中断: {msg[:180]}）",
                container_code=env.container_code,
                checked_at=checked_at,
            )
        else:
            result = CheckResult(
                machine_id=env.machine_id,
                email=email,
                status=AccountStatus.UNKNOWN,
                detail=f"脚本异常: {exc}",
                container_code=env.container_code,
                checked_at=checked_at,
            )

    assert result is not None
    # 需人工处理的状态只保持后台打开，不自动弹窗；点「打开」再显示当前页
    maybe_stop_browser(api_base, env.container_code, api_key, result.status, close_browser)
    return result


def fill_only_one(
    api_base: str,
    api_key: str,
    login_url: str,
    env: EnvRef,
    email: str,
    password: str,
    recovery_email: str = "",
    recovery_password: str = "",
    recovery_config: dict | None = None,
    close_browser: bool = True,
) -> CheckResult:
    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result: CheckResult | None = None
    port, err = ensure_browser_port(api_base, env.container_code, login_url, api_key)
    if port is None:
        result = CheckResult(
            machine_id=env.machine_id,
            email=email,
            status=AccountStatus.OPEN_FAILED,
            detail=err,
            container_code=env.container_code,
            checked_at=checked_at,
        )
        maybe_stop_browser(api_base, env.container_code, api_key, result.status, close_browser)
        return result

    time.sleep(0.1 if _fast_fill_enabled() else 0.3)

    try:
        playwright, browser, page = acquire_login_page(port, login_url)
        try:
            status, login_detail = auto_login(page, email, password)

            # 密码页绝不能以「网卡」结案：重新走登录填密码
            if status == AccountStatus.NETWORK_CARD:
                try:
                    if (
                        is_password_entry_page(page)
                        or find_password_input(page) is not None
                    ):
                        log_step("当前为密码输入页，忽略网卡误判并重新提交密码")
                        status, login_detail = auto_login(page, email, password)
                except Exception:
                    pass

            if status in {
                AccountStatus.ACCOUNT_NOT_FOUND,
                AccountStatus.BAD_PASSWORD,
                AccountStatus.LOCKED,
                AccountStatus.BAD_CREDENTIALS,
                AccountStatus.ALREADY_BOUND,
                AccountStatus.NETWORK_CARD,
            }:
                result = finalize_check_result(
                    machine_id=env.machine_id,
                    email=email,
                    status=status,
                    detail=login_detail,
                    container_code=env.container_code,
                    final_url=page.url,
                    checked_at=checked_at,
                    page=page,
                )
            elif status in {
                AccountStatus.OK,
                AccountStatus.LOGIN_OK,
                AccountStatus.AMZ_BANNED,
            }:
                # 已进邮箱：再扫封号（防漏判成「登入」）
                if status != AccountStatus.AMZ_BANNED:
                    status, login_detail = evaluate_inbox_account_status(page)
                result = finalize_check_result(
                    machine_id=env.machine_id,
                    email=email,
                    status=status,
                    detail=login_detail,
                    container_code=env.container_code,
                    final_url=page.url,
                    checked_at=checked_at,
                    page=page,
                )
            else:
                # 可能需要绑定辅助邮箱：先等页面加载完成/绑定页出现，再按需从池领取
                context = browser.contexts[0]
                page = get_active_flow_page(context, login_url, timeout_sec=10)
                used_recovery = (recovery_email or "").strip()
                used_recovery_password = recovery_password or ""

                settle_sec = 55.0 if _keep_background() else 48.0
                status, login_detail, page = wait_for_auth_page_settled(
                    page,
                    context,
                    login_url,
                    timeout_sec=settle_sec,
                )

                # 超时结果若页面已是绑定页：改判并继续领辅助邮箱，不要直接结案
                if status == AccountStatus.TIMEOUT and _page_looks_like_recovery_bind(page):
                    log_step("超时瞬间已到绑定页，改判为需要绑定辅助邮箱")
                    status = AccountStatus.NEED_RECOVERY
                    login_detail = "需要绑定辅助邮箱"

                if status in _NON_RECOVERY_FLOW_STATUSES:
                    result = finalize_check_result(
                        machine_id=env.machine_id,
                        email=email,
                        status=status,
                        detail=login_detail,
                        container_code=env.container_code,
                        final_url=page.url,
                        checked_at=checked_at,
                        page=page,
                    )
                else:
                    # 电话认证：绝不领辅助邮箱
                    try:
                        if is_phone_verify_page(page):
                            status = AccountStatus.NEED_PHONE
                            login_detail = (
                                "需要电话认证（确认手机号末位，非辅助邮箱流程）"
                            )
                            used_recovery = ""
                            used_recovery_password = ""
                    except Exception:
                        pass

                    # 必须真到绑定/验证码页才领取；仅凭 NEED_RECOVERY 状态不够
                    on_recovery = (
                        status != AccountStatus.NEED_PHONE
                        and (
                            is_definitely_recovery_flow_page(page)
                            or _page_looks_like_recovery_bind(page)
                        )
                    )
                    if on_recovery:
                        if not used_recovery:
                            used_recovery, used_recovery_password = resolve_recovery_credentials(
                                env.machine_id,
                                used_recovery,
                                used_recovery_password,
                                recovery_config or {},
                            )
                        if not used_recovery:
                            status = AccountStatus.NEED_RECOVERY
                            login_detail = (
                                "需要绑定辅助邮箱，但辅助邮箱池已空（请在左侧补充）"
                            )
                        else:
                            log_step(f"绑定辅助邮箱: {used_recovery}")
                            bind_result = try_bind_recovery_on_page(
                                page,
                                used_recovery,
                                used_recovery_password,
                                (recovery_config or {}).get("recovery_imap") or {},
                                recovery_config or {},
                                log=log_step,
                            )
                            manual = bool(
                                (recovery_config or {}).get("recovery_manual_code", True)
                            )
                            if bind_result is not None:
                                ok, bind_detail = bind_result
                                if ok and manual:
                                    status = AccountStatus.WAIT_CODE
                                    login_detail = bind_detail
                                    register_pending_recovery_code(
                                        machine_id=env.machine_id,
                                        container_code=env.container_code,
                                        port=port,
                                        recovery_email=used_recovery,
                                        login_email=email,
                                    )
                                    remember_debug_port(env.container_code, port)
                                elif ok:
                                    status = AccountStatus.RECOVERY_BOUND
                                    login_detail = bind_detail
                                else:
                                    release_recovery_if_needed(
                                        env.machine_id,
                                        recovery_config or {},
                                        reason="绑定失败归还",
                                    )
                                    used_recovery = ""
                                    used_recovery_password = ""
                                    if "电话认证" in (bind_detail or ""):
                                        status = AccountStatus.NEED_PHONE
                                        login_detail = bind_detail
                                    else:
                                        status = AccountStatus.RECOVERY_FAILED
                                        login_detail = bind_detail
                            else:
                                status = AccountStatus.RECOVERY_FAILED
                                login_detail = "绑定页识别失败，未能执行填表"
                                release_recovery_if_needed(
                                    env.machine_id,
                                    recovery_config or {},
                                    reason="绑定页识别失败归还",
                                )
                                used_recovery = ""
                                used_recovery_password = ""
                    elif status == AccountStatus.NEED_PHONE:
                        used_recovery = ""
                        used_recovery_password = ""
                    elif status == AccountStatus.NEED_RECOVERY:
                        if is_definitely_recovery_flow_page(page):
                            # 绑定页已出现但未走进 on_recovery 分支时的兜底
                            if not used_recovery:
                                used_recovery, used_recovery_password = resolve_recovery_credentials(
                                    env.machine_id,
                                    used_recovery,
                                    used_recovery_password,
                                    recovery_config or {},
                                )
                            if used_recovery:
                                log_step(f"绑定辅助邮箱: {used_recovery}")
                                bind_result = try_bind_recovery_on_page(
                                    page,
                                    used_recovery,
                                    used_recovery_password,
                                    (recovery_config or {}).get("recovery_imap") or {},
                                    recovery_config or {},
                                    log=log_step,
                                )
                                if bind_result is not None:
                                    ok, bind_detail = bind_result
                                    manual = bool(
                                        (recovery_config or {}).get("recovery_manual_code", True)
                                    )
                                    if ok and manual:
                                        status = AccountStatus.WAIT_CODE
                                        login_detail = bind_detail
                                        register_pending_recovery_code(
                                            machine_id=env.machine_id,
                                            container_code=env.container_code,
                                            port=port,
                                            recovery_email=used_recovery,
                                            login_email=email,
                                        )
                                        remember_debug_port(env.container_code, port)
                                    elif ok:
                                        status = AccountStatus.RECOVERY_BOUND
                                        login_detail = bind_detail
                                    else:
                                        release_recovery_if_needed(
                                            env.machine_id,
                                            recovery_config or {},
                                            reason="绑定失败归还",
                                        )
                                        used_recovery = ""
                                        used_recovery_password = ""
                                        status = AccountStatus.RECOVERY_FAILED
                                        login_detail = bind_detail
                            else:
                                login_detail = (
                                    "需要绑定辅助邮箱，但辅助邮箱池已空（请在左侧补充）"
                                )
                        elif is_microsoft_login_page_loading(page) or is_ms_auth_page_shell(page):
                            login_detail = (
                                (login_detail or "需要绑定辅助邮箱")
                                + "（页面仍在加载，请点打开查看）"
                            )
                            used_recovery = ""
                            used_recovery_password = ""
                        else:
                            login_detail = (
                                (login_detail or "需要绑定辅助邮箱")
                                + "（绑定页未出现，未领取辅助邮箱）"
                            )
                            used_recovery = ""
                            used_recovery_password = ""
                    elif status == AccountStatus.OK:
                        if is_outlook_inbox(page):
                            status, login_detail = evaluate_inbox_account_status(page)
                        else:
                            login_detail = login_detail or "登录成功，未出现辅助邮箱绑定页"
                    else:
                        # 未进入绑定页：不领取辅助邮箱
                        login_detail = login_detail or f"当前状态: {status.value}"
                        used_recovery = used_recovery if status in {
                            AccountStatus.WAIT_CODE,
                            AccountStatus.RECOVERY_BOUND,
                        } else ""

                    result = finalize_check_result(
                        machine_id=env.machine_id,
                        email=email,
                        status=status,
                        detail=login_detail,
                        container_code=env.container_code,
                        final_url=page.url,
                        checked_at=checked_at,
                        page=page,
                        recovery_email=used_recovery,
                        awaiting_code=status == AccountStatus.WAIT_CODE,
                    )
        finally:
            playwright.stop()
    except Exception as exc:
        msg = str(exc)
        if _is_transient_nav_error(exc) or "ERR_CONNECTION" in msg or "网卡打开登录页" in msg:
            result = CheckResult(
                machine_id=env.machine_id,
                email=email,
                status=AccountStatus.NETWORK_CARD,
                detail=f"应更换IP（打开登录页连接中断: {msg[:180]}）",
                container_code=env.container_code,
                checked_at=checked_at,
            )
        else:
            result = CheckResult(
                machine_id=env.machine_id,
                email=email,
                status=AccountStatus.UNKNOWN,
                detail=f"脚本异常: {exc}",
                container_code=env.container_code,
                checked_at=checked_at,
            )

    assert result is not None
    # 需要绑定辅助邮箱 / 异常状态：只保持后台打开，不自动弹窗；靠手动点「打开」
    maybe_stop_browser(api_base, env.container_code, api_key, result.status, close_browser)
    return result


def save_results(results: list[CheckResult], output_path: Path) -> None:
    fieldnames = [
        "machine_id",
        "email",
        "status",
        "detail",
        "container_code",
        "final_url",
        "checked_at",
    ]
    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for item in results:
            row = asdict(item)
            row["status"] = item.status.value
            writer.writerow(row)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="HubStudio Outlook 可用性检测")
    parser.add_argument(
        "--open-only",
        action="store_true",
        help="只打开环境并进入登录页，不自动填账号密码",
    )
    parser.add_argument(
        "--fill-only",
        action="store_true",
        help="打开环境并自动填账号密码，不等待登录结果",
    )
    parser.add_argument(
        "--accounts",
        default="accounts.csv",
        help="账号 CSV 路径（默认 accounts.csv）",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_dir = data_dir()
    config_path = base_dir / "config.json"
    accounts_path = Path(args.accounts)
    if not accounts_path.is_absolute():
        accounts_path = base_dir / accounts_path

    if not config_path.exists():
        print("缺少 config.json")
        return 1
    if not accounts_path.exists():
        print(f"缺少账号文件: {accounts_path}")
        return 1

    config = load_json(config_path)
    accounts = load_accounts(accounts_path)
    if not accounts:
        print("accounts.csv 中没有有效机子号")
        return 1

    api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = config.get("local_api_key", "")
    login_url = config["login_url"]
    wait_timeout_sec = int(config.get("wait_timeout_sec", 45))
    manual_mfa_timeout_sec = int(config.get("manual_mfa_timeout_sec", 300))
    close_browser = bool(config.get("close_browser_after_check", True))

    results: list[CheckResult] = []
    total = len(accounts)
    mode = "仅打开环境" if args.open_only else ("仅填表" if args.fill_only else "自动登录检测")
    print(f"模式: {mode} | 共 {total} 条 | API: {api_base}\n")

    for idx, acc in enumerate(accounts, start=1):
        machine_id = acc["machine_id"]
        email = acc["email"]
        print(f"[{idx}/{total}] 机子号: {machine_id}" + (f" | 邮箱: {email}" if email else ""))

        env = resolve_env(api_base, machine_id, api_key)
        if env is None:
            result = CheckResult(
                machine_id=machine_id,
                email=email,
                status=AccountStatus.RESOLVE_FAILED,
                detail="未在 HubStudio 找到该环境名称/备注/序号",
                checked_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
            results.append(result)
            print(f"    => {result.status.value} | {result.detail}\n")
            continue

        print(
            f"    已解析: containerCode={env.container_code}, "
            f"serialNumber={env.serial_number}, name={env.container_name}"
        )

        if args.open_only:
            result = open_only_one(api_base, api_key, login_url, env)
        elif args.fill_only:
            if not email or not acc["password"]:
                result = CheckResult(
                    machine_id=machine_id,
                    email=email,
                    status=AccountStatus.UNKNOWN,
                    detail="请先在 accounts.csv 填入邮箱和密码",
                    container_code=env.container_code,
                    checked_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                )
            else:
                result = fill_only_one(
                    api_base, api_key, login_url, env, email, acc["password"]
                )
        else:
            if not email or not acc["password"] or "请替换" in email or "请替换" in acc["password"]:
                result = CheckResult(
                    machine_id=machine_id,
                    email=email,
                    status=AccountStatus.UNKNOWN,
                    detail="请先在 accounts.csv 填入真实邮箱和密码；或改用 --open-only",
                    container_code=env.container_code,
                    checked_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                )
            else:
                result = check_one_account(
                    api_base=api_base,
                    api_key=api_key,
                    login_url=login_url,
                    env=env,
                    email=email,
                    password=acc["password"],
                    wait_timeout_sec=wait_timeout_sec,
                    manual_mfa_timeout_sec=manual_mfa_timeout_sec,
                    close_browser=close_browser,
                )

        results.append(result)
        print(f"    => {result.status.value} | {result.detail}\n")

    output_path = base_dir / f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    save_results(results, output_path)

    ok_count = sum(
        1
        for r in results
        if r.status in {AccountStatus.OK, AccountStatus.LOGIN_OK}
    )
    print("=" * 50)
    print(f"完成: 成功/可用 {ok_count}/{total}")
    print(f"结果已保存: {output_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n用户中断")
        raise SystemExit(130)
