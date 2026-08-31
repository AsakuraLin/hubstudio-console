# -*- coding: utf-8 -*-
"""Foxmail 桌面客户端：自动添加「其它邮箱」账号。"""

from __future__ import annotations

import re
import subprocess
import threading
import time
from pathlib import Path
from typing import Any, Callable

LogFn = Callable[[str], None]

_foxmail_ui_lock = threading.Lock()
_parked_foxmail_rects: dict[int, tuple[int, int, int, int]] = {}
_FOXMAIL_OFFSCREEN_X = -2600
_FOXMAIL_OFFSCREEN_Y = 40
_foxmail_cancel = threading.Event()


def request_foxmail_cancel() -> None:
    """用户点停止时调用：打断 Foxmail/等待，并拉回桌面。"""
    _foxmail_cancel.set()
    try:
        _restore_foxmail_to_desktop(log=_noop_log, then_minimize=False)
    except Exception:
        pass


def clear_foxmail_cancel() -> None:
    _foxmail_cancel.clear()


def is_foxmail_cancel_requested() -> bool:
    return _foxmail_cancel.is_set()

try:
    import win32api
    import win32con
    import win32gui
    import win32process

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

_minimized_windows: list[int] = []
_window_lock = threading.Lock()
_tracked_debug_ports: set[int] = set()
_tracked_browser_pids: set[int] = set()
_PROCESS_QUERY_LIMITED = 0x1000


def _pid_for_listening_port(port: int) -> int | None:
    try:
        out = subprocess.check_output(
            ["netstat", "-ano"],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
            text=True,
            errors="ignore",
        )
        token = f":{int(port)}"
        for line in out.splitlines():
            upper = line.upper()
            if "LISTENING" not in upper or token not in line:
                continue
            parts = line.split()
            if parts and parts[-1].isdigit():
                return int(parts[-1])
    except Exception:
        pass
    return None


def register_automation_browser(debug_port: int | None = None) -> None:
    """登记 HubStudio 调试端口，后续只最小化/恢复对应浏览器窗口。"""
    if not debug_port or not HAS_WIN32:
        return
    port = int(debug_port)
    with _window_lock:
        _tracked_debug_ports.add(port)
        pid = _pid_for_listening_port(port)
        if pid:
            _tracked_browser_pids.add(pid)


def _is_hubstudio_client_shell(title: str) -> bool:
    """HubStudio 客户端/任务栏壳窗口，不是浏览器内容页。"""
    t = (title or "").strip()
    if not t:
        return False
    if t.lower() in {"hubstudio", "hub studio"}:
        return True
    if t.isdigit() and len(t) <= 8:
        return True
    if re.match(r"^HubStudio[\s\-–—]*\d+$", t, re.I):
        return True
    return False


_PAGE_TITLE_MARKERS = (
    "microsoft",
    "outlook",
    "sign in",
    "登录",
    "パスワード",
    "live.com",
    "microsoftonline",
    "mail",
    "邮箱",
    "account",
)


def _collect_browser_pids(debug_port: int | None = None) -> set[int]:
    pids: set[int] = set()
    if not debug_port or not HAS_WIN32:
        return pids
    root = _pid_for_listening_port(int(debug_port))
    if not root:
        return pids
    pids.add(root)
    try:
        out = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"(Get-CimInstance Win32_Process -Filter 'ParentProcessId={root}').ProcessId",
            ],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
            text=True,
            errors="ignore",
        )
        for line in out.splitlines():
            line = line.strip()
            if line.isdigit():
                pids.add(int(line))
    except Exception:
        pass
    return pids


def _is_browser_content_window(hwnd: int, target_pids: set[int] | None = None) -> bool:
    """真实浏览器内容窗口（Chrome 页面），排除 HubStudio 客户端壳。"""
    if not HAS_WIN32 or not win32gui.IsWindow(hwnd):
        return False
    title = win32gui.GetWindowText(hwnd) or ""
    if "Foxmail" in title or _is_hubstudio_client_shell(title):
        return False
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        return False
    cls = win32gui.GetClassName(hwnd)
    title_l = title.lower()
    if cls in {"Chrome_WidgetWin_1", "Chrome_WidgetWin_0"}:
        if target_pids and pid in target_pids:
            return True
        if any(m in title_l for m in _PAGE_TITLE_MARKERS):
            return True
        if title and "hubstudio" not in title_l and len(title) > 8:
            return True
    with _window_lock:
        if pid in _tracked_browser_pids and cls in {
            "Chrome_WidgetWin_1",
            "Chrome_WidgetWin_0",
        }:
            return True
    exe = _get_process_exe(pid).lower()
    if "hubstudio" in exe.replace("\\", "/") and cls in {
        "Chrome_WidgetWin_1",
        "Chrome_WidgetWin_0",
    }:
        with _window_lock:
            _tracked_browser_pids.add(pid)
        return True
    return False


def _get_process_exe(pid: int) -> str:
    if not HAS_WIN32 or pid <= 0:
        return ""
    try:
        handle = win32api.OpenProcess(_PROCESS_QUERY_LIMITED, False, pid)
        try:
            return win32process.GetModuleFileNameEx(handle, 0) or ""
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        return ""


def _is_automation_browser_window(hwnd: int) -> bool:
    """仅 HubStudio 自动化浏览器，不误伤用户其它 Chrome/Outlook 窗口。"""
    if not HAS_WIN32 or not win32gui.IsWindow(hwnd):
        return False
    title = win32gui.GetWindowText(hwnd) or ""
    if "Foxmail" in title:
        return False
    if _is_hubstudio_client_shell(title):
        return False
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        return False
    with _window_lock:
        if pid in _tracked_browser_pids:
            return True
    exe = _get_process_exe(pid).lower()
    if "hubstudio" in exe.replace("\\", "/"):
        with _window_lock:
            _tracked_browser_pids.add(pid)
        return True
    cls = win32gui.GetClassName(hwnd)
    if cls in {"Chrome_WidgetWin_1", "Chrome_WidgetWin_0"}:
        with _window_lock:
            if pid in _tracked_browser_pids:
                return True
        return False
    return False

def _noop_log(_: str) -> None:
    pass


def _foxmail_cfg(config: dict[str, Any]) -> dict[str, Any]:
    return config.get("foxmail") or {}


def _enum_foxmail_windows(
    *,
    visible_only: bool = False,
) -> list[tuple[int, str, str, tuple[int, int, int, int]]]:
    """枚举 Foxmail 顶层窗口（可含不可见，便于托盘/隐藏态后台操作）。"""
    rows: list[tuple[int, str, str, tuple[int, int, int, int]]] = []

    def cb(hwnd, _):
        try:
            if visible_only and not win32gui.IsWindowVisible(hwnd):
                return True
            if not _is_foxmail_process(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            cls = win32gui.GetClassName(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            rows.append((hwnd, cls, title, rect))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return rows


def _enum_visible_windows(
    *,
    include_offscreen: bool = True,
) -> list[tuple[int, str, str, tuple[int, int, int, int]]]:
    rows: list[tuple[int, str, str, tuple[int, int, int, int]]] = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        if rect[0] <= -30000:
            return True
        if not include_offscreen and rect[0] < -100:
            return True
        rows.append((hwnd, cls, title, rect))
        return True

    win32gui.EnumWindows(cb, None)
    return rows


def _park_foxmail_offscreen(log: LogFn = _noop_log) -> int:
    """临时把 Foxmail 移到屏外（仅自动化期间）；结束后必须 restore。"""
    global _parked_foxmail_rects
    if not HAS_WIN32:
        return 0
    count = 0
    flags = (
        win32con.SWP_NOSIZE
        | win32con.SWP_NOZORDER
        | win32con.SWP_NOACTIVATE
        | getattr(win32con, "SWP_NOSENDCHANGING", 0)
    )
    for hwnd, cls, _title, rect in _enum_foxmail_windows(visible_only=False):
        if not any(
            k in cls
            for k in (
                "TFoxMainFrm",
                "TOptionForm",
                "TAccCreate",
                "TFoxCompose",
                "OrayUI",
            )
        ):
            continue
        left, top, right, bottom = rect
        w, h = right - left, bottom - top
        if w <= 2 or h <= 2:
            continue
        try:
            if win32gui.IsIconic(hwnd) or not win32gui.IsWindowVisible(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
                time.sleep(0.05)
                rect = win32gui.GetWindowRect(hwnd)
                left, top, right, bottom = rect
                w, h = right - left, bottom - top
            if left <= _FOXMAIL_OFFSCREEN_X + 50 and win32gui.IsWindowVisible(hwnd):
                continue
            if hwnd not in _parked_foxmail_rects and left > -100:
                _parked_foxmail_rects[hwnd] = (left, top, right, bottom)
            win32gui.SetWindowPos(
                hwnd,
                0,
                _FOXMAIL_OFFSCREEN_X + (count % 3) * 40,
                _FOXMAIL_OFFSCREEN_Y + (count % 5) * 30,
                0,
                0,
                flags,
            )
            count += 1
        except Exception:
            pass
    if count:
        log(f"已将 {count} 个 Foxmail 窗口临时移到屏外")
    return count


def _restore_foxmail_to_desktop(log: LogFn = _noop_log, *, then_minimize: bool = False) -> int:
    """把屏外/隐藏的 Foxmail 拉回桌面，点击任务栏即可正常显示。"""
    global _parked_foxmail_rects
    if not HAS_WIN32:
        return 0
    count = 0
    screen_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    flags = (
        win32con.SWP_NOZORDER
        | win32con.SWP_NOACTIVATE
        | getattr(win32con, "SWP_SHOWWINDOW", 0x0040)
    )

    targets: list[tuple[int, str, tuple[int, int, int, int]]] = []
    for hwnd, cls, _title, rect in _enum_foxmail_windows(visible_only=False):
        if not any(
            k in cls
            for k in (
                "TFoxMainFrm",
                "TOptionForm",
                "TAccCreate",
                "TFoxCompose",
                "OrayUI",
            )
        ):
            continue
        left, top, right, bottom = rect
        w, h = max(1, right - left), max(1, bottom - top)
        if w <= 2 or h <= 2:
            continue
        targets.append((hwnd, cls, (left, top, right, bottom)))

    for hwnd, cls, rect in targets:
        try:
            left, top, right, bottom = rect
            w, h = max(400, right - left), max(300, bottom - top)
            saved = _parked_foxmail_rects.get(hwnd)
            if saved:
                sl, st, sr, sb = saved
                w, h = max(400, sr - sl), max(300, sb - st)
                x, y = sl, st
            else:
                x, y = 60, 40
            # 保证在可见屏内
            x = max(0, min(x, screen_w - 200))
            y = max(0, min(y, screen_h - 120))
            w = min(w, screen_w - 40)
            h = min(h, screen_h - 60)

            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
            win32gui.SetWindowPos(hwnd, 0, x, y, w, h, flags)
            if then_minimize and "TFoxMainFrm" in cls:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            count += 1
        except Exception:
            pass

    _parked_foxmail_rects.clear()
    if count:
        if then_minimize:
            log(f"已恢复 {count} 个 Foxmail 窗口到桌面并最小化（点任务栏可显示）")
        else:
            log(f"已恢复 {count} 个 Foxmail 窗口到桌面")
    return count


def _keep_foxmail_background(log: LogFn = _noop_log) -> None:
    """不再把窗口藏到屏外（否则任务栏点不开）。默认不抢前台即可。"""
    return


def _should_park_offscreen(config: dict[str, Any] | None = None) -> bool:
    fx = _foxmail_cfg(config or {})
    return bool(fx.get("park_offscreen", False))


def _title_match(title: str, keywords: list[str]) -> bool:
    t = title or ""
    return any(k in t for k in keywords)


def _is_foxmail_process(hwnd: int) -> bool:
    if not HAS_WIN32:
        return False
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        return False
    exe = _get_process_exe(pid).lower().replace("\\", "/")
    return "foxmail" in exe


def find_foxmail_main_window() -> tuple[int | None, tuple[int, int, int, int] | None]:
    """主窗口标题常是当前邮箱地址或空；含不可见/屏外窗口。"""
    if not HAS_WIN32:
        return None, None
    best = None
    for hwnd, cls, title, rect in _enum_foxmail_windows(visible_only=False):
        if any(k in (title or "") for k in ("系统设置", "新建帐号", "新建账号", "接收服务器", "请输入")):
            continue
        if "TOptionForm" in cls or "TAccCreate" in cls:
            continue
        if "TFoxMainFrm" not in cls and cls not in {"TApplication", "OrayUI", "TXGuiFoundation"}:
            # 仍允许大尺寸主窗
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            if w < 500 or h < 400:
                continue
        w = rect[2] - rect[0]
        h = rect[3] - rect[1]
        if w < 200 or h < 150:
            continue
        area = max(1, w * h)
        score = area
        if "TFoxMainFrm" in cls:
            score += 5_000_000
        if best is None or score > best[0]:
            best = (score, hwnd, rect)
    if best:
        return best[1], best[2]
    return None, None


def find_window(
    keywords: list[str],
    class_names: set[str] | None = None,
    *,
    foxmail_only: bool = False,
) -> tuple[int | None, tuple[int, int, int, int] | None]:
    if not HAS_WIN32:
        return None, None
    best = None
    source = (
        _enum_foxmail_windows(visible_only=False)
        if foxmail_only
        else _enum_visible_windows(include_offscreen=True)
    )
    for hwnd, cls, title, rect in source:
        if class_names and cls not in class_names:
            continue
        if foxmail_only and not _is_foxmail_process(hwnd):
            continue
        if _title_match(title, keywords):
            area = max(1, (rect[2] - rect[0]) * (rect[3] - rect[1]))
            if best is None or area > best[0]:
                best = (area, hwnd, rect)
    if best:
        return best[1], best[2]
    return None, None


_MGMT_KEYWORDS = ["系统设置", "帐号管理", "账号管理"]
_NEW_ACCOUNT_KEYWORDS = ["新建帐号", "新建账号"]
_FORM_KEYWORDS = [
    "接收服务器类型",
    "请输入账号密码",
    "请输入帐号密码",
    "账号密码",
    "帐号密码",
    "邮件帐号",
    "邮件账号",
    "E-mail",
]


def _walk_children(root: int, visitor) -> None:
    """深度优先遍历子控件。"""

    def cb(hwnd, _):
        try:
            visitor(hwnd)
            _walk_children(hwnd, visitor)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(root, cb, None)
    except Exception:
        pass


def _find_controls(
    root: int,
    *,
    text: str | None = None,
    text_contains: str | None = None,
    class_contains: str | None = None,
    visible_only: bool = True,
) -> list[int]:
    found: list[int] = []
    seen: set[int] = set()

    def visit(hwnd: int) -> None:
        if hwnd in seen:
            return
        try:
            if visible_only and not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd) or ""
            cls = win32gui.GetClassName(hwnd)
        except Exception:
            return
        if text is not None and title != text:
            return
        if text_contains is not None and text_contains not in title:
            return
        if class_contains is not None and class_contains not in cls:
            return
        seen.add(hwnd)
        found.append(hwnd)

    if root:
        _walk_children(root, visit)
    return found


def _force_activate_no_mouse(hwnd: int) -> None:
    """短暂把窗口设为前台（不移动鼠标）。用 Alt 解锁突破 Windows 前台限制。"""
    if not hwnd or not HAS_WIN32:
        return
    try:
        import ctypes

        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.05)
        # 允许跨进程切前台
        try:
            ctypes.windll.user32.AllowSetForegroundWindow(-1)
        except Exception:
            pass
        # 模拟一下 Alt（不移动鼠标），否则 SetForegroundWindow 常被系统忽略
        try:
            VK_MENU = 0x12
            KEYEVENTF_KEYUP = 0x0002
            ctypes.windll.user32.keybd_event(VK_MENU, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
        except Exception:
            pass

        fg = win32gui.GetForegroundWindow()
        cur_tid = win32api.GetCurrentThreadId()
        fg_tid = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
        tgt_tid = win32process.GetWindowThreadProcessId(hwnd)[0]
        attached_fg = False
        attached_cur = False
        try:
            if fg_tid and fg_tid != tgt_tid:
                attached_fg = bool(win32process.AttachThreadInput(fg_tid, tgt_tid, True))
            if cur_tid != tgt_tid:
                attached_cur = bool(win32process.AttachThreadInput(cur_tid, tgt_tid, True))
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.BringWindowToTop(hwnd)
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            try:
                win32gui.SetActiveWindow(hwnd)
            except Exception:
                pass
            try:
                win32gui.SetFocus(hwnd)
            except Exception:
                pass
        finally:
            if attached_cur:
                win32process.AttachThreadInput(cur_tid, tgt_tid, False)
            if attached_fg:
                win32process.AttachThreadInput(fg_tid, tgt_tid, False)
    except Exception:
        pass


def _dialog_text_blob(dlg: int) -> str:
    """收集对话框可见文本，用于识别企业微信推广页等。"""
    parts: list[str] = []
    try:
        seen: set[int] = set()

        def walk(h: int) -> None:
            if h in seen:
                return
            seen.add(h)
            try:
                if win32gui.IsWindowVisible(h):
                    t = win32gui.GetWindowText(h) or ""
                    if t.strip():
                        parts.append(t)
                child = win32gui.GetWindow(h, win32con.GW_CHILD)
                while child:
                    walk(child)
                    child = win32gui.GetWindow(child, win32con.GW_HWNDNEXT)
            except Exception:
                return

        walk(dlg)
    except Exception:
        pass
    return "\n".join(parts)


def _is_wecom_promo_page(dlg: int) -> bool:
    """新建帐号里的「企业微信/腾讯企业邮升级」推广页（不是邮箱类型列表）。"""
    if not dlg:
        return False
    blob = _dialog_text_blob(dlg)
    keys = (
        "企业微信",
        "下载企业微信",
        "腾讯企业邮已升级",
        "企业微信是腾讯企业邮",
        "点击此处添加",
    )
    return any(k in blob for k in keys)


def _ensure_provider_list_page(log: LogFn = _noop_log) -> int | None:
    """
    确保当前是邮箱类型列表（含「其它邮箱」）。
    若停在企业微信推广页/错误类型页，点取消退回或重开。
    """
    end = time.time() + 20.0
    while time.time() < end:
        if is_foxmail_cancel_requested():
            return None
        dlg, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        if not dlg:
            return None
        if _is_password_entry_page(dlg):
            return dlg
        if _page_visible(dlg, "startpage"):
            trees = _find_controls(dlg, class_contains="TVirtualDrawTree", visible_only=True)
            if trees:
                return dlg
        # 企业微信推广页：必须退回，不能当列表用
        if _is_wecom_promo_page(dlg):
            log("检测到「企业微信」推广页（非其它邮箱列表），正在退回…")
            with _brief_activate(dlg, restore=False):
                # 优先点取消（通常退回列表）；不要点「下载企业微信」
                for btn in _find_controls(dlg, text="取消", class_contains="TFMXButton"):
                    try:
                        bl, bt, br, bb = win32gui.GetWindowRect(btn)
                        _bg_click_at(dlg, (bl + br) / 2, (bt + bb) / 2)
                        _bg_click_hwnd(btn)
                    except Exception:
                        pass
                    time.sleep(0.5)
                    break
                else:
                    left, top, right, bottom = win32gui.GetWindowRect(dlg)
                    _bg_click_at(dlg, right - 18, top + 12)
                    _bg_send_key(dlg, win32con.VK_ESCAPE, times=2)
            time.sleep(0.45)
            dlg2, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
            if dlg2 and (
                _page_visible(dlg2, "startpage")
                or _find_controls(dlg2, class_contains="TVirtualDrawTree", visible_only=True)
            ):
                log("已从企业微信页退回邮箱类型列表")
                _keep_foxmail_dialog_active(dlg2)
                return dlg2
            # 整窗关了：由上层重开
            if not dlg2:
                log("企业微信页已关闭，需重新点新建")
                return None
            continue
        # 其它非列表页：尝试取消退回
        with _brief_activate(dlg, restore=False):
            for btn in _find_controls(dlg, text="取消", class_contains="TFMXButton"):
                try:
                    bl, bt, br, bb = win32gui.GetWindowRect(btn)
                    _bg_click_at(dlg, (bl + br) / 2, (bt + bb) / 2)
                except Exception:
                    _bg_click_hwnd(btn)
                time.sleep(0.4)
                break
        time.sleep(0.35)
    return None


def _brief_activate(hwnd: int, *, restore: bool = True):
    """
    激活目标窗口执行操作。
    restore=False：操作后不把前台还给原窗口（Foxmail FMX 列表必须保持激活，
    否则「新建」弹出后立刻失焦，第一次点击全部无效——正是卡第一次的原因）。
    不移动真实鼠标。
    """
    from contextlib import contextmanager

    @contextmanager
    def _cm():
        prev = None
        try:
            prev = win32gui.GetForegroundWindow()
        except Exception:
            prev = None
        if hwnd:
            _force_activate_no_mouse(hwnd)
            time.sleep(0.08)
        try:
            yield
        finally:
            if restore and prev and prev != hwnd:
                try:
                    if win32gui.IsWindow(prev):
                        _force_activate_no_mouse(prev)
                except Exception:
                    pass

    return _cm()


def _keep_foxmail_dialog_active(dlg: int) -> None:
    """保持新建帐号对话框可接收点击（不移动鼠标）。"""
    if not dlg:
        return
    _force_activate_no_mouse(dlg)
    try:
        win32gui.BringWindowToTop(dlg)
    except Exception:
        pass


def _bg_click_at(hwnd: int, screen_x: float, screen_y: float) -> None:
    """向控件发送点击消息（不移动真实鼠标）。优先点到坐标下的子控件。"""
    if not HAS_WIN32:
        return
    try:
        sx, sy = int(screen_x), int(screen_y)
        target = hwnd
        try:
            cx, cy = win32gui.ScreenToClient(hwnd, (sx, sy))
            child = win32gui.ChildWindowFromPoint(hwnd, (cx, cy))
            if child:
                target = child
        except Exception:
            pass
        cx, cy = win32gui.ScreenToClient(target, (sx, sy))
        lp = win32api.MAKELONG(max(0, cx), max(0, cy))
        win32gui.SendMessage(target, win32con.WM_MOUSEMOVE, 0, lp)
        win32gui.SendMessage(target, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
        time.sleep(0.04)
        win32gui.SendMessage(target, win32con.WM_LBUTTONUP, 0, lp)
        win32gui.PostMessage(target, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
        win32gui.PostMessage(target, win32con.WM_LBUTTONUP, 0, lp)
        # 同时给父对话框一份，避免 FMX 只吃父窗消息
        if target != hwnd:
            cx2, cy2 = win32gui.ScreenToClient(hwnd, (sx, sy))
            lp2 = win32api.MAKELONG(max(0, cx2), max(0, cy2))
            win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp2)
            win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp2)
    except Exception:
        pass


def _bg_dblclick_at(hwnd: int, screen_x: float, screen_y: float) -> None:
    """后台双击（不移动真实鼠标）。列表项常需双击才进入下一页。"""
    if not HAS_WIN32:
        return
    try:
        cx, cy = win32gui.ScreenToClient(hwnd, (int(screen_x), int(screen_y)))
        lp = win32api.MAKELONG(max(0, cx), max(0, cy))
        win32gui.SendMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lp)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)
        time.sleep(0.04)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lp)
        win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lp)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)
    except Exception:
        pass


def _bg_set_focus(hwnd: int) -> None:
    """仅向目标控件发 WM_SETFOCUS，不 AttachThreadInput / SetFocus（避免抢前台）。"""
    if not hwnd or not HAS_WIN32:
        return
    try:
        win32gui.SendMessage(hwnd, win32con.WM_SETFOCUS, 0, 0)
    except Exception:
        pass


def _close_new_account_dialog_until_gone(
    log: LogFn = _noop_log, timeout_sec: float = 8.0
) -> bool:
    """关闭「新建帐号」直到消失。关闭瞬间短暂激活对话框（不移动鼠标），否则取消无效。"""
    end = time.time() + timeout_sec
    while time.time() < end:
        if is_foxmail_cancel_requested():
            return False
        dlg, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        if not dlg:
            return True
        with _brief_activate(dlg):
            closed = False
            for btn in _find_controls(dlg, text="取消", class_contains="TFMXButton"):
                try:
                    bl, bt, br, bb = win32gui.GetWindowRect(btn)
                    _bg_click_at(dlg, (bl + br) / 2, (bt + bb) / 2)
                    _bg_click_hwnd(btn)
                except Exception:
                    pass
                time.sleep(0.35)
                closed = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)[0] is None
                if closed:
                    break
            if not closed:
                try:
                    left, top, right, bottom = win32gui.GetWindowRect(dlg)
                    for dx in (14, 22, 30, 40):
                        _bg_click_at(dlg, right - dx, top + 12)
                        time.sleep(0.12)
                except Exception:
                    pass
                try:
                    win32gui.PostMessage(dlg, win32con.WM_CLOSE, 0, 0)
                    win32gui.PostMessage(dlg, win32con.WM_SYSCOMMAND, win32con.SC_CLOSE, 0)
                    win32gui.SendMessage(dlg, win32con.WM_SYSCOMMAND, win32con.SC_CLOSE, 0)
                except Exception:
                    pass
                # Esc
                _bg_send_key(dlg, win32con.VK_ESCAPE, times=2)
        time.sleep(0.3)
    gone = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)[0] is None
    if not gone:
        log("关闭「新建帐号」超时，窗口可能仍在")
    return gone


def _wait_new_account_list_ready(timeout_sec: float = 4.0) -> int | None:
    """等待邮箱类型列表就绪；若出现企业微信推广页则先退回。"""
    end = time.time() + timeout_sec
    while time.time() < end:
        if is_foxmail_cancel_requested():
            return None
        dlg, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        if not dlg:
            time.sleep(0.15)
            continue
        if _is_password_entry_page(dlg):
            return dlg
        if _is_wecom_promo_page(dlg):
            got = _ensure_provider_list_page(_noop_log)
            if got:
                return got
            time.sleep(0.2)
            continue
        if _page_visible(dlg, "startpage"):
            trees = _find_controls(dlg, class_contains="TVirtualDrawTree", visible_only=True)
            if trees:
                try:
                    l, t, r, b = win32gui.GetWindowRect(trees[0])
                    if (r - l) > 80 and (b - t) > 60:
                        return dlg
                except Exception:
                    return dlg
        time.sleep(0.15)
    # 最后再强退一次企业微信页
    return _ensure_provider_list_page(_noop_log)


def open_new_account_ready_for_other(log: LogFn = _noop_log) -> bool:
    """
    打开邮箱类型列表并保持对话框激活，立刻可供点「其它邮箱」。
    禁止：打开后把前台还回去（会导致第一次点击全部无效）。
    """
    dlg, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
    if dlg:
        got = _ensure_provider_list_page(log)
        if got and (
            _page_visible(got, "startpage")
            or _find_controls(got, class_contains="TVirtualDrawTree", visible_only=True)
        ):
            _keep_foxmail_dialog_active(got)
            log("已有邮箱类型列表，保持激活并使用")
            return True
        # 不是列表才关；不要无故关掉第一次打开的列表
        if got and _is_wecom_promo_page(got):
            _close_new_account_dialog_until_gone(log)
        elif not got:
            _close_new_account_dialog_until_gone(log)

    log("打开「新建帐号」列表…")
    if not click_new_account(log=log):
        return False

    dlg = _wait_new_account_list_ready(6.0)
    if not dlg:
        log("列表未就绪，关闭后重开一次…")
        _close_new_account_dialog_until_gone(log)
        time.sleep(0.35)
        if not click_new_account(log=log):
            return False
        dlg = _wait_new_account_list_ready(6.0)
    if not dlg:
        log("仍未见到邮箱类型列表")
        return False

    dlg2 = _ensure_provider_list_page(log)
    if not dlg2:
        log("无法进入邮箱类型列表")
        return False
    # 关键：保持新建帐号在前台，马上点其它邮箱；不要 sleep 后还焦点
    _keep_foxmail_dialog_active(dlg2)
    time.sleep(0.15)
    log("「新建帐号」邮箱类型列表已就绪（保持激活）")
    return True


def _bg_send_key(hwnd: int, vk: int, *, times: int = 1) -> None:
    """向指定窗口投递按键（不占用真实键盘硬件输入）。"""
    if not hwnd or not HAS_WIN32:
        return
    for _ in range(max(1, times)):
        try:
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)
        except Exception:
            try:
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk, 0)
            except Exception:
                pass
        time.sleep(0.06)


def _click_screen_restore_cursor(screen_x: float, screen_y: float) -> None:
    """禁用真实鼠标：绝不抢用户光标。"""
    return


def _uia_select_other_mailbox(dlg: int, log: LogFn = _noop_log) -> bool:
    """UIA 选「其它邮箱」。可能在 Delphi 窗口上卡死，调用方必须加超时。"""
    if not dlg:
        return False
    try:
        from pywinauto import Desktop
    except Exception:
        return False

    def _entered() -> bool:
        d, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        return bool(d and _is_password_entry_page(d))

    def _try_activate(item) -> bool:
        try:
            r = item.rectangle()
            cx = (int(r.left) + int(r.right)) / 2
            cy = (int(r.top) + int(r.bottom)) / 2
            _bg_dblclick_at(dlg, cx, cy)
            trees = _find_controls(dlg, class_contains="TVirtualDrawTree", visible_only=True)
            if trees:
                _bg_dblclick_at(trees[0], cx, cy)
            _bg_send_key(dlg, win32con.VK_RETURN, times=2)
        except Exception:
            pass
        for action in ("select", "invoke"):
            try:
                getattr(item, action)()
            except Exception:
                pass
        try:
            ia = getattr(item, "iface_legacy_iaccessible", None)
            if ia is not None:
                ia.DoDefaultAction()
        except Exception:
            pass
        time.sleep(0.4)
        return _entered()

    try:
        win = Desktop(backend="uia").window(handle=dlg)
        # 只找标题含其它邮箱的，避免全树 descendants 过慢
        for name in ("其它邮箱", "其他邮箱"):
            try:
                items = win.descendants(title=name)
            except Exception:
                items = []
            for item in items[:3]:
                if _try_activate(item):
                    log(f"已选择「其它邮箱」（UIA:{name}）")
                    return True
    except Exception as exc:
        log(f"UIA 选其它邮箱失败: {exc}")
    return False


def _uia_select_other_mailbox_timed(dlg: int, log: LogFn = _noop_log, timeout: float = 1.8) -> bool:
    """带超时的 UIA，防止在列表页卡死无点击。"""
    box: list[bool] = [False]

    def _run() -> None:
        try:
            box[0] = bool(_uia_select_other_mailbox(dlg, log=_noop_log))
        except Exception:
            box[0] = False

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    t.join(timeout)
    if t.is_alive():
        log("UIA 枚举超时，改用坐标点击「其它邮箱」")
        return False
    if box[0]:
        log("已选择「其它邮箱」（UIA）")
    return box[0]


def _click_other_mailbox_by_row(
    tree: int,
    dlg: int,
    log: LogFn = _noop_log,
    *,
    total_rows: int = 8,
    other_index: int = 7,
) -> bool:
    """
    按行号点击「其它邮箱」。
    列表共 8 项：腾讯企业邮 / QQ / Exchange / M365国际 / M365国内 / Gmail / 163 / 其它邮箱（末项）。
    """
    if not tree or not dlg:
        return False
    left, top, right, bottom = win32gui.GetWindowRect(tree)
    height = max(1, bottom - top)
    width = max(1, right - left)
    row_h = height / float(max(1, total_rows))
    # 末行中心；也点底部偏上，避开下面的「企业邮箱」链接
    ys = [
        top + row_h * other_index + row_h * 0.50,
        top + row_h * other_index + row_h * 0.40,
        top + row_h * other_index + row_h * 0.60,
        bottom - max(18, int(row_h * 0.45)),
        bottom - max(28, int(row_h * 0.55)),
    ]
    xs = [
        left + width * 0.30,
        left + width * 0.45,
        left + width * 0.55,
    ]

    _keep_foxmail_dialog_active(dlg)
    time.sleep(0.12)
    _bg_set_focus(tree)
    log(f"按行点击「其它邮箱」（第 {other_index + 1}/{total_rows} 行）…")

    # 先单击选中末行
    for sy in ys[:3]:
        for sx in xs:
            _bg_click_at(tree, sx, sy)
            _bg_click_at(dlg, sx, sy)
        time.sleep(0.06)
    _bg_send_key(tree, win32con.VK_END, times=2)
    _bg_send_key(dlg, win32con.VK_END, times=1)
    time.sleep(0.1)

    # 双击 / 回车进入
    for sy in ys:
        if sy <= top + 5 or sy >= bottom - 2:
            continue
        _keep_foxmail_dialog_active(dlg)
        for sx in xs:
            _bg_dblclick_at(tree, sx, sy)
            _bg_dblclick_at(dlg, sx, sy)
        _bg_send_key(tree, win32con.VK_RETURN, times=2)
        _bg_send_key(dlg, win32con.VK_RETURN, times=1)
        time.sleep(0.25)
        d, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        if d and _is_password_entry_page(d):
            log("已进入请输入帐号密码（按行双击）")
            _keep_foxmail_dialog_active(d)
            return True
        if d and not _page_visible(d, "startpage") and not _is_password_entry_page(d):
            if _is_wecom_promo_page(d):
                log("误点到腾讯企业邮/企业微信，将退回")
            break

    _keep_foxmail_dialog_active(dlg)
    _bg_send_key(tree, win32con.VK_END, times=2)
    _bg_send_key(tree, win32con.VK_RETURN, times=3)
    time.sleep(0.45)
    d, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
    if d and _is_password_entry_page(d):
        _keep_foxmail_dialog_active(d)
        return True
    return False


def _bg_click_hwnd(hwnd: int) -> None:
    if not hwnd:
        return
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        _bg_click_at(hwnd, (left + right) / 2, (top + bottom) / 2)
        try:
            win32gui.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)
        except Exception:
            pass
    except Exception:
        pass


def _bg_click_rel(hwnd: int, rx: float, ry: float) -> None:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    _bg_click_at(
        hwnd,
        left + (right - left) * rx,
        top + (bottom - top) * ry,
    )


def _is_password_entry_page(dlg: int) -> bool:
    """是否已到「请输入帐号密码」页（新建→其它邮箱之后）。"""
    if not dlg:
        return False
    if _page_visible(dlg, "accountEmailPage"):
        return True
    if _page_visible(dlg, "serverConfigPage"):
        return True
    has_manual = bool(
        _find_controls(dlg, text="手动设置", class_contains="TFMXButton", visible_only=True)
    )
    has_create = bool(
        _find_controls(dlg, text="创建", class_contains="TFMXButton", visible_only=True)
    )
    edits = _find_controls(dlg, class_contains="TFMEdit", visible_only=True)
    uniq = list({h for h in edits})
    return has_manual and has_create and len(uniq) >= 2


def _set_edit_text(hwnd: int, text: str) -> bool:
    """后台写入编辑框（先清空再写入，并校验）。禁止再发尾字符 WM_CHAR，否则 .com 的 m 会插到开头变成 mxxx@aol.com。"""
    if not hwnd:
        return False
    try:
        _bg_set_focus(hwnd)
        time.sleep(0.03)
        # 先彻底清空
        try:
            win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, "")
        except Exception:
            pass
        try:
            win32gui.SendMessage(hwnd, win32con.EM_SETSEL, 0, -1)
            win32gui.SendMessage(hwnd, win32con.WM_CLEAR, 0, 0)
        except Exception:
            pass
        # 写入目标文本
        win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
        try:
            win32gui.SendMessage(hwnd, win32con.EM_SETSEL, len(text), len(text))
        except Exception:
            pass
        # 通知 Delphi/父窗口内容变更（不用 WM_CHAR，避免插字符）
        try:
            parent = win32gui.GetParent(hwnd)
            cid = win32gui.GetDlgCtrlID(hwnd)
            if parent:
                wparam = win32api.MAKELONG(cid & 0xFFFF, 0x0300)  # EN_CHANGE
                win32gui.SendMessage(parent, win32con.WM_COMMAND, wparam, hwnd)
                wparam2 = win32api.MAKELONG(cid & 0xFFFF, 0x0400)  # EN_UPDATE
                win32gui.SendMessage(parent, win32con.WM_COMMAND, wparam2, hwnd)
        except Exception:
            pass
        got = (_read_edit_text(hwnd) or "").strip()
        if got == text:
            return True
        # 再试一次：全选后 EM_REPLACESEL
        try:
            win32gui.SendMessage(hwnd, win32con.EM_SETSEL, 0, -1)
            win32gui.SendMessage(hwnd, win32con.EM_REPLACESEL, 1, text)
        except Exception:
            pass
        got = (_read_edit_text(hwnd) or "").strip()
        return got == text
    except Exception:
        return False


def _read_edit_text(hwnd: int) -> str:
    try:
        import ctypes

        n = int(win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0) or 0)
        buf = ctypes.create_unicode_buffer(max(n, 1) + 8)
        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, n + 8, buf)
        return buf.value or ""
    except Exception:
        return ""


def _page_visible(dlg: int, page_name: str) -> bool:
    for hwnd in _find_controls(
        dlg, text=page_name, class_contains="TPage", visible_only=False
    ):
        try:
            if win32gui.IsWindowVisible(hwnd):
                return True
        except Exception:
            pass
    return False


def _close_new_account_dialog(log: LogFn = _noop_log) -> None:
    _close_new_account_dialog_until_gone(log, timeout_sec=5.0)


def dismiss_foxmail_success_dialogs(log: LogFn = _noop_log, timeout_sec: float = 12.0) -> int:
    """创建后点掉「设置成功」里的「完成」，否则账号不会开始收信。"""
    if not HAS_WIN32:
        return 0
    clicked = 0
    end = time.time() + timeout_sec
    while time.time() < end:
        if is_foxmail_cancel_requested():
            break
        found_btn = False

        def cb(hwnd, _):
            nonlocal clicked, found_btn
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if not _is_foxmail_process(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd) or ""
                # 标题或子控件含「设置成功」/「完成」
                interesting = any(k in title for k in ("设置成功", "设置完成", "添加成功", "创建成功"))
                btns = _find_controls(hwnd, text="完成", class_contains="TFMXButton", visible_only=True)
                if not btns:
                    btns = _find_controls(hwnd, text="完成", visible_only=True)
                if not btns and interesting:
                    # 无类名匹配时按文本扫
                    btns = _find_controls(hwnd, text="完成", visible_only=False)
                if not btns and not interesting:
                    # 仍检查子窗口标题
                    for child in _find_controls(hwnd, text_contains="设置成功", visible_only=False)[:1]:
                        interesting = True
                        break
                if not btns:
                    return True
                found_btn = True
                with _brief_activate(hwnd):
                    for btn in btns[:2]:
                        try:
                            bl, bt, br, bb = win32gui.GetWindowRect(btn)
                            _bg_click_at(hwnd, (bl + br) / 2, (bt + bb) / 2)
                        except Exception:
                            pass
                        _bg_click_hwnd(btn)
                        try:
                            win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
                        except Exception:
                            pass
                        clicked += 1
                        log("已点击「设置成功」→ 完成")
                        time.sleep(0.5)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(cb, None)
        if not found_btn and clicked:
            break
        if not found_btn:
            time.sleep(0.35)
        else:
            time.sleep(0.4)
    return clicked


def click_new_account(log: LogFn = _noop_log) -> bool:
    """点击系统设置里的「新建」。打开后保持新建窗激活，不把焦点还回去。"""
    _keep_foxmail_background()
    if find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)[0]:
        _close_new_account_dialog_until_gone(log)
        time.sleep(0.3)

    mgmt_hwnd, _ = find_window(_MGMT_KEYWORDS, foxmail_only=True)
    if not mgmt_hwnd:
        log("帐号管理/系统设置窗口不可见")
        return False

    buttons = _find_controls(mgmt_hwnd, text="新建", class_contains="TFMXButton")
    if not buttons:
        log("未找到「新建」按钮控件")
        return False

    for btn in buttons[:2]:
        # restore=False：点完新建不要抢回 HubStudio，否则列表第一次点不动
        with _brief_activate(mgmt_hwnd, restore=False):
            _bg_click_hwnd(btn)
            try:
                bl, bt, br, bb = win32gui.GetWindowRect(btn)
                _bg_click_at(mgmt_hwnd, (bl + br) / 2, (bt + bb) / 2)
            except Exception:
                pass
        time.sleep(0.55)
        dlg = _wait_new_account_list_ready(4.0)
        if dlg:
            _keep_foxmail_dialog_active(dlg)
            log("已打开「新建帐号」（保持激活）")
            return True
        if find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)[0]:
            # 开了但不是列表：关掉再试
            _close_new_account_dialog_until_gone(log)
            time.sleep(0.3)

    log("未检测到「新建帐号」列表页")
    return False


def ensure_foxmail_running(exe_path: str, log: LogFn = _noop_log) -> bool:
    if not HAS_WIN32:
        return False
    hwnd, _ = find_foxmail_main_window()
    if hwnd:
        return True
    hwnd, _ = find_window(["Foxmail"], foxmail_only=True)
    if hwnd:
        return True
    path = Path(exe_path or r"D:\APP\foxmail\Foxmail.exe")
    if not path.exists():
        log(f"未找到 Foxmail: {path}")
        return False
    log("正在启动 Foxmail...")
    subprocess.Popen([str(path)], cwd=str(path.parent))
    deadline = time.time() + 25
    while time.time() < deadline:
        hwnd, _ = find_foxmail_main_window()
        if hwnd:
            time.sleep(1.2)
            return True
        time.sleep(0.5)
    return False


def _minimize_foxmail_after(log: LogFn = _noop_log) -> None:
    """结束后恢复到桌面再最小化，避免一直停在屏外点任务栏看不见。"""
    if not HAS_WIN32:
        return
    try:
        _restore_foxmail_to_desktop(log=log, then_minimize=True)
    except Exception:
        pass


def _find_mgmt_hwnd() -> int | None:
    mgmt_hwnd, _ = find_window(_MGMT_KEYWORDS, foxmail_only=True)
    if not mgmt_hwnd:
        return None
    try:
        if win32gui.IsIconic(mgmt_hwnd):
            win32gui.ShowWindow(mgmt_hwnd, win32con.SW_SHOWNOACTIVATE)
    except Exception:
        pass
    return mgmt_hwnd


def _mgmt_already_open(log: LogFn = _noop_log) -> bool:
    if not _find_mgmt_hwnd():
        return False
    log("Foxmail 帐号管理（系统设置）已打开")
    return True


def _wait_for_mgmt_window(
    log: LogFn = _noop_log,
    *,
    timeout_sec: float = 120.0,
) -> bool:
    """
    Foxmail 右上角三横线是自绘 FMX 按钮，后台消息点不开。
    提示用户手动点开后轮询等待「系统设置」，不移动鼠标、不抢光标。
    """
    if _find_mgmt_hwnd():
        log("Foxmail 帐号管理（系统设置）已打开")
        return True

    log(
        "请手动点击 Foxmail 右上角「≡」→「帐号管理」，打开「系统设置」窗口；"
        f"打开后自动继续（最多等 {int(timeout_sec)} 秒，可点停止）…"
    )
    end = time.time() + timeout_sec
    last_ping = 0.0
    while time.time() < end:
        if is_foxmail_cancel_requested():
            log("用户已停止，取消等待帐号管理")
            return False
        if _find_mgmt_hwnd():
            log("已检测到「系统设置」，继续…")
            return True
        now = time.time()
        if now - last_ping >= 15.0:
            left = max(0, int(end - now))
            log(f"仍在等待你打开「系统设置」（剩余约 {left}s）…")
            last_ping = now
        time.sleep(0.4)
    log("等待「系统设置」超时")
    return False


def open_account_management(log: LogFn = _noop_log) -> bool:
    """
    确保「系统设置」可用。
    右上角菜单无法后台点击（会卡住空转），改为提示手动打开后继续；
    窗口一旦打开，后续「新建→其它邮箱」仍走后台消息。
    """
    if _mgmt_already_open(log):
        return True

    main_hwnd, _ = find_foxmail_main_window()
    if not main_hwnd:
        log("未找到 Foxmail 主窗口，请先打开 Foxmail")
        return False

    try:
        if win32gui.IsIconic(main_hwnd):
            win32gui.ShowWindow(main_hwnd, win32con.SW_SHOWNOACTIVATE)
            time.sleep(0.2)
    except Exception:
        pass

    # 不再做上百次坐标盲点（会长时间卡在「点击右上角菜单」）
    return _wait_for_mgmt_window(log=log, timeout_sec=120.0)


def open_settings_window(log: LogFn = _noop_log) -> bool:
    return open_account_management(log=log)


def select_other_mailbox(log: LogFn = _noop_log) -> bool:
    """列表出现后立刻点「其它邮箱」（最后一项）。全程保持对话框激活，不还焦点。"""
    if is_foxmail_cancel_requested():
        log("用户已停止")
        return False

    dlg = _ensure_provider_list_page(log)
    if not dlg:
        dlg, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
    if not dlg:
        log("新建帐号对话框不可见")
        return False
    if _is_password_entry_page(dlg):
        log("已在「请输入帐号密码」页")
        return True
    if _is_wecom_promo_page(dlg):
        log("仍停在企业微信推广页，无法选其它邮箱")
        return False

    trees = _find_controls(dlg, class_contains="TVirtualDrawTree", visible_only=True)
    if not trees and not _page_visible(dlg, "startpage"):
        log("不在邮箱类型列表")
        return False
    if not trees:
        log("未找到邮箱类型列表（TVirtualDrawTree）")
        return False
    tree = trees[0]
    _keep_foxmail_dialog_active(dlg)

    def _refresh() -> bool:
        nonlocal dlg, tree
        d, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        if not d:
            return False
        dlg = d
        if _is_password_entry_page(dlg):
            return True
        ts = _find_controls(dlg, class_contains="TVirtualDrawTree", visible_only=True)
        if not ts:
            return False
        tree = ts[0]
        return True

    # 当前界面：8 项，其它邮箱是最后一项（腾讯企业邮…163…其它邮箱）
    layouts = [(8, 7), (8, 7), (7, 6)]
    log("列表已打开，立即点击「其它邮箱」（第 8 项）…")
    for round_i in range(1, 5):
        if is_foxmail_cancel_requested():
            return False
        if not _refresh():
            time.sleep(0.15)
            continue
        if _is_password_entry_page(dlg):
            log("已进入请输入帐号密码")
            return True
        if _is_wecom_promo_page(dlg):
            log("误入企业微信页，退回…")
            if not _ensure_provider_list_page(log) or not _refresh():
                return False
            continue

        total_rows, other_index = layouts[(round_i - 1) % len(layouts)]
        # 不 restore：点选过程中绝不能把前台还给浏览器
        with _brief_activate(dlg, restore=False):
            _keep_foxmail_dialog_active(dlg)
            time.sleep(0.12)
            log(f"第 {round_i} 轮：坐标点击其它邮箱（按 {total_rows} 行布局）")
            if _click_other_mailbox_by_row(
                tree, dlg, log=log, total_rows=total_rows, other_index=other_index
            ):
                return True
            if _uia_select_other_mailbox_timed(dlg, log=log, timeout=1.5):
                return True

        if _refresh() and _is_password_entry_page(dlg):
            return True
        if _refresh() and not _page_visible(dlg, "startpage") and not _is_password_entry_page(dlg):
            log("误进其它类型，退回重试…")
            _ensure_provider_list_page(log)
            _keep_foxmail_dialog_active(
                find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)[0] or 0
            )
        time.sleep(0.15)

    if _refresh() and _is_password_entry_page(dlg):
        return True
    log("未能自动选择「其它邮箱」")
    return False



def force_foreground(hwnd: int) -> None:
    if not HAS_WIN32:
        return
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        fg = win32gui.GetForegroundWindow()
        fg_tid = win32process.GetWindowThreadProcessId(fg)[0]
        tgt_tid = win32process.GetWindowThreadProcessId(hwnd)[0]
        if fg_tid != tgt_tid:
            win32process.AttachThreadInput(fg_tid, tgt_tid, True)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)
        if fg_tid != tgt_tid:
            win32process.AttachThreadInput(fg_tid, tgt_tid, False)
    except Exception:
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass


def minimize_hubstudio_windows(log: LogFn = _noop_log) -> int:
    """最小化 HubStudio 自动化浏览器窗口（不影响用户其它程序）。"""
    if not HAS_WIN32:
        return 0
    count = 0

    def cb(hwnd, _):
        nonlocal count
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if not _is_automation_browser_window(hwnd):
            return True
        if not win32gui.IsIconic(hwnd):
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                count += 1
            except Exception:
                pass
        return True

    win32gui.EnumWindows(cb, None)
    if count:
        log(f"已最小化 {count} 个 HubStudio 窗口（后台运行）")
    return count


def restore_hubstudio_windows(log: LogFn = _noop_log) -> int:
    """截图前短暂恢复 HubStudio 窗口（仅非后台模式使用）。"""
    if not HAS_WIN32:
        return 0
    count = 0

    def cb(hwnd, _):
        nonlocal count
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if not _is_automation_browser_window(hwnd):
            return True
        if win32gui.IsIconic(hwnd):
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                count += 1
            except Exception:
                pass
        return True

    win32gui.EnumWindows(cb, None)
    if count:
        log(f"已恢复 {count} 个 HubStudio 窗口用于截图")
    return count


def _move_window_to_visible(hwnd: int) -> None:
    """将屏外/过小的 HubStudio 窗口移到屏幕可见区域。"""
    if not HAS_WIN32:
        return
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top
        if width <= 0 or height <= 0:
            width, height = 1280, 800
        screen_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        screen_h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        # 静默模式常用 --window-position=-24000,-24000，阈值须低于该值
        off_screen = (
            left < -1000
            or top < -1000
            or left >= screen_w - 40
            or top >= screen_h - 40
            or width < 200
            or height < 200
        )
        if not off_screen:
            return
        x = max(40, (screen_w - min(width, screen_w - 80)) // 2)
        y = max(40, (screen_h - min(height, screen_h - 120)) // 2)
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOP,
            x,
            y,
            min(width, screen_w - 80),
            min(height, screen_h - 120),
            win32con.SWP_SHOWWINDOW,
        )
    except Exception:
        pass


def show_hubstudio_windows_for_debug(
    log: LogFn = _noop_log,
    debug_port: int | None = None,
) -> int:
    """调试打开：恢复屏外/最小化浏览器窗口并置前（不含 HubStudio 客户端壳）。"""
    if not HAS_WIN32:
        return 0
    target_pids = _collect_browser_pids(debug_port)

    candidates: list[tuple[int, int, str]] = []

    def cb(hwnd, _):
        if not _is_browser_content_window(hwnd, target_pids or None):
            return True
        title = win32gui.GetWindowText(hwnd) or ""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            area = max(1, (rect[2] - rect[0]) * (rect[3] - rect[1]))
        except Exception:
            area = 1
        score = area
        title_l = title.lower()
        if any(m in title_l for m in _PAGE_TITLE_MARKERS):
            score += 1_000_000
        candidates.append((score, hwnd, title))
        return True

    win32gui.EnumWindows(cb, None)
    if not candidates:
        return 0

    candidates.sort(key=lambda item: item[0], reverse=True)
    count = 0
    for idx, (_, hwnd, title) in enumerate(candidates[:2]):
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            _move_window_to_visible(hwnd)
            if idx == 0:
                force_foreground(hwnd)
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNA)
            count += 1
        except Exception:
            pass
    if count:
        top_title = candidates[0][2][:60] if candidates else ""
        log(f"已显示浏览器窗口用于调试: {top_title or count}")
    return count


def reveal_hubstudio_windows_for_screenshot(log: LogFn = _noop_log) -> int:
    """截图专用：显示 HubStudio 窗口但不抢焦点（仅非后台模式使用）。"""
    if not HAS_WIN32:
        return 0
    count = 0
    show_cmd = getattr(win32con, "SW_SHOWNOACTIVATE", 4)

    def cb(hwnd, _):
        nonlocal count
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if not _is_automation_browser_window(hwnd):
            return True
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            else:
                win32gui.ShowWindow(hwnd, show_cmd)
            count += 1
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    if count:
        log(f"已显示 {count} 个 HubStudio 窗口用于截图")
    return count


def minimize_blocking_windows(log: LogFn = _noop_log) -> None:
    """临时最小化 HubStudio 窗口，避免挡住 Foxmail。"""
    global _minimized_windows
    _minimized_windows = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if not _is_automation_browser_window(hwnd):
            return True
        if not win32gui.IsIconic(hwnd):
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                _minimized_windows.append(hwnd)
            except Exception:
                pass
        return True

    if HAS_WIN32:
        win32gui.EnumWindows(cb, None)
        if _minimized_windows:
            log(f"已临时最小化 {len(_minimized_windows)} 个浏览器窗口")
        time.sleep(0.4)


def restore_minimized_windows(log: LogFn = _noop_log) -> None:
    global _minimized_windows
    for hwnd in _minimized_windows:
        try:
            if win32gui.IsWindow(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass
    if _minimized_windows:
        log("已恢复浏览器窗口")
    _minimized_windows = []


def focus_window(hwnd: int) -> None:
    if not HAS_WIN32:
        return
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    except Exception:
        pass
    force_foreground(hwnd)


def _click_rel(rect: tuple[int, int, int, int], rx: float, ry: float) -> None:
    """兼容旧调用：改为后台消息点击，不再移动真实鼠标。"""
    # 无 hwnd 时无法 SendMessage；此函数仅保留兼容，实际 Foxmail 流程已改用 _bg_click_*
    return


def _smtp_for_email(email: str) -> tuple[str, int]:
    domain = (email.split("@")[-1] or "").lower()
    mapping = {
        "aol.com": ("smtp.aol.com", 587),
        "qq.com": ("smtp.qq.com", 465),
        "foxmail.com": ("smtp.qq.com", 465),
        "163.com": ("smtp.163.com", 465),
        "126.com": ("smtp.126.com", 465),
        "gmail.com": ("smtp.gmail.com", 587),
        "outlook.com": ("smtp.office365.com", 587),
        "hotmail.com": ("smtp.office365.com", 587),
        "live.com": ("smtp.office365.com", 587),
    }
    return mapping.get(domain, (f"smtp.{domain}", 587))


def _imap_for_email(email: str) -> tuple[str, int]:
    domain = (email.split("@")[-1] or "").lower()
    mapping = {
        "aol.com": ("imap.aol.com", 993),
        "qq.com": ("imap.qq.com", 993),
        "foxmail.com": ("imap.qq.com", 993),
        "163.com": ("imap.163.com", 993),
        "126.com": ("imap.126.com", 993),
        "gmail.com": ("imap.gmail.com", 993),
        "outlook.com": ("outlook.office365.com", 993),
        "hotmail.com": ("outlook.office365.com", 993),
        "live.com": ("outlook.office365.com", 993),
    }
    return mapping.get(domain, (f"imap.{domain}", 993))


def fill_other_mailbox_form(
    recovery_email: str,
    recovery_password: str,
    log: LogFn = _noop_log,
    *,
    _depth: int = 0,
) -> bool:
    """在「请输入帐号密码」页填写辅助邮箱+密码并点创建。"""
    deadline = time.time() + 15
    dlg = None
    while time.time() < deadline:
        dlg, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        if dlg and _is_password_entry_page(dlg):
            break
        dlg = None
        time.sleep(0.35)

    if not dlg:
        log("未找到「请输入帐号密码」页面（需先：新建→其它邮箱）")
        return False

    # 填写/点创建时短暂激活（不移动鼠标），否则 FMX 控件不吃消息
    _force_activate_no_mouse(dlg)
    time.sleep(0.1)

    imap_host, imap_port = _imap_for_email(recovery_email)
    smtp_host, smtp_port = _smtp_for_email(recovery_email)

    def _visible_edits() -> list[int]:
        edits = _find_controls(dlg, class_contains="TFMEdit", visible_only=True)
        uniq: list[int] = []
        seen: set[int] = set()
        for h in edits:
            if h in seen:
                continue
            seen.add(h)
            uniq.append(h)
        uniq.sort(key=lambda h: win32gui.GetWindowRect(h)[1])
        return uniq

    if _page_visible(dlg, "serverConfigPage"):
        log(f"检测到服务器设置页，后台填写 {recovery_email}")
        edits = _visible_edits()
        values = [
            recovery_email,
            recovery_password,
            imap_host,
            str(imap_port),
            smtp_host,
            str(smtp_port),
        ]
        for hwnd, val in zip(edits, values):
            _set_edit_text(hwnd, val)
            time.sleep(0.08)
    else:
        edits = _visible_edits()
        if len(edits) < 2:
            log("请输入帐号密码页未找到邮箱/密码输入框")
            return False
        log(f"填写请输入帐号密码: {recovery_email}")
        ok1 = _set_edit_text(edits[0], recovery_email)
        time.sleep(0.12)
        ok2 = _set_edit_text(edits[1], recovery_password)
        time.sleep(0.12)
        # 强制校验：邮箱必须与设定完全一致（防止残留字符变成 mxxx@...）
        got_email = (_read_edit_text(edits[0]) or "").strip()
        if got_email != recovery_email.strip():
            log(f"邮箱写入异常，当前为 [{got_email}]，重写为 [{recovery_email}]")
            win32gui.SendMessage(edits[0], win32con.WM_SETTEXT, 0, "")
            time.sleep(0.05)
            ok1 = _set_edit_text(edits[0], recovery_email)
            got_email = (_read_edit_text(edits[0]) or "").strip()
            if got_email != recovery_email.strip():
                log(f"写入辅助邮箱失败，读回仍是 [{got_email}]")
                return False
        # Tab 触发校验，使「创建」可点
        try:
            win32gui.PostMessage(edits[1], win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
            win32gui.PostMessage(edits[1], win32con.WM_KEYUP, win32con.VK_TAB, 0)
        except Exception:
            pass
        time.sleep(0.2)
        if not ok2 and not (_read_edit_text(edits[1]) or "").strip():
            log("写入辅助邮箱密码失败")
            return False
        log(f"已确认填写辅助邮箱: {got_email}")
    create_btns = _find_controls(dlg, text="创建", class_contains="TFMXButton")
    if not create_btns:
        log("未找到「创建」按钮")
        return False

    for _ in range(4):
        btn = create_btns[0]
        enabled = True
        try:
            enabled = bool(win32gui.IsWindowEnabled(btn))
        except Exception:
            pass
        if not enabled:
            edits = _visible_edits()
            if len(edits) >= 2:
                _set_edit_text(edits[0], recovery_email)
                _set_edit_text(edits[1], recovery_password)
                time.sleep(0.25)
        # 点创建（消息点父窗中心也试一次）
        _bg_click_hwnd(btn)
        try:
            bl, bt, br, bb = win32gui.GetWindowRect(btn)
            _bg_click_at(dlg, (bl + br) / 2, (bt + bb) / 2)
        except Exception:
            pass
        time.sleep(1.0)
        # 对话框消失或离开密码页视为成功
        still, _ = find_window(_NEW_ACCOUNT_KEYWORDS, foxmail_only=True)
        if not still:
            log("已点击创建，新建帐号对话框已关闭")
            return True
        if still and not _is_password_entry_page(still):
            if _page_visible(still, "serverConfigPage") and _depth < 1:
                return fill_other_mailbox_form(
                    recovery_email, recovery_password, log=log, _depth=_depth + 1
                )
            log("已提交创建（页面已切换）")
            return True

    log("已尝试点击创建")
    return True


def account_exists_in_foxmail(recovery_email: str, foxmail_dir: str) -> bool:
    base = Path(foxmail_dir or r"D:\APP\foxmail")
    storage = base / "Storage" / recovery_email.lower()
    if storage.exists():
        return True
    listing = base / "FMStorage.list"
    if listing.exists():
        needle = recovery_email.lower()
        try:
            text = listing.read_text(encoding="utf-8", errors="ignore").lower()
            return needle in text
        except Exception:
            pass
    return False


def create_other_mailbox_account(
    recovery_email: str,
    recovery_password: str,
    config: dict[str, Any],
    log: LogFn = _noop_log,
    force_ui: bool = False,
) -> tuple[bool, str]:
    """
    Foxmail：系统设置 → 新建 → 其它邮箱 → 请输入帐号密码 → 创建。
    并行任务会串行占用 Foxmail UI（屏外消息操作，结束后恢复到桌面）。
    """
    if not recovery_email or not recovery_password:
        return False, "缺少辅助邮箱或密码"

    fx = _foxmail_cfg(config)
    if not fx.get("enabled", True):
        return True, "Foxmail 已禁用"

    # 未强制、且未开启 use_ui / auto_create 时跳过界面
    if (
        not force_ui
        and not fx.get("use_ui", False)
        and not fx.get("auto_create", True)
    ):
        return True, "已使用 IMAP 后台收信，跳过 Foxmail 界面"

    if not HAS_WIN32:
        return False, "缺少 Windows 自动化依赖（pywin32）"

    foxmail_dir = fx.get("data_dir") or r"D:\APP\foxmail"
    exe_path = fx.get("exe_path") or str(Path(foxmail_dir) / "Foxmail.exe")

    # 可中断地拿锁：点停止后最多约 0.4s 就能退出排队，不会一直卡住
    while True:
        if is_foxmail_cancel_requested():
            return False, "用户已停止"
        acquired = _foxmail_ui_lock.acquire(timeout=0.4)
        if acquired:
            break

    try:
        if is_foxmail_cancel_requested():
            return False, "用户已停止"

        if account_exists_in_foxmail(recovery_email, foxmail_dir):
            log(f"Foxmail 已有账号 {recovery_email}，跳过创建")
            return True, "Foxmail 账号已存在"

        if not ensure_foxmail_running(exe_path, log=log):
            return False, "Foxmail 未运行且无法启动"

        # 窗口消息操作，不抢前台；窗口留在桌面，点任务栏可显示
        try:
            if fx.get("park_offscreen", False):
                _park_foxmail_offscreen(log=log)

            if fx.get("require_settings_open", False):
                if not _find_mgmt_hwnd():
                    log("配置要求先打开系统设置，正在等待你手动打开…")
                    if not _wait_for_mgmt_window(log=log, timeout_sec=120.0):
                        return False, "请手动打开 Foxmail 帐号管理（系统设置）后再运行"

            if not open_account_management(log=log):
                return False, "未打开 Foxmail「系统设置」（请点右上角≡→帐号管理）"

            if is_foxmail_cancel_requested():
                return False, "用户已停止"

            # 列表一旦出现就立刻选其它邮箱；不要先关窗空转
            selected = False
            for attempt in range(1, 4):
                if is_foxmail_cancel_requested():
                    return False, "用户已停止"
                log(f"准备选择「其它邮箱」（第 {attempt}/3 轮）…")
                if not open_new_account_ready_for_other(log=log):
                    continue
                log("列表就绪，开始点击「其它邮箱」…")
                if select_other_mailbox(log=log):
                    selected = True
                    break
                log("本轮未进入请输入帐号密码，关闭后重试…")
                _close_new_account_dialog_until_gone(log)
                time.sleep(0.45)

            if not selected:
                return False, "无法选择「其它邮箱」"

            if is_foxmail_cancel_requested():
                return False, "用户已停止"

            if not fill_other_mailbox_form(recovery_email, recovery_password, log=log):
                return False, "无法在 Foxmail 填写账密"

            deadline = time.time() + int(fx.get("create_confirm_sec", 15))
            created = False
            while time.time() < deadline:
                if is_foxmail_cancel_requested():
                    return False, "用户已停止"
                # 创建后常弹出「设置成功」→ 必须点「完成」才开始收信
                dismiss_foxmail_success_dialogs(log=log, timeout_sec=1.5)
                if account_exists_in_foxmail(recovery_email, foxmail_dir):
                    created = True
                    break
                time.sleep(0.4)

            dismiss_foxmail_success_dialogs(log=log, timeout_sec=8.0)
            if created:
                log(f"Foxmail 已创建账号: {recovery_email}")
                return True, "Foxmail 其它邮箱创建成功"
            return True, "Foxmail 已提交创建（目录尚未更新，继续尝试收信）"
        finally:
            # 不抢前台：仅在曾屏外停靠时才无激活拉回
            if fx.get("park_offscreen", False):
                _restore_foxmail_to_desktop(
                    log=log,
                    then_minimize=bool(fx.get("minimize_after", False)),
                )
    finally:
        _foxmail_ui_lock.release()
