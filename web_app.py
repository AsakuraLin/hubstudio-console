# -*- coding: utf-8 -*-
"""HubStudio 批量打开 Web 控制台（支持并行）。"""

from __future__ import annotations

import json
import re
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request

from app_paths import config_path, data_dir, ensure_user_config, templates_dir
from hubstudio_client import ensure_hubstudio_ready, hubstudio_api_ok, resolve_hubstudio_exe

from outlook_checker import (
    AccountStatus,
    CheckResult,
    apply_manual_recovery_code,
    apply_runtime_config,
    build_env_cache,
    check_one_account,
    clear_env_cache,
    clear_task_context,
    fill_only_one,
    get_pending_recovery_code,
    init_screenshot_session,
    load_json,
    open_only_one,
    resolve_env,
    save_results,
    set_env_cache,
    set_task_context,
    stop_all_browsers,
    stop_browsers_by_machine_ids,
    open_browser_for_debug,
)
from excel_sync import (
    excel_sync_config,
    is_machine_not_found_message,
    list_sheet_names,
    update_machine_status,
)
from foxmail_automation import clear_foxmail_cancel, request_foxmail_cancel

BASE_DIR = data_dir()
ensure_user_config()
CONFIG_PATH = config_path()

app = Flask(__name__, template_folder=str(templates_dir()))


def cleanup_old_results(base_dir: Path, keep: int = 5) -> tuple[int, int]:
    """只保留最近 N 份 results CSV，避免堆积占磁盘。返回 (删除数, 保留数)。"""
    files = sorted(
        base_dir.glob("results_*.csv"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    removed = 0
    for old in files[keep:]:
        try:
            old.unlink()
            removed += 1
        except OSError:
            pass
    kept = min(len(files), keep)
    return removed, kept


def cleanup_screenshots(base_dir: Path) -> tuple[int, int]:
    """删除 imgs 目录下全部截图。返回 (删除文件数, 释放字节数)。"""
    imgs = base_dir / "imgs"
    if not imgs.exists():
        return 0, 0
    removed = 0
    freed = 0
    for path in imgs.rglob("*"):
        if not path.is_file():
            continue
        try:
            freed += path.stat().st_size
            path.unlink()
            removed += 1
        except OSError:
            pass
    for path in sorted(imgs.rglob("*"), reverse=True):
        if path.is_dir():
            try:
                path.rmdir()
            except OSError:
                pass
    imgs.mkdir(parents=True, exist_ok=True)
    return removed, freed


def _format_bytes(size: int) -> str:
    if size >= 1024 * 1024:
        return f"{size / (1024 * 1024):.1f} MB"
    if size >= 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size} B"


@app.get("/files/<path:rel_path>")
def serve_project_file(rel_path: str):
    """提供 imgs 目录下的截图访问。"""
    from urllib.parse import unquote

    if ".." in rel_path.replace("\\", "/"):
        return jsonify({"ok": False, "error": "invalid path"}), 400
    # 兼容中文文件名（超时_、AMZ账号被封_ 等）
    norm = unquote(rel_path).replace("\\", "/")
    if not norm.startswith("imgs/"):
        return jsonify({"ok": False, "error": "not found"}), 404
    target = (BASE_DIR / norm).resolve()
    if not str(target).startswith(str(BASE_DIR.resolve())) or not target.is_file():
        return jsonify({"ok": False, "error": "not found"}), 404
    from flask import send_file

    return send_file(target)

_jobs: dict[str, dict] = {}
_jobs_lock = threading.Lock()
_MAX_JOBS_KEPT = 20


def _sanitize_pair(pair: dict) -> dict:
    row = dict(pair)
    row.pop("password", None)
    row.pop("recovery_password", None)
    return row


def _sanitize_job(job: dict) -> dict:
    safe = dict(job)
    safe["pairs"] = [_sanitize_pair(p) for p in (job.get("pairs") or [])]
    pool = job.get("recovery_pool") or []
    claimed = job.get("recovery_claimed") or []
    safe["recovery_emails_remaining"] = [x.get("email", "") for x in pool]
    safe["recovery_passwords_remaining"] = [x.get("password", "") for x in pool]
    safe["recovery_claimed"] = [
        {
            "machine_id": c.get("machine_id", ""),
            "recovery_email": c.get("recovery_email", ""),
            "recovery_password": c.get("recovery_password", ""),
            "at": c.get("at", ""),
        }
        for c in claimed
    ]
    # 不把含密码的原始 pool 发给前端
    safe.pop("recovery_pool", None)
    return safe


def _prune_old_jobs(keep: int = _MAX_JOBS_KEPT) -> None:
    with _jobs_lock:
        if len(_jobs) <= keep:
            return
        finished = [
            (jid, j)
            for jid, j in _jobs.items()
            if j.get("status") in {"done", "cancelled"}
        ]
        finished.sort(key=lambda x: x[1].get("finished_at") or x[1].get("started_at") or "")
        for jid, _ in finished[: max(0, len(_jobs) - keep)]:
            _jobs.pop(jid, None)


def load_config() -> dict:
    if CONFIG_PATH.exists():
        return load_json(CONFIG_PATH)
    return {
        "hubstudio_api": "http://127.0.0.1:6873/api/v1",
        "hubstudio_exe_path": "",
        "local_api_key": "",
        "login_url": "https://outlook.live.com/mail/",
        "wait_timeout_sec": 30,
        "manual_mfa_timeout_sec": 120,
        "close_browser_after_check": True,
        "batch_concurrency": 6,
    }


def save_config(config: dict) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def public_settings(config: dict | None = None) -> dict:
    cfg = config or load_config()
    exe_raw = cfg.get("hubstudio_exe_path") or ""
    exe = resolve_hubstudio_exe(exe_raw)
    excel = cfg.get("excel_sync") or {}
    api_base = cfg.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = cfg.get("local_api_key", "")
    return {
        "hubstudio_exe_path": exe_raw,
        "hubstudio_exe_resolved": str(exe) if exe else "",
        "hubstudio_api": api_base,
        "local_api_key": cfg.get("local_api_key", ""),
        "login_url": cfg.get("login_url") or "",
        "excel_file_path": excel.get("file_path", ""),
        "api_connected": hubstudio_api_ok(api_base, api_key),
        "config_path": str(CONFIG_PATH),
    }


def parse_lines(text: str) -> list[str]:
    rows: list[str] = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        rows.append(line)
    return rows


def build_pairs(
    machine_text: str,
    email_text: str,
    password_text: str,
    recovery_email_text: str = "",
    recovery_password_text: str = "",
    mode: str = "fill_only",
) -> tuple[list[dict], str | None]:
    machines = parse_lines(machine_text)
    emails = parse_lines(email_text)
    passwords = parse_lines(password_text)
    recovery_emails = parse_lines(recovery_email_text)
    recovery_passwords = parse_lines(recovery_password_text)

    if not machines:
        return [], "请至少输入一个机子号"

    # 仅打开环境：只需机子号
    if mode == "open_only":
        pairs = []
        for i, machine_id in enumerate(machines):
            pairs.append(
                {
                    "index": i + 1,
                    "machine_id": machine_id,
                    "email": emails[i] if i < len(emails) else "",
                    "password": passwords[i] if i < len(passwords) else "",
                    # 辅助邮箱不预分配，运行时按需领取
                    "recovery_email": "",
                    "recovery_password": "",
                }
            )
        return pairs, None

    if not emails:
        return [], "请至少输入一个邮箱"
    if not passwords:
        return [], "请至少输入一个密码"

    # 辅助邮箱是共享池，不限制机子/邮箱配对数量
    count = min(len(machines), len(emails), len(passwords))

    warnings: list[str] = []
    lengths = [len(machines), len(emails), len(passwords)]
    if len(set(lengths)) > 1:
        warnings.append(
            f"各列行数不一致（机子号 {len(machines)} / 邮箱 {len(emails)} / 密码 {len(passwords)}"
            + f"），将按顺序配对前 {count} 条"
        )
    if recovery_emails:
        warnings.append(
            f"辅助邮箱 {len(recovery_emails)} 条为共享池："
            "只有真正进入「绑定辅助邮箱」页的机子才会领取一条；"
            "领取后从输入框移除并显示在该机子号旁"
        )
    if any(not p for p in passwords[:count]):
        warnings.append("存在空密码行，请检查对应关系")
    if recovery_emails and recovery_passwords:
        if len(recovery_passwords) < len(recovery_emails):
            warnings.append(
                f"辅助邮箱密码行数({len(recovery_passwords)})少于辅助邮箱({len(recovery_emails)})，"
                "缺密码的条目在领取时可能无法自动收信（人工验证码模式仍可用）"
            )

    pairs = []
    for i in range(count):
        pairs.append(
            {
                "index": i + 1,
                "machine_id": machines[i],
                "email": emails[i],
                "password": passwords[i],
                "recovery_email": "",
                "recovery_password": "",
            }
        )
    warning = "；".join(warnings) if warnings else None
    return pairs, warning


def build_recovery_pool(
    recovery_email_text: str = "",
    recovery_password_text: str = "",
) -> list[dict]:
    """辅助邮箱共享池：email + password，运行时按需领取。"""
    emails = parse_lines(recovery_email_text)
    passwords = parse_lines(recovery_password_text)
    pool: list[dict] = []
    for i, em in enumerate(emails):
        pool.append(
            {
                "email": em,
                "password": passwords[i] if i < len(passwords) else "",
            }
        )
    return pool


def claim_recovery_from_job(job_id: str, machine_id: str) -> tuple[str, str] | None:
    """
    机子真正进入辅助邮箱绑定页时领取一条。
    同一机子重复领取返回已领的那条；池空返回 None。

    recovery_assign_ordered=True 时：严格按机子列表顺序领取——
    只有序号更靠前的机子都已「出结果或已领过」后，当前机子才能从池头取下一条。
    这样并行跑时，辅助邮箱仍按你粘贴进池子的先后顺序，对应到机子列表从前到后需要绑定的机子。
    """
    mid = (machine_id or "").strip()
    if not mid:
        return None

    deadline = time.time() + 600.0  # 最多等 10 分钟轮到自己
    while time.time() < deadline:
        claimed_email = ""
        claimed_password = ""
        should_wait = False
        with _jobs_lock:
            job = _jobs.get(job_id)
            if not job or job.get("cancelled"):
                return None
            for row in job.get("recovery_claimed") or []:
                if row.get("machine_id") == mid:
                    return (
                        (row.get("recovery_email") or "").strip(),
                        (row.get("recovery_password") or ""),
                    )
            pool = job.setdefault("recovery_pool", [])
            if not pool:
                return None

            ordered = bool(job.get("recovery_assign_ordered"))
            if ordered:
                pairs = job.get("pairs") or []
                my_index = None
                for p in pairs:
                    if (p.get("machine_id") or "").strip() == mid:
                        my_index = int(p.get("index") or 0)
                        break
                if my_index is not None:
                    claimed_ids = {
                        (c.get("machine_id") or "").strip()
                        for c in (job.get("recovery_claimed") or [])
                    }
                    finished_ids = {
                        (r.get("machine_id") or "").strip()
                        for r in (job.get("results") or [])
                    }
                    blockers = []
                    for p in pairs:
                        idx = int(p.get("index") or 0)
                        if idx <= 0 or idx >= my_index:
                            continue
                        pmid = (p.get("machine_id") or "").strip()
                        if not pmid:
                            continue
                        # 前面的机子：已领过或已出最终结果 → 不挡；否则等待
                        if pmid in claimed_ids or pmid in finished_ids:
                            continue
                        blockers.append(pmid)
                    if blockers:
                        should_wait = True

            if not should_wait:
                item = pool.pop(0)
                claimed_email = (item.get("email") or "").strip()
                claimed_password = item.get("password") or ""
                job.setdefault("recovery_claimed", []).append(
                    {
                        "machine_id": mid,
                        "recovery_email": claimed_email,
                        "recovery_password": claimed_password,
                        "at": datetime.now().strftime("%H:%M:%S"),
                    }
                )

        if should_wait:
            time.sleep(0.35)
            continue

        if claimed_email:
            _append_activity(job_id, mid, f"领取辅助邮箱: {claimed_email}")
            return claimed_email, claimed_password
        return None

    _append_activity(job_id, mid, "按顺序领取辅助邮箱超时（前面机子未结束）")
    return None


def release_recovery_claim(job_id: str, machine_id: str, *, reason: str = "") -> bool:
    """误领/绑定失败时把辅助邮箱还回池头，避免打乱后续顺序。"""
    mid = (machine_id or "").strip()
    if not mid or not job_id:
        return False
    returned_email = ""
    returned_password = ""
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return False
        claimed = job.setdefault("recovery_claimed", [])
        keep = []
        item = None
        for row in claimed:
            if (row.get("machine_id") or "").strip() == mid and item is None:
                item = row
            else:
                keep.append(row)
        if item is None:
            return False
        job["recovery_claimed"] = keep
        returned_email = (item.get("recovery_email") or "").strip()
        returned_password = item.get("recovery_password") or ""
        if returned_email:
            job.setdefault("recovery_pool", []).insert(
                0, {"email": returned_email, "password": returned_password}
            )
    if returned_email:
        tip = f"归还辅助邮箱到池: {returned_email}"
        if reason:
            tip += f"（{reason}）"
        _append_activity(job_id, mid, tip)
        return True
    return False


def result_to_dict(result) -> dict:
    row = asdict(result)
    row["status"] = result.status.value
    return row


def _append_activity(job_id: str, machine_id: str, step: str) -> None:
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        ts = datetime.now().strftime("%H:%M:%S")
        job.setdefault("running_steps", {})[machine_id] = {"step": step, "at": ts}
        log = job.setdefault("activity_log", [])
        log.append({"at": ts, "machine_id": machine_id, "step": step})
        if len(log) > 100:
            job["activity_log"] = log[-100:]


def _make_step_reporter(job_id: str):
    def reporter(machine_id: str, step: str) -> None:
        _append_activity(job_id, machine_id, step)

    return reporter


def process_one_pair(
    pair: dict,
    mode: str,
    api_base: str,
    api_key: str,
    login_url: str,
    wait_timeout_sec: int,
    manual_mfa_timeout_sec: int,
    close_browser: bool,
    recovery_config: dict,
    job_id: str = "",
) -> dict:
    machine_id = pair["machine_id"]
    email = pair["email"]
    password = pair["password"]
    recovery_email = pair.get("recovery_email", "")
    recovery_password = pair.get("recovery_password", "")
    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 注入「进入绑定页才领取」回调；不预分配辅助邮箱
    cfg = dict(recovery_config or {})
    if job_id:
        cfg["claim_recovery"] = lambda mid, jid=job_id: claim_recovery_from_job(jid, mid)
        cfg["release_recovery"] = (
            lambda mid, reason="", jid=job_id: release_recovery_claim(
                jid, mid, reason=reason
            )
        )

    if job_id:
        set_task_context(machine_id, _make_step_reporter(job_id))
        _append_activity(job_id, machine_id, "开始处理")
        _append_activity(job_id, machine_id, "查找 HubStudio 环境")
    try:
        env = resolve_env(api_base, machine_id, api_key)
        if env is None:
            return {
                "index": pair["index"],
                "machine_id": machine_id,
                "email": email,
                "status": AccountStatus.RESOLVE_FAILED.value,
                "detail": "未找到对应 HubStudio 环境",
                "container_code": "",
                "final_url": "",
                "checked_at": checked_at,
            }

        try:
            if mode == "open_only":
                result = open_only_one(api_base, api_key, login_url, env)
            elif mode == "fill_only":
                result = fill_only_one(
                    api_base,
                    api_key,
                    login_url,
                    env,
                    email,
                    password,
                    recovery_email=recovery_email,
                    recovery_password=recovery_password,
                    recovery_config=cfg,
                    close_browser=close_browser,
                )
            else:
                result = check_one_account(
                    api_base=api_base,
                    api_key=api_key,
                    login_url=login_url,
                    env=env,
                    email=email,
                    password=password,
                    wait_timeout_sec=wait_timeout_sec,
                    manual_mfa_timeout_sec=manual_mfa_timeout_sec,
                    close_browser=close_browser,
                    recovery_email=recovery_email,
                    recovery_password=recovery_password,
                    recovery_config=cfg,
                )
            item = result_to_dict(result)
            item["index"] = pair["index"]
            if not item.get("recovery_email"):
                item["recovery_email"] = getattr(result, "recovery_email", "") or recovery_email
            if job_id:
                with _jobs_lock:
                    job = _jobs.get(job_id)
                    if job:
                        for c in job.get("recovery_claimed") or []:
                            if c.get("machine_id") == machine_id:
                                if not item.get("recovery_email"):
                                    item["recovery_email"] = c.get("recovery_email") or ""
                                item["recovery_password"] = c.get("recovery_password") or ""
                                break
            if result.status == AccountStatus.WAIT_CODE:
                item["awaiting_code"] = True
            return item
        except Exception as exc:
            detail = str(exc)
            if any(k in detail for k in ("Read timed out", "ConnectionPool", "CDP 连接失败", "打开浏览器超时")):
                status = AccountStatus.OPEN_FAILED.value
                detail = f"浏览器连接超时（并行过多，建议降至并行 3）: {exc}"
            else:
                status = AccountStatus.UNKNOWN.value
                detail = f"执行异常: {exc}"
            return {
                "index": pair["index"],
                "machine_id": machine_id,
                "email": email,
                "status": status,
                "detail": detail,
                "container_code": env.container_code if env else "",
                "final_url": "",
                "checked_at": checked_at,
            }
    finally:
        if job_id:
            with _jobs_lock:
                job = _jobs.get(job_id)
                if job:
                    job.get("running_steps", {}).pop(machine_id, None)
            clear_task_context()


def _job_cancelled(job_id: str) -> bool:
    with _jobs_lock:
        job = _jobs.get(job_id)
        return bool(job and job.get("cancelled"))


def _update_job_running(job_id: str, machine_id: str, add: bool) -> None:
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        running = job.setdefault("running_machines", [])
        if add and machine_id not in running:
            running.append(machine_id)
        elif not add and machine_id in running:
            running.remove(machine_id)
        job["running_count"] = len(running)


def _close_running_job_browsers(job_id: str) -> None:
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        machine_ids = list(job.get("running_machines") or [])
    if not machine_ids:
        return
    config = load_config()
    api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = config.get("local_api_key", "")
    stop_browsers_by_machine_ids(api_base, machine_ids, api_key)


def _append_job_result(job_id: str, item: dict) -> None:
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        mid = item.get("machine_id")
        ov = (job.get("status_overrides") or {}).get(mid) if mid else None
        if ov:
            item = dict(item)
            item["status"] = ov.get("status") or item.get("status")
            item["detail"] = ov.get("detail") or item.get("detail")
            item["awaiting_code"] = False
        results = job.setdefault("results", [])
        if mid:
            results[:] = [r for r in results if r.get("machine_id") != mid]
        results.append(item)
        results.sort(key=lambda x: x.get("index", 0))
        job["completed"] = len(results)


def _maybe_sync_excel(job_id: str, item: dict) -> None:
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        excel = job.get("excel_sync") or {}
    if not excel.get("enabled"):
        return
    file_path = excel.get("file_path") or ""
    sheet_name = excel.get("sheet_name") or ""
    if not file_path or not sheet_name:
        return

    machine_id = item.get("machine_id", "")
    status = item.get("status", "")
    detail = item.get("detail", "")
    ok, msg = update_machine_status(
        file_path,
        sheet_name,
        machine_id,
        status,
        detail,
        machine_headers=excel.get("machine_headers"),
        status_headers=excel.get("status_headers"),
        include_detail=bool(excel.get("include_detail")),
        write_while_open=bool(excel.get("write_while_open", True)),
    )
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        sync = job.setdefault("excel_sync", {})
        if ok:
            sync["updated"] = int(sync.get("updated", 0)) + 1
            _append_activity(job_id, machine_id, f"Excel 回写: {status}")
        elif is_machine_not_found_message(msg):
            sync["skipped"] = int(sync.get("skipped", 0)) + 1
        else:
            errors = sync.setdefault("errors", [])
            errors.append({"machine_id": machine_id, "message": msg})
            if len(errors) > 30:
                sync["errors"] = errors[-30:]
            _append_activity(job_id, machine_id, f"Excel 回写失败: {msg}")


def run_batch_job(job_id: str, pairs: list[dict], mode: str, concurrency: int) -> None:
    config = load_config()
    apply_runtime_config(config)
    init_screenshot_session(BASE_DIR / "imgs")
    api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = config.get("local_api_key", "")
    login_url = config["login_url"]
    wait_timeout_sec = int(config.get("wait_timeout_sec", 30))
    manual_mfa_timeout_sec = int(config.get("manual_mfa_timeout_sec", 120))
    close_browser = bool(config.get("close_browser_after_check", True))
    results_keep = int(config.get("results_keep_count", 5))

    stagger_sec = float(config.get("browser_start_stagger_sec", 1.5))

    try:
        cache = build_env_cache(api_base, api_key)
        set_env_cache(cache)
    except Exception:
        set_env_cache(None)

    try:
        _run_batch_job_body(
            job_id,
            pairs,
            mode,
            concurrency,
            api_base,
            api_key,
            login_url,
            wait_timeout_sec,
            manual_mfa_timeout_sec,
            close_browser,
            results_keep,
            stagger_sec,
            config,
        )
    finally:
        clear_env_cache()


def _run_batch_job_body(
    job_id: str,
    pairs: list[dict],
    mode: str,
    concurrency: int,
    api_base: str,
    api_key: str,
    login_url: str,
    wait_timeout_sec: int,
    manual_mfa_timeout_sec: int,
    close_browser: bool,
    results_keep: int,
    stagger_sec: float,
    config: dict,
) -> None:
    def _run_pair(pair: dict) -> dict | None:
        if _job_cancelled(job_id):
            return None
        machine_id = pair["machine_id"]
        # 错峰启动：仅在同一并行批次内错开，避免第 N 个任务等 (N-1)*1.5 秒
        slot = (int(pair.get("index", 1)) - 1) % max(1, concurrency)
        delay = slot * stagger_sec
        if delay:
            time.sleep(delay)
        _update_job_running(job_id, machine_id, add=True)
        with _jobs_lock:
            job = _jobs.get(job_id)
            if job:
                job["current_machine"] = machine_id
        try:
            return process_one_pair(
                pair,
                mode,
                api_base,
                api_key,
                login_url,
                wait_timeout_sec,
                manual_mfa_timeout_sec,
                close_browser,
                config,
                job_id=job_id,
            )
        finally:
            _update_job_running(job_id, machine_id, add=False)

    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        futures = [executor.submit(_run_pair, pair) for pair in pairs]
        for future in as_completed(futures):
            if _job_cancelled(job_id):
                with _jobs_lock:
                    job = _jobs.get(job_id)
                    if job:
                        job["status"] = "cancelled"
                _close_running_job_browsers(job_id)
                break
            item = future.result()
            if item:
                _append_job_result(job_id, item)
                threading.Thread(
                    target=_maybe_sync_excel,
                    args=(job_id, item),
                    daemon=True,
                ).start()

    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        if job.get("status") != "cancelled":
            job["status"] = "done"
        job["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        job["running_machines"] = []
        job["running_count"] = 0
        if job["results"]:
            out = BASE_DIR / f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            check_results = []
            for row in job["results"]:
                status_text = row.get("status", AccountStatus.UNKNOWN.value)
                try:
                    status = AccountStatus(status_text)
                except ValueError:
                    status = AccountStatus.UNKNOWN
                check_results.append(
                    CheckResult(
                        machine_id=row["machine_id"],
                        email=row.get("email", ""),
                        status=status,
                        detail=row.get("detail", ""),
                        container_code=row.get("container_code", ""),
                        final_url=row.get("final_url", ""),
                        checked_at=row.get("checked_at", ""),
                    )
                )
            save_results(check_results, out)
            cleanup_old_results(BASE_DIR, keep=results_keep)
            job["result_file"] = str(out)
    _prune_old_jobs()


@app.route("/")
def index():
    return render_template("index.html")


@app.post("/api/preview")
def preview():
    data = request.get_json(force=True) or {}
    pairs, warning = build_pairs(
        data.get("machines", ""),
        data.get("emails", ""),
        data.get("passwords", ""),
        data.get("recovery_emails", ""),
        data.get("recovery_passwords", ""),
        mode=data.get("mode", "fill_only"),
    )
    if not pairs and warning:
        return jsonify({"ok": False, "error": warning}), 400
    return jsonify({"ok": True, "pairs": pairs, "warning": warning})


@app.get("/api/settings")
def get_settings():
    return jsonify({"ok": True, "settings": public_settings()})


@app.post("/api/settings")
def save_settings():
    data = request.get_json(force=True) or {}
    config = load_config()
    if "hubstudio_exe_path" in data:
        config["hubstudio_exe_path"] = (data.get("hubstudio_exe_path") or "").strip()
    if "hubstudio_api" in data:
        config["hubstudio_api"] = (data.get("hubstudio_api") or "").strip() or config.get(
            "hubstudio_api", "http://127.0.0.1:6873/api/v1"
        )
    if "local_api_key" in data:
        config["local_api_key"] = (data.get("local_api_key") or "").strip()
    if "excel_file_path" in data:
        excel = config.setdefault("excel_sync", {})
        excel["file_path"] = (data.get("excel_file_path") or "").strip()
        excel.setdefault("enabled", True)
        excel.setdefault("write_while_open", True)
    save_config(config)
    return jsonify({"ok": True, "settings": public_settings(config)})


@app.post("/api/settings/test-hubstudio")
def test_hubstudio_settings():
    config = load_config()
    data = request.get_json(force=True) or {}
    exe_path = (data.get("hubstudio_exe_path") or config.get("hubstudio_exe_path") or "").strip()
    api_base = (data.get("hubstudio_api") or config.get("hubstudio_api") or "").strip()
    api_key = (data.get("local_api_key") or config.get("local_api_key") or "").strip()
    ok, detail = ensure_hubstudio_ready(api_base, exe_path, api_key)
    return jsonify({"ok": ok, "detail": detail, "settings": public_settings(config)})


@app.get("/api/excel/sheets")
def excel_list_sheets():
    config = load_config()
    excel_cfg = excel_sync_config(config)
    file_path = excel_cfg.get("file_path") or ""
    if not file_path:
        return jsonify({"ok": False, "error": "未配置 excel_sync.file_path"}), 400
    names, err = list_sheet_names(file_path)
    if err:
        return jsonify({"ok": False, "error": err}), 400
    return jsonify({"ok": True, "sheets": names, "file_path": file_path})


@app.get("/api/excel/config")
def excel_get_config():
    config = load_config()
    excel_cfg = excel_sync_config(config)
    return jsonify({"ok": True, "excel_sync": excel_cfg})


@app.post("/api/batch/start")
def batch_start():
    data = request.get_json(force=True) or {}
    mode = data.get("mode", "fill_only")
    if mode not in {"open_only", "fill_only", "full_check"}:
        return jsonify({"ok": False, "error": "无效模式"}), 400

    concurrency = int(data.get("concurrency") or load_config().get("batch_concurrency", 6))
    concurrency = max(1, min(concurrency, 12))

    pairs, warning = build_pairs(
        data.get("machines", ""),
        data.get("emails", ""),
        data.get("passwords", ""),
        data.get("recovery_emails", ""),
        data.get("recovery_passwords", ""),
        mode=data.get("mode", "fill_only"),
    )
    if not pairs:
        return jsonify({"ok": False, "error": warning or "没有可执行任务"}), 400

    recovery_pool = build_recovery_pool(
        data.get("recovery_emails", ""),
        data.get("recovery_passwords", ""),
    )
    recovery_assign_ordered = bool(data.get("recovery_assign_ordered", True))
    if recovery_pool and recovery_assign_ordered:
        warning = (
            (warning + "；") if warning else ""
        ) + "辅助邮箱按机子列表顺序领取（池内第1条给最先需要绑定的靠前机子）"

    config = load_config()
    excel_cfg = excel_sync_config(config)
    excel_sheet = (data.get("excel_sheet") or "").strip()
    # 暂时关闭 Excel 回写：即使填了工作表名也不写入（config.excel_sync.enabled=false）
    excel_enabled = bool(
        excel_cfg.get("enabled")
        and data.get("excel_sync", True)
        and excel_sheet
    )
    if excel_enabled and not excel_cfg.get("file_path"):
        return jsonify(
            {
                "ok": False,
                "error": "请先在「环境设置」中填写 Excel 总表路径，或留空工作表名跳过 Excel 回写",
            }
        ), 400

    api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = config.get("local_api_key", "")
    exe_path = config.get("hubstudio_exe_path", "")
    hs_ok, hs_detail = ensure_hubstudio_ready(api_base, exe_path, api_key)
    if not hs_ok:
        return jsonify({"ok": False, "error": hs_detail}), 400

    job_id = uuid.uuid4().hex[:12]
    excel_sync_job = None
    if excel_enabled:
        excel_sync_job = {
            "enabled": True,
            "file_path": excel_cfg["file_path"],
            "sheet_name": excel_sheet,
            "machine_headers": excel_cfg["machine_headers"],
            "status_headers": excel_cfg["status_headers"],
            "include_detail": excel_cfg["include_detail"],
            "write_while_open": excel_cfg["write_while_open"],
            "updated": 0,
            "skipped": 0,
            "errors": [],
        }

    with _jobs_lock:
        _jobs[job_id] = {
            "id": job_id,
            "status": "running",
            "mode": mode,
            "concurrency": concurrency,
            "total": len(pairs),
            "completed": 0,
            "running_count": 0,
            "running_machines": [],
            "running_steps": {},
            "activity_log": [],
            "pairs": pairs,
            "recovery_pool": recovery_pool,
            "recovery_claimed": [],
            "recovery_assign_ordered": recovery_assign_ordered,
            "status_overrides": {},
            "current_machine": "",
            "results": [],
            "warning": warning,
            "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "finished_at": "",
            "result_file": "",
            "cancelled": False,
            "excel_sync": excel_sync_job,
        }
    clear_foxmail_cancel()

    thread = threading.Thread(
        target=run_batch_job,
        args=(job_id, pairs, mode, concurrency),
        daemon=True,
    )
    thread.start()
    return jsonify({
        "ok": True,
        "job_id": job_id,
        "warning": warning,
        "total": len(pairs),
        "concurrency": concurrency,
        "pairs": [_sanitize_pair(p) for p in pairs],
    })


@app.get("/api/batch/status/<job_id>")
def batch_status(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        return jsonify({"ok": False, "error": "任务不存在"}), 404
    return jsonify({"ok": True, "job": _sanitize_job(job)})


@app.post("/api/browser/open")
def browser_open_one():
    try:
        data = request.get_json(silent=True) or {}
        machine_id = (data.get("machine_id") or "").strip()
        if not machine_id:
            return jsonify({"ok": False, "error": "缺少机子号"}), 400
        config = load_config()
        apply_runtime_config(config)
        api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
        api_key = config.get("local_api_key", "")
        login_url = config["login_url"]
        ok, container_code, detail = open_browser_for_debug(
            api_base, machine_id, login_url, api_key
        )
        return jsonify(
            {
                "ok": ok,
                "machine_id": machine_id,
                "container_code": container_code,
                "detail": detail,
            }
        )
    except Exception as exc:
        # 环境可能已弹出，始终返回 JSON，避免前端 JSON.parse 报 doctype
        return jsonify(
            {
                "ok": False,
                "error": f"打开异常（若窗口已弹出可忽略）: {exc}",
                "detail": f"打开异常（若窗口已弹出可忽略）: {exc}",
            }
        ), 500


@app.post("/api/recovery/submit-code")
def recovery_submit_code():
    """人工输入验证码 → 回填到对应机子浏览器验证码页 → 自动完成登录。"""
    data = request.get_json(force=True) or {}
    machine_id = (data.get("machine_id") or "").strip()
    code = (data.get("code") or "").strip()
    job_id = (data.get("job_id") or "").strip()
    if not machine_id:
        return jsonify({"ok": False, "error": "缺少机子号"}), 400
    if not code:
        return jsonify({"ok": False, "error": "请输入验证码"}), 400

    config = load_config()
    apply_runtime_config(config)
    api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = config.get("local_api_key", "")
    login_url = config.get("login_url") or ""

    pending = get_pending_recovery_code(machine_id) or {}
    recovery_email = (pending.get("recovery_email") or "").strip()

    def _status_from_step(step: str) -> str:
        s = step or ""
        if "封号" in s or "AMZ" in s:
            return "检查封号中"
        if any(
            x in s
            for x in (
                "收件箱",
                "保持登录",
                "通行密钥",
                "登录收尾",
                "同意",
                "账户选择",
                "打开 Outlook",
            )
        ):
            return "登录收尾中"
        return "验证码处理中"

    def _patch_job_progress(step: str, *, status_text: str | None = None) -> None:
        if not job_id:
            return
        label = status_text or _status_from_step(step)
        with _jobs_lock:
            job = _jobs.get(job_id)
            if not job:
                return
            ts = datetime.now().strftime("%H:%M:%S")
            job.setdefault("running_steps", {})[machine_id] = {"step": step, "at": ts}
            log = job.setdefault("activity_log", [])
            log.append({"at": ts, "machine_id": machine_id, "step": step})
            if len(log) > 100:
                job["activity_log"] = log[-100:]
            for row in job.get("results") or []:
                if row.get("machine_id") == machine_id:
                    row["status"] = label
                    row["detail"] = step
                    row["awaiting_code"] = False
                    if recovery_email and not row.get("recovery_email"):
                        row["recovery_email"] = recovery_email
                    break

    _patch_job_progress("正在连接浏览器并回填验证码…")

    ok, detail, status = apply_manual_recovery_code(
        api_base,
        api_key,
        login_url,
        machine_id,
        code,
        on_progress=_patch_job_progress,
    )

    # 登入 / AMZ账号被封 / 其它状态直接用枚举文案
    if ok and status == AccountStatus.OK:
        status_text = AccountStatus.LOGIN_OK.value
    else:
        status_text = status.value

    # 更新当前任务结果行
    if job_id:
        with _jobs_lock:
            job = _jobs.get(job_id)
            if job:
                for row in job.get("results") or []:
                    if row.get("machine_id") == machine_id:
                        row["status"] = status_text
                        row["detail"] = detail
                        row["awaiting_code"] = status == AccountStatus.WAIT_CODE
                        # 终态截图路径写入字段，便于前端只展示最新一张
                        m = re.search(r"截图:\s*(\S+)", detail or "")
                        if m:
                            row["screenshot_path"] = m.group(1)
                        if recovery_email and not row.get("recovery_email"):
                            row["recovery_email"] = recovery_email
                        break
                ts = datetime.now().strftime("%H:%M:%S")
                # 结束后清掉进行中步骤，避免 UI 仍显示「验证码处理中」
                job.setdefault("running_steps", {}).pop(machine_id, None)
                log = job.setdefault("activity_log", [])
                log.append({"at": ts, "machine_id": machine_id, "step": f"验证码提交: {detail}"})
                if len(log) > 100:
                    job["activity_log"] = log[-100:]

    return jsonify(
        {
            "ok": ok,
            "machine_id": machine_id,
            "status": status_text,
            "detail": detail,
            "awaiting_code": status == AccountStatus.WAIT_CODE,
            "recovery_email": recovery_email,
        }
    )


@app.post("/api/result/mark-status")
def mark_result_status():
    """人工标记某机子状态（如辅助邮箱密码错误）。"""
    data = request.get_json(force=True) or {}
    job_id = (data.get("job_id") or "").strip()
    machine_id = (data.get("machine_id") or "").strip()
    status_text = (data.get("status") or "").strip()
    detail = (data.get("detail") or "").strip() or f"人工标记：{status_text}"
    if not machine_id:
        return jsonify({"ok": False, "error": "缺少机子号"}), 400
    if not status_text:
        return jsonify({"ok": False, "error": "缺少状态"}), 400

    updated = None
    with _jobs_lock:
        job = _jobs.get(job_id) if job_id else None
        if not job and job_id:
            return jsonify({"ok": False, "error": "任务不存在"}), 404
        if job:
            job.setdefault("status_overrides", {})[machine_id] = {
                "status": status_text,
                "detail": detail,
            }
            found = False
            for row in job.get("results") or []:
                if row.get("machine_id") == machine_id:
                    row["status"] = status_text
                    row["detail"] = detail
                    row["awaiting_code"] = False
                    updated = dict(row)
                    found = True
                    break
            if not found:
                pair = next(
                    (p for p in (job.get("pairs") or []) if p.get("machine_id") == machine_id),
                    None,
                )
                updated = {
                    "index": (pair or {}).get("index", 0),
                    "machine_id": machine_id,
                    "email": (pair or {}).get("email", ""),
                    "status": status_text,
                    "detail": detail,
                    "awaiting_code": False,
                    "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                job.setdefault("results", []).append(updated)
                job["results"].sort(key=lambda x: x.get("index", 0))
                job["completed"] = len(job["results"])

    if job_id and updated is not None:
        _append_activity(job_id, machine_id, f"人工标记: {status_text}")
        threading.Thread(
            target=_maybe_sync_excel,
            args=(job_id, updated),
            daemon=True,
        ).start()

    return jsonify({"ok": True, "machine_id": machine_id, "status": status_text, "detail": detail})


@app.post("/api/cleanup/csv")
def cleanup_csv_api():
    config = load_config()
    keep = int(config.get("results_keep_count", 5))
    removed, kept = cleanup_old_results(BASE_DIR, keep=keep)
    return jsonify(
        {
            "ok": True,
            "removed": removed,
            "kept": kept,
            "detail": f"已删除 {removed} 个旧 CSV，保留最近 {kept} 份",
        }
    )


@app.post("/api/cleanup/screenshots")
def cleanup_screenshots_api():
    removed, freed = cleanup_screenshots(BASE_DIR)
    return jsonify(
        {
            "ok": True,
            "removed": removed,
            "freed_bytes": freed,
            "detail": f"已删除 {removed} 张截图，释放 {_format_bytes(freed)}",
        }
    )


@app.post("/api/open-imgs-folder")
def open_imgs_folder():
    """在资源管理器中打开截图文件夹（可指定 imgs/时间戳 子目录）。"""
    import os

    data = request.get_json(force=True) or {}
    rel = (data.get("folder") or data.get("path") or "imgs").strip().replace("\\", "/")
    if ".." in rel or rel.startswith("/") or (len(rel) > 1 and rel[1] == ":"):
        return jsonify({"ok": False, "error": "非法路径"}), 400
    if not rel.startswith("imgs"):
        rel = "imgs/" + rel.lstrip("/")
    # 若传入文件路径，取父目录
    if rel.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".gif")):
        rel = "/".join(rel.split("/")[:-1]) or "imgs"
    target = (BASE_DIR / rel).resolve()
    imgs_root = (BASE_DIR / "imgs").resolve()
    try:
        target.relative_to(imgs_root)
    except ValueError:
        if target != imgs_root:
            return jsonify({"ok": False, "error": "只能打开 imgs 目录"}), 400
    if not target.exists():
        imgs_root.mkdir(parents=True, exist_ok=True)
        target = imgs_root
    try:
        if os.name == "nt":
            os.startfile(str(target))  # type: ignore[attr-defined]
        else:
            import sys
            import subprocess

            if sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
    except Exception as exc:
        return jsonify({"ok": False, "error": f"打开失败: {exc}"}), 500
    return jsonify({"ok": True, "detail": f"已打开: {target}", "path": str(target)})


@app.post("/api/browser/stop-all")
def browser_stop_all():
    config = load_config()
    api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = config.get("local_api_key", "")
    ok, detail = stop_all_browsers(api_base, api_key, clear_opening=True)
    return jsonify({"ok": ok, "detail": detail})


@app.post("/api/browser/stop")
def browser_stop_machines():
    data = request.get_json(force=True) or {}
    machines = parse_lines(data.get("machines", ""))
    if not machines:
        return jsonify({"ok": False, "error": "请至少输入一个机子号"}), 400
    config = load_config()
    api_base = config.get("hubstudio_api", "http://127.0.0.1:6873/api/v1")
    api_key = config.get("local_api_key", "")
    results = stop_browsers_by_machine_ids(api_base, machines, api_key)
    ok_count = sum(1 for r in results if r.get("ok"))
    return jsonify(
        {
            "ok": ok_count > 0,
            "detail": f"已关闭 {ok_count}/{len(results)} 个环境",
            "results": results,
        }
    )


@app.post("/api/batch/cancel/<job_id>")
def batch_cancel(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
        if job:
            job["cancelled"] = True
            job["status"] = "cancelled"
    request_foxmail_cancel()
    _close_running_job_browsers(job_id)
    return jsonify({"ok": True, "detail": "已请求停止，正在关闭环境并恢复 Foxmail"})


if __name__ == "__main__":
    from run_app import main

    main()
