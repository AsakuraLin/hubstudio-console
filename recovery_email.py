# -*- coding: utf-8 -*-
"""Microsoft 辅助邮箱绑定 + IMAP 收取验证码。"""

from __future__ import annotations

import email
import imaplib
import re
import time
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from email.utils import parsedate_to_datetime
from typing import Any, Callable

from playwright.sync_api import Page

from foxmail_automation import create_other_mailbox_account

LogFn = Callable[[str], None]

RECOVERY_PAGE_PATTERNS = [
    "add a way to verify",
    "add another way",
    "alternate email",
    "alternative email",
    "backup email",
    "recovery email",
    "security info",
    "help us protect your account",
    "let's protect your account",
    "添加安全信息",
    "备用电子邮件",
    "备用邮箱",
    "辅助电子邮件",
    "辅助邮箱",
    "別のメール",
    "代替のメール",
    "代替メール",
    "セキュリティ情報",
    "アカウントを保護",
    "アカウントを保護しましょう",
    "someone@example.com",
    "メールアドレスを追加",
    "追加のセキュリティ",
    "proofs",
]

RECOVERY_EMAIL_SELECTORS = [
    "#EmailAddress",
    "#iProofEmail",
    "#ProofEmail",
    "#iEmail",
    'input[name="EmailAddress"]',
    'input[name="email"]',
    'input[name="ProofEmail"]',
    'input[name="iProofEmail"]',
    'input[type="email"]',
    'input[aria-label*="mail" i]',
    'input[aria-label*="メール" i]',
    'input[aria-label*="電子メール" i]',
    'input[placeholder*="example" i]',
    'input[placeholder="someone@example.com"]',
    'input[placeholder*="@"]',
    'input[data-testid*="email" i]',
]

CODE_PAGE_PATTERNS = [
    # 仅用「明确在输入验证码」的文案；勿用单独「セキュリティコード」
    # （绑定页会写「セキュリティコードが送信されます」导致误判）
    "enter the code",
    "enter code",
    "verification code",
    "one-time code",
    "输入验证码",
    "输入代码",
    "コードを入力",
    "コードの入力",
    "にお送りしたコード",
    "確認コードを入力",
    "コードを入力してください",
]

# 电话认证（确认手机号末 4 位）——绝不是辅助邮箱验证码页
PHONE_VERIFY_PATTERNS = [
    "電話番号を確認する",
    "電話番号を確認",
    "欠落している番号の最後の 4 桁",
    "最後の 4 桁の数字を入力",
    "番号の最後の 4 桁",
    "Confirm your phone number",
    "confirm your phone number",
    "Enter the last 4 digits",
    "enter the last 4 digits",
    "last 4 digits of the phone",
    "last four digits of your phone",
    "We will send a code to",
    "确认电话号码",
    "验证你的电话号码",
    "验证电话号码",
    "输入电话号码的最后",
    "请输入电话号码的最后",
]

CODE_INPUT_SELECTORS = [
    "#iOttText",
    "#idTxtBx_SAOTCC_OTC",
    'input[name="otc"]',
    'input[name="code"]',
    'input[autocomplete="one-time-code"]',
    'input[aria-label*="コード" i]',
    'input[aria-label*="code" i]',
    'input[placeholder="コード"]',
    'input[placeholder*="Code" i]',
]

SUBMIT_SELECTORS = [
    "#iNext",
    "#idSIButton9",
    "#iBtn_action",
    'input[type="submit"]',
    'button[type="submit"]',
    'button:has-text("Send code")',
    'button:has-text("发送")',
    'button:has-text("送信")',
    'button:has-text("次へ")',
    'input[value="次へ"]',
    'input[value="Next"]',
    'input[value="Send"]',
    'button:has-text("Next")',
    'button[data-testid*="primary" i]',
]

DEFAULT_IMAP_HOSTS = {
    "qq.com": ("imap.qq.com", 993),
    "foxmail.com": ("imap.qq.com", 993),
    "163.com": ("imap.163.com", 993),
    "126.com": ("imap.126.com", 993),
    "gmail.com": ("imap.gmail.com", 993),
    "outlook.com": ("outlook.office365.com", 993),
    "hotmail.com": ("outlook.office365.com", 993),
    "live.com": ("outlook.office365.com", 993),
    "aol.com": ("imap.aol.com", 993),
}


def _noop_log(_: str) -> None:
    pass


_KEEP_BACKGROUND = True


def set_keep_background(enabled: bool) -> None:
    global _KEEP_BACKGROUND
    _KEEP_BACKGROUND = enabled


def maybe_bring_to_front(page: Page) -> None:
    if _KEEP_BACKGROUND:
        return
    try:
        page.bring_to_front()
    except Exception:
        pass


def page_text(page: Page) -> str:
    """读取页面可见文案；Microsoft 登录常把表单放在 iframe，需合并各 frame。"""
    parts: list[str] = []
    try:
        t = page.inner_text("body")
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
                t = frame.inner_text("body")
                if t and str(t).strip():
                    parts.append(str(t))
            except Exception:
                continue
    except Exception:
        pass
    return "\n".join(parts)


def page_text_lower(page: Page) -> str:
    return page_text(page).lower()


def safe_wait(page: Page, ms: int) -> None:
    try:
        page.wait_for_timeout(ms)
    except Exception:
        pass


def wait_and_click(page: Page, selectors: list[str], timeout_ms: int = 1500) -> bool:
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=timeout_ms)
            loc.click()
            return True
        except Exception:
            continue
    return False


def fill_first_visible(page: Page, selectors: list[str], value: str) -> bool:
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if not loc.is_visible(timeout=800):
                continue
            loc.click(timeout=500)
            loc.fill(value, timeout=1500)
            if loc.input_value(timeout=400).strip() == value.strip():
                return True
        except Exception:
            continue
    return False


def find_code_input(page: Page):
    # 电话认证末 4 位输入框不是邮件验证码框
    try:
        if is_phone_verify_page(page):
            return None
    except Exception:
        pass
    for sel in CODE_INPUT_SELECTORS:
        for ctx in _iter_contexts(page):
            loc = ctx.locator(sel).first
            try:
                if not (loc.count() and loc.is_visible(timeout=300)):
                    continue
                typ = (loc.get_attribute("type") or "").lower()
                ph = (loc.get_attribute("placeholder") or "").lower()
                name = (loc.get_attribute("name") or "").lower()
                # 排除邮箱框（绑定页 placeholder=someone@example.com）
                if typ in {"email", "password", "hidden", "checkbox", "radio"}:
                    continue
                if "example.com" in ph or "@" in ph or "mail" in name or "email" in name:
                    continue
                return loc
            except Exception:
                continue
    # 宽松兜底：数字验证码框（排除邮箱）
    for ctx in _iter_contexts(page):
        try:
            for loc in ctx.locator(
                'input[type="tel"], input[inputmode="numeric"], input[type="text"]'
            ).all()[:6]:
                if not loc.is_visible(timeout=150):
                    continue
                typ = (loc.get_attribute("type") or "").lower()
                ph = (loc.get_attribute("placeholder") or "").lower()
                name = (loc.get_attribute("name") or "").lower()
                aria = (loc.get_attribute("aria-label") or "").lower()
                if typ == "email" or "example.com" in ph or "@" in ph:
                    continue
                blob = f"{ph} {name} {aria}"
                if any(
                    k in blob
                    for k in ("code", "otc", "コード", "验证码", "確認")
                ):
                    return loc
        except Exception:
            continue
    return None


def _is_still_recovery_email_page(page: Page) -> bool:
    """仍在「填辅助邮箱」页（アカウントを保護 / someone@example.com）。"""
    try:
        text = page_text(page)
    except Exception:
        text = ""
    text_lower = text.lower()
    if find_recovery_email_input(page):
        if any(
            p in text or p in text_lower
            for p in (
                "アカウントを保護",
                "protect your account",
                "someone@example.com",
                "連絡用メール",
                "alternate email",
                "recovery email",
            )
        ):
            return True
        # 有邮箱输入框且没有明确的验证码标题 → 仍算绑定页
        if not any(
            p in text or p in text_lower
            for p in ("コードの入力", "コードを入力", "enter the code", "にお送りしたコード")
        ):
            return True
    return False


def is_phone_verify_page(page: Page) -> bool:
    """
    Microsoft「電話番号を確認する」：要填手机号末 4 位才能发短信。
    绝不能当成辅助邮箱验证码页去领池。
    """
    try:
        text = page_text(page) or ""
    except Exception:
        return False
    text_lower = text.lower()
    if any(p in text or p in text_lower for p in PHONE_VERIFY_PATTERNS):
        return True
    # 掩码号 + 末四位
    if (
        ("電話" in text or "phone number" in text_lower or "phone" in text_lower)
        and (
            "最後の 4 桁" in text
            or "最後の4桁" in text
            or "last 4 digit" in text_lower
            or "last four digit" in text_lower
            or "最后4位" in text
            or "最后 4 位" in text
        )
    ):
        return True
    # 「********60 にコードを送信」类短信预确认
    if "にコードを送信" in text and ("欠落" in text or "最後の" in text or "電話" in text):
        return True
    if "send a code to" in text_lower and (
        "last 4" in text_lower or "missing" in text_lower or "****" in text
    ):
        return True
    return False


def is_code_verify_page(page: Page) -> bool:
    # 电话认证页有数字框，绝不能当成邮件验证码页
    if is_phone_verify_page(page):
        return False
    # 绑定辅助邮箱页文案含「セキュリティコードが送信されます」，绝不能当成验证码页
    if _is_still_recovery_email_page(page):
        return False

    if find_code_input(page):
        return True

    text = page_text(page)
    text_lower = text.lower()
    if any(p in text or p in text_lower for p in CODE_PAGE_PATTERNS):
        # 再挡一层：含「電話番号を確認」等时不算邮件码页
        if is_phone_verify_page(page):
            return False
        return True
    return False


def _iter_contexts(page: Page):
    yield page
    for frame in page.frames:
        try:
            if frame != page.main_frame:
                yield frame
        except Exception:
            continue


def find_recovery_email_input(page: Page):
    # 登录首页「メール、電話、Skype」绝不是辅助邮箱输入框
    try:
        home_text = page_text(page) or ""
        if any(
            m in home_text
            for m in (
                "メール、電話、Skype",
                "Email, phone, or Skype",
                "Email, phone, Skype",
            )
        ):
            # 除非同时出现真正的绑定页标题
            if "アカウントを保護" not in home_text and "someone@example.com" not in home_text.lower():
                return None
    except Exception:
        pass

    placeholder_needles = (
        "someone@example.com",
        "example.com",
    )
    login_field_deny = (
        "loginfmt",
        "i0116",
        "usernameentry",
        "メール、電話",
        "email, phone",
        "skype",
    )
    for ctx in _iter_contexts(page):
        try:
            loc = ctx.get_by_placeholder("someone@example.com").first
            if loc.is_visible(timeout=400):
                return loc
        except Exception:
            pass
        try:
            loc = ctx.get_by_label(re.compile(r"mail|メール|電子メール|email", re.I)).first
            if loc.is_visible(timeout=400):
                name = (loc.get_attribute("name") or "").lower()
                pid = (loc.get_attribute("id") or "").lower()
                if "loginfmt" in name or pid in {"i0116", "usernameentry"}:
                    pass
                else:
                    return loc
        except Exception:
            pass
        for sel in RECOVERY_EMAIL_SELECTORS:
            if any(x in sel.lower() for x in ("loginfmt", "i0116")):
                continue
            loc = ctx.locator(sel).first
            try:
                if loc.count() and loc.is_visible(timeout=400):
                    name = (loc.get_attribute("name") or "").lower()
                    pid = (loc.get_attribute("id") or "").lower()
                    ph = (loc.get_attribute("placeholder") or "").lower()
                    if any(d in name or d in pid or d in ph for d in login_field_deny):
                        continue
                    return loc
            except Exception:
                continue
        try:
            for loc in ctx.locator(
                'input[type="email"], input[type="text"], input:not([type]):not([type="password"]):not([type="hidden"]):not([type="checkbox"]):not([type="radio"])'
            ).all():
                if not loc.is_visible(timeout=150):
                    continue
                ph = (loc.get_attribute("placeholder") or "").lower()
                name = (loc.get_attribute("name") or "").lower()
                aria = (loc.get_attribute("aria-label") or "").lower()
                pid = (loc.get_attribute("id") or "").lower()
                blob = f"{ph} {name} {aria} {pid}"
                if any(d in blob for d in login_field_deny):
                    continue
                if any(n.lower() in blob for n in placeholder_needles) or "someone@example" in ph:
                    return loc
        except Exception:
            continue

    text = page_text(page)
    if any(
        p in text
        for p in [
            "アカウントを保護",
            "protect your account",
            "someone@example.com",
            "代替メール",
        ]
    ):
        for ctx in _iter_contexts(page):
            for sel in (
                '#EmailAddress',
                '#iProofEmail',
                'input[name="EmailAddress"]',
                'input[name="ProofEmail"]',
                'input[placeholder="someone@example.com"]',
            ):
                loc = ctx.locator(sel).first
                try:
                    if loc.count() and loc.is_visible(timeout=500):
                        return loc
                except Exception:
                    continue
    return None


def _read_locator_value(loc, *, timeout_ms: int = 1200) -> str:
    """读取输入框值；后台/静默模式 input_value 常超时，用 evaluate 兜底。"""
    for _ in range(3):
        try:
            v = loc.input_value(timeout=timeout_ms)
            if (v or "").strip():
                return v.strip()
        except Exception:
            pass
        try:
            v = loc.evaluate("el => (el && (el.value ?? el.textContent)) || ''")
            if str(v or "").strip():
                return str(v).strip()
        except Exception:
            pass
        time.sleep(0.15)
    return ""


def _input_matches(loc, value: str) -> bool:
    target = (value or "").strip()
    if not target:
        return False
    cur = _read_locator_value(loc)
    return cur == target


def _recovery_flow_advanced(page: Page, before_url: str = "") -> bool:
    """
    辅助邮箱已提交并离开绑定输入页（即使未到验证码页也算前进）。
    避免「其实已填完在跳转，脚本却报超时失败」。
    """
    try:
        if is_code_verify_page(page):
            return True
    except Exception:
        pass
    try:
        if is_phone_verify_page(page):
            return True
    except Exception:
        pass
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
        return True
    if before_url and (page.url or "") != before_url:
        if not _is_still_recovery_email_page(page):
            return True
        try:
            if not is_recovery_bind_page(page):
                return True
        except Exception:
            return True
    try:
        if not _is_still_recovery_email_page(page) and not is_recovery_bind_page(page):
            return True
    except Exception:
        pass
    text = page_text(page).lower()
    if any(
        x in text
        for x in (
            "stay signed in",
            "keep me signed in",
            "保持登录",
            "サインインしたまま",
            "サインインのまま",
        )
    ):
        return True
    return False


def fill_input_value(page: Page, loc, value: str) -> bool:
    try:
        loc.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass

    def _ok() -> bool:
        return _input_matches(loc, value)

    # 0) 先 focus + JS 赋值（HubStudio 后台/不可见时 fill 常失败）
    try:
        loc.evaluate(
            """(el, v) => {
                el.focus();
                el.click();
                const proto = window.HTMLInputElement.prototype;
                const desc = Object.getOwnPropertyDescriptor(proto, 'value');
                if (desc && desc.set) desc.set.call(el, '');
                else el.value = '';
                el.dispatchEvent(new Event('input', { bubbles: true }));
                if (desc && desc.set) desc.set.call(el, v);
                else el.value = v;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
            }""",
            value,
        )
        safe_wait(page, 200)
        if _ok():
            return True
    except Exception:
        pass

    # 1) 常规 fill
    try:
        loc.click(timeout=2000, force=True)
        loc.fill("")
        loc.fill(value, timeout=4000, force=True)
        safe_wait(page, 150)
        if _ok():
            return True
    except Exception:
        pass

    # 2) 清空后逐字输入
    try:
        loc.click(timeout=1500, force=True)
        loc.press("Control+a")
        loc.press("Backspace")
        loc.press_sequential(value, delay=15)
        safe_wait(page, 150)
        if _ok():
            return True
    except Exception:
        pass

    # 3) 键盘 type
    try:
        loc.click(timeout=1200, force=True)
        page.keyboard.press("Control+a")
        page.keyboard.type(value, delay=20)
        safe_wait(page, 150)
        if _ok():
            return True
    except Exception:
        pass

    # 4) 读回失败但 DOM 里已有值（React 受控组件常见误报）
    return _ok()


def _click_next_button(page: Page) -> bool:
    if wait_and_click(page, SUBMIT_SELECTORS, timeout_ms=2200):
        return True
    for label in ["次へ", "Next", "发送", "Send", "送信", "继续", "Continue"]:
        try:
            page.get_by_role("button", name=label).first.click(timeout=1500)
            return True
        except Exception:
            pass
        try:
            page.locator(f'input[type="submit"][value="{label}"]').first.click(timeout=1200)
            return True
        except Exception:
            pass
        try:
            page.get_by_text(label, exact=True).first.click(timeout=1200)
            return True
        except Exception:
            pass
    for ctx in _iter_contexts(page):
        try:
            btn = ctx.locator("#iNext, input[type='submit'], button[type='submit']").first
            if btn.is_visible(timeout=500):
                btn.click(timeout=1500)
                return True
        except Exception:
            continue
    try:
        page.keyboard.press("Enter")
        return True
    except Exception:
        return False


def submit_recovery_email(page: Page, recovery_email: str, log: LogFn = _noop_log) -> bool:
    try:
        maybe_bring_to_front(page)
    except Exception:
        pass
    safe_wait(page, 500)

    if is_code_verify_page(page):
        log("已在验证码页，跳过填辅助邮箱")
        return True

    if not wait_for_recovery_input(page, timeout_sec=22):
        if _recovery_flow_advanced(page):
            log("输入框未读到但页面已前进，视为已提交")
            return True
        log("未找到辅助邮箱输入框")
        return False

    loc = find_recovery_email_input(page)
    if loc is None:
        if _recovery_flow_advanced(page):
            return True
        log("辅助邮箱输入框定位失败")
        return False

    before_url = ""
    try:
        before_url = page.url or ""
    except Exception:
        pass

    log(f"填写辅助邮箱: {recovery_email}")
    filled = fill_input_value(page, loc, recovery_email)
    if not filled:
        safe_wait(page, 800)
        loc = find_recovery_email_input(page)
        if loc is not None:
            filled = fill_input_value(page, loc, recovery_email)
    if not filled:
        # 值可能已写入 DOM 但读回超时；有值仍尝试点下一步
        loc = find_recovery_email_input(page)
        if loc is not None and _input_matches(loc, recovery_email):
            filled = True
            log("辅助邮箱读回超时但 DOM 已有值，继续提交")
        else:
            log("辅助邮箱填写失败")
            return False

    if not _click_next_button(page):
        if _recovery_flow_advanced(page, before_url):
            log("未能点到次へ但页面已前进")
            return True
        log("未能点击次へ/Next")
        return False

    safe_wait(page, 2000)
    if _recovery_flow_advanced(page, before_url):
        return True
    if is_code_verify_page(page):
        return True

    # 若仍停在同一输入框且值被清空，视为提交失败
    try:
        still = find_recovery_email_input(page)
        if still and is_recovery_bind_page(page) and not is_code_verify_page(page):
            cur = _read_locator_value(still)
            if page.url == before_url:
                _click_next_button(page)
                safe_wait(page, 1800)
            if is_code_verify_page(page) or _recovery_flow_advanced(page, before_url):
                return True
            if cur == recovery_email.strip() or cur == "":
                return True
    except Exception:
        pass
    return True


IDENTITY_VERIFY_PATTERNS = [
    "お客様のアカウントのセキュリティ保護にご協力ください",
    "異常なアクティビティが検出された",
    "異常なアクティビティ",
    "本人確認とパスワードの変更",
    "資格情報が危険にさらされている",
    "オンラインで確認",
    "unusual activity",
    "credentials may have been compromised",
    "verify your identity and change your password",
    "verify online",
]


def is_identity_verification_page(page: Page) -> bool:
    """异常活动 / 本人确认页，不是辅助邮箱绑定。"""
    text = page_text(page)
    text_lower = text.lower()
    if any(p in text or p in text_lower for p in IDENTITY_VERIFY_PATTERNS):
        return True
    for label in ("オンラインで確認", "Verify online", "在线验证"):
        try:
            if page.get_by_text(label, exact=False).first.is_visible(timeout=400):
                return True
        except Exception:
            continue
    return False


def is_ms_auth_page_shell(page: Page) -> bool:
    """
    Microsoft 登录/验证半加载壳：仅邮箱+Logo，密码/绑定表单尚未渲染。
    此状态绝不能判「需要绑定辅助邮箱」或「未知状态」并结束。
    """
    # 绑定表单常在 iframe：主框看起来像壳，但已可填辅助邮箱 → 绝不是半加载
    try:
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
                "メールの追加",
            )
        ):
            return False
    except Exception:
        pass
    try:
        return bool(
            page.evaluate(
                """() => {
                    const url = (location.href || '').toLowerCase();
                    if (!['login.microsoftonline.com','login.live.com','account.live.com',
                          'signup.live.com','account.microsoft.com']
                        .some(h => url.includes(h))) return false;
                    const t = ((document.body && document.body.innerText) || '')
                        .replace(/\\s+/g, ' ').trim();
                    const settled = [
                        'アカウントを保護', 'protect your account', 'someone@example.com',
                        'パスワード', 'password', 'コード', 'verify',
                    ];
                    if (settled.some(m => t.toLowerCase().includes(m.toLowerCase()))) {
                        return false;
                    }
                    const hasSubmit = !!document.querySelector(
                        '#idSIButton9, input[type="submit"]:not([disabled]), button[type="submit"]:not([disabled])'
                    );
                    const hasBanner = !!document.querySelector(
                        '[data-testid="identityBanner"], #displayName, #identityBanner, #bannerText'
                    );
                    const hasEmailInText = /@\\S+\\.\\S+/.test(t);
                    const visibleInputs = Array.from(
                        document.querySelectorAll('input:not([type=hidden])')
                    ).filter(el => {
                        try {
                            const r = el.getBoundingClientRect();
                            return r.width > 0 && r.height > 0 && el.offsetParent !== null;
                        } catch (e) { return false; }
                    });
                    const hasPwd = !!document.querySelector(
                        '#i0118, #passwordEntry, input[name="passwd"], input[type="password"]'
                    );
                    if ((hasBanner || hasEmailInText) && !hasSubmit && !hasPwd
                        && visibleInputs.length === 0 && t.length < 320) {
                        return true;
                    }
                    if (/microsoft/i.test(t) && t.length < 140 && !hasSubmit && !hasPwd
                        && visibleInputs.length === 0) {
                        return true;
                    }
                    return false;
                }"""
            )
        )
    except Exception:
        return False


def is_recovery_bind_page(page: Page) -> bool:
    if is_ms_auth_page_shell(page):
        return False
    if is_identity_verification_page(page):
        return False
    text = page_text(page)
    text_lower = text.lower()
    url = page.url.lower()

    # 登录邮箱首页（メール、電話、Skype）绝不是辅助邮箱绑定页
    if any(
        m in text
        for m in (
            "メール、電話、Skype",
            "Email, phone, or Skype",
            "Email, phone, Skype",
        )
    ):
        if "アカウントを保護" not in text and "アカウントの保護" not in text and "someone@example.com" not in text_lower:
            return False

    # 已绑辅助邮箱的「确认邮箱/发送代码」页，绝不是去新绑辅助邮箱
    if (
        "メールをご確認ください" in text
        or "コードの送信" in text
        or "ご自身のものであることを確認" in text
        or "please check your email" in text_lower
        or "verify your email" in text_lower
        or re.search(r"[A-Za-z0-9._%+-]*\*+[A-Za-z0-9._%+-]*@", text)
    ):
        if "アカウントを保護しましょう" not in text and "アカウントの保護" not in text and "someone@example.com" not in text_lower:
            return False

    recovery_title_patterns = [
        "アカウントを保護しましょう",
        "アカウントを保護",
        "アカウントの保護にご協力ください",
        "アカウントの保護にご協力",
        "アカウントの保護",
        "help us protect your account",
        "let's protect your account",
        "メールの追加",
        "add email",
        "add an email",
    ]
    if any(p in text or p in text_lower for p in recovery_title_patterns):
        return True

    # 图三：只有「メールの追加」按钮的保护账户页
    for label in ("メールの追加", "Add email", "Add an email", "添加电子邮件", "添加邮箱"):
        try:
            if page.get_by_role("button", name=re.compile(label, re.I)).first.is_visible(timeout=250):
                return True
        except Exception:
            pass
        try:
            if page.get_by_text(label, exact=False).first.is_visible(timeout=250):
                if "protect" in text_lower or "保護" in text or "メール" in text:
                    return True
        except Exception:
            continue

    if "someone@example.com" in text and find_recovery_email_input(page):
        return True

    if any(x in url for x in ["proofs", "recover", "securityinfo", "proof", "addproof"]):
        if find_recovery_email_input(page):
            return True
        if any(p in text for p in recovery_title_patterns):
            return True
        return False

    if find_recovery_email_input(page) or find_code_input(page):
        # 勿用过宽词（如単独セキュリティ）误伤登录页
        if any(p in text_lower for p in RECOVERY_PAGE_PATTERNS):
            return True
        if any(
            p in text_lower
            for p in ["alternate email", "backup email", "备用", "辅助邮箱", "別のメール", "代替メール"]
        ):
            return True

    return False


def is_definitely_recovery_flow_page(page: Page) -> bool:
    """仅当可确认已进入绑定/验证码页时才允许领取辅助邮箱。"""
    try:
        if is_ms_auth_page_shell(page):
            return False
        if is_phone_verify_page(page):
            return False
    except Exception:
        pass
    try:
        # 登录邮箱首页绝不可领
        text = page_text(page) or ""
        if any(
            m in text
            for m in (
                "メール、電話、Skype",
                "Email, phone, or Skype",
                "Email, phone, Skype",
            )
        ):
            if "アカウントを保護" not in text and "アカウントの保護" not in text and "someone@example.com" not in text.lower():
                return False
    except Exception:
        pass
    try:
        if is_code_verify_page(page):
            return True
    except Exception:
        pass
    try:
        return bool(is_recovery_bind_page(page))
    except Exception:
        return False


def resolve_imap_settings(recovery_email: str, imap_config: dict[str, Any]) -> tuple[str, int, bool]:
    domain = recovery_email.split("@")[-1].lower()
    domain_hosts = imap_config.get("domain_hosts") or {}
    if domain in domain_hosts:
        host = domain_hosts[domain]
        port = int(imap_config.get("port", 993))
        use_ssl = bool(imap_config.get("use_ssl", True))
        return host, port, use_ssl

    host = (imap_config.get("host") or "").strip()
    if not host and domain in DEFAULT_IMAP_HOSTS:
        host, port = DEFAULT_IMAP_HOSTS[domain]
        return host, port, True
    if not host:
        host = f"imap.{domain}"
    port = int(imap_config.get("port", 993))
    use_ssl = bool(imap_config.get("use_ssl", True))
    return host, port, use_ssl


def _decode_mime_header(value: str) -> str:
    parts: list[str] = []
    for chunk, enc in decode_header(value):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(enc or "utf-8", errors="replace"))
        else:
            parts.append(str(chunk))
    return "".join(parts)


def _extract_code_from_text(text: str) -> str | None:
    patterns = [
        r"(?:security\s*code|verification\s*code|one-time\s*code|验证码|確認コード|セキュリティ[\s\u3000]*コード)[:\s：]*([0-9]{6,8})",
        r"セキュリティ[\s\u3000]*コード[：:\s]*([0-9]{6,8})",
        r"(?:コード|code)[：:\s]*([0-9]{6,8})",
        r"\b([0-9]{6})\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            return m.group(1)
    return None


def _message_datetime(msg: email.message.Message) -> datetime | None:
    raw = msg.get("Date")
    if not raw:
        return None
    try:
        dt = parsedate_to_datetime(raw)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def _message_plain_text(msg: email.message.Message) -> str:
    texts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype not in {"text/plain", "text/html"}:
                continue
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            charset = part.get_content_charset() or "utf-8"
            texts.append(payload.decode(charset, errors="replace"))
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            texts.append(payload.decode(charset, errors="replace"))
    return "\n".join(texts)


def _is_microsoft_mail(subject: str, body: str) -> bool:
    s = subject.lower()
    b = body.lower()
    keys = [
        "microsoft",
        "accountprotection.microsoft",
        "account security",
        "security code",
        "verification code",
        "アカウント",
        "セキュリティ",
        "microsoft 帐户",
        "microsoft 账户",
        "account-security-noreply",
    ]
    return any(k in s or k in b for k in keys)


def fetch_microsoft_verification_code(
    recovery_email: str,
    recovery_password: str,
    imap_config: dict[str, Any],
    not_before: datetime | None = None,
    timeout_sec: int = 180,
    poll_sec: int = 8,
    log: LogFn = _noop_log,
) -> tuple[str | None, str]:
    """通过 IMAP 收取 Microsoft 发送到辅助邮箱的验证码。"""
    host, port, use_ssl = resolve_imap_settings(recovery_email, imap_config)
    # 放宽时间窗，避免时区/时钟偏差把刚到的信滤掉
    if not_before is None:
        not_before = datetime.now(timezone.utc) - timedelta(minutes=5)
    else:
        if not_before.tzinfo is None:
            not_before = not_before.replace(tzinfo=timezone.utc)
        not_before = not_before - timedelta(minutes=5)
    deadline = time.time() + max(60, int(timeout_sec))
    last_error = ""
    poll_sec = max(3, int(poll_sec))

    while time.time() < deadline:
        try:
            if use_ssl:
                conn = imaplib.IMAP4_SSL(host, port)
            else:
                conn = imaplib.IMAP4(host, port)
            conn.login(recovery_email, recovery_password)

            mailboxes: list[str] = ["INBOX", "Inbox", "Junk", "Spam", "Bulk Mail", "ゴミ箱", "迷惑メール"]
            try:
                typ, boxes = conn.list()
                if typ == "OK" and boxes:
                    for raw in boxes:
                        try:
                            line = raw.decode("utf-8", errors="ignore") if isinstance(raw, bytes) else str(raw)
                            # 取最后一个引号或空格后的邮箱名
                            name = line.split('"')[-2] if line.count('"') >= 2 else line.split()[-1]
                            name = name.strip()
                            if name and name not in mailboxes and name.upper() != "INBOX":
                                mailboxes.append(name)
                        except Exception:
                            continue
            except Exception:
                pass

            for mailbox in mailboxes[:20]:
                try:
                    typ, _ = conn.select(mailbox, readonly=True)
                    if typ != "OK":
                        continue
                except Exception:
                    continue
                try:
                    status, data = conn.search(None, "ALL")
                except imaplib.IMAP4.error:
                    continue
                if status != "OK" or not data or not data[0]:
                    continue
                ids = data[0].split()
                for msg_id in reversed(ids[-40:]):
                    status, msg_data = conn.fetch(msg_id, "(RFC822)")
                    if status != "OK" or not msg_data:
                        continue
                    raw = msg_data[0][1]
                    msg = email.message_from_bytes(raw)
                    subject = _decode_mime_header(msg.get("Subject", ""))
                    from_addr = _decode_mime_header(msg.get("From", ""))
                    msg_dt = _message_datetime(msg)
                    if msg_dt and msg_dt < not_before:
                        continue
                    body = _message_plain_text(msg)
                    blob = subject + "\n" + from_addr + "\n" + body
                    if not _is_microsoft_mail(subject, blob):
                        continue
                    code = _extract_code_from_text(blob)
                    if code:
                        try:
                            conn.logout()
                        except Exception:
                            pass
                        log(f"IMAP 收到验证码: {code}（邮箱夹:{mailbox}）")
                        return code, "ok"
            try:
                conn.logout()
            except Exception:
                pass
            remain = int(max(0, deadline - time.time()))
            last_error = f"尚未收到 Microsoft 验证码邮件（剩余约 {remain}s）"
        except imaplib.IMAP4.error as exc:
            err = str(exc)
            if "SELECTED" in err.upper() or "SEARCH" in err.upper():
                last_error = (
                    f"IMAP 邮箱未就绪({exc})，请确认辅助邮箱密码正确且已开启 IMAP"
                )
            else:
                last_error = (
                    f"IMAP 登录失败（请确认邮箱密码/授权码正确并开启 IMAP）: {exc}"
                )
        except Exception as exc:
            last_error = f"IMAP 异常: {exc}"

        log(f"等待验证码邮件... ({last_error or '轮询中'})")
        try:
            from foxmail_automation import is_foxmail_cancel_requested

            if is_foxmail_cancel_requested():
                return None, "用户已停止"
        except Exception:
            pass
        time.sleep(poll_sec)

    return None, last_error or "收取验证码超时"


def test_imap_login(
    recovery_email: str,
    recovery_password: str,
    imap_config: dict[str, Any],
) -> tuple[bool, str]:
    """检测辅助邮箱是否已在邮件服务器上存在且可登录。"""
    host, port, use_ssl = resolve_imap_settings(recovery_email, imap_config)
    try:
        if use_ssl:
            conn = imaplib.IMAP4_SSL(host, port)
        else:
            conn = imaplib.IMAP4(host, port)
        conn.login(recovery_email, recovery_password)
        conn.logout()
        return True, "ok"
    except imaplib.IMAP4.error as exc:
        return False, str(exc)
    except Exception as exc:
        return False, str(exc)


def ensure_recovery_mailbox_ready(
    recovery_email: str,
    recovery_password: str,
    imap_config: dict[str, Any],
    full_config: dict[str, Any],
    log: LogFn = _noop_log,
) -> tuple[bool, str]:
    """
    确保辅助邮箱可用（Microsoft 提交前执行）：
    1. IMAP 试探登录 — 成功说明邮箱已存在
    2. 失败且开启 auto_create → Foxmail 添加「其它邮箱」
    3. 再次 IMAP 试探
    """
    fox_cfg = full_config.get("foxmail") or {}
    create_wait = int(full_config.get("recovery_mail_create_wait_sec", 45))

    ok, detail = test_imap_login(recovery_email, recovery_password, imap_config)
    if ok:
        log(f"辅助邮箱 IMAP 已可用: {recovery_email}")
        return True, "辅助邮箱已就绪"

    log(f"辅助邮箱 IMAP 暂不可用: {detail}")

    if not fox_cfg.get("auto_create", True):
        return False, (
            f"辅助邮箱无法登录({detail})。"
            "请先在邮件服务器/AOL 注册创建该邮箱，或在 config 开启 foxmail.auto_create"
        )

    log("尝试通过 Foxmail 添加「其它邮箱」（此步需 Foxmail 窗口可见）...")
    fx_ok, fx_detail = create_other_mailbox_account(
        recovery_email, recovery_password, full_config, log=log, force_ui=True
    )
    if not fx_ok:
        return False, (
            f"Foxmail 添加失败: {fx_detail}。"
            "说明：Foxmail 只能添加已有邮箱，无法在 AOL 等服务器上凭空建号。"
        )

    deadline = time.time() + create_wait
    while time.time() < deadline:
        ok2, detail2 = test_imap_login(recovery_email, recovery_password, imap_config)
        if ok2:
            log("Foxmail 添加后 IMAP 已可用")
            return True, "辅助邮箱已通过 Foxmail 添加并就绪"
        time.sleep(3)

    return False, (
        f"Foxmail 已操作但 IMAP 仍不可用({detail})。"
        "该辅助邮箱可能尚未在 AOL/邮件服务器上真正创建，请先手动注册后再跑。"
    )


def wait_for_recovery_input(page: Page, timeout_sec: float = 12.0) -> bool:
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if find_recovery_email_input(page):
            return True
        if is_code_verify_page(page):
            return True
        safe_wait(page, 350)
    return False


def submit_verification_code(page: Page, code: str, log: LogFn = _noop_log) -> bool:
    loc = find_code_input(page)
    if loc is None:
        if not is_code_verify_page(page):
            return False
        loc = page.locator('input[type="text"]').first
    log(f"填写验证码: {code}")
    try:
        loc.click(timeout=500)
        loc.fill(code, timeout=1500)
    except Exception:
        return False
    wait_and_click(page, SUBMIT_SELECTORS, timeout_ms=2000)
    safe_wait(page, 400)
    return True


def bind_recovery_email(
    page: Page,
    recovery_email: str,
    recovery_password: str,
    imap_config: dict[str, Any],
    *,
    full_config: dict[str, Any] | None = None,
    code_timeout_sec: int = 180,
    poll_sec: int = 8,
    create_wait_sec: int = 30,
    log: LogFn = _noop_log,
) -> tuple[bool, str]:
    """
    绑定辅助邮箱：
    1. Microsoft 填辅助邮箱 → 点「次へ」
    2. 默认 IMAP 后台收验证码（无需 Foxmail，不被浏览器遮挡）
    3. 可选 Foxmail UI（use_ui=true 时，会临时最小化浏览器再操作 Foxmail）
    4. 回填验证码
    """
    if not recovery_email or not recovery_password:
        return False, "未提供辅助邮箱或密码"

    full_config = full_config or {}
    fox_cfg = full_config.get("foxmail") or {}

    try:
        maybe_bring_to_front(page)
    except Exception:
        pass

    # ── 步骤 1：Microsoft 填辅助邮箱并点下一步 ──
    if is_code_verify_page(page):
        send_time = datetime.now(timezone.utc)
        log("已在验证码输入页，跳过填辅助邮箱")
    elif is_recovery_bind_page(page) or wait_for_recovery_input(page, timeout_sec=15):
        if not submit_recovery_email(page, recovery_email, log=log):
            return False, "未能填写辅助邮箱或点击次へ"
        send_time = datetime.now(timezone.utc)
        log("已在 Microsoft 提交辅助邮箱并点击次へ，Microsoft 正在发送验证码...")
        safe_wait(page, 1500)
    else:
        return False, "未检测到绑定辅助邮箱页面"

    # ── 步骤 2：先在 Foxmail「新建」已有辅助邮箱，再 IMAP 收验证码 ──
    # 邮箱账号本身已存在，只需加进 Foxmail 后才能稳定收信
    if fox_cfg.get("enabled", True) and fox_cfg.get("auto_create", True):
        log(f"Foxmail：新建/添加已有辅助邮箱 {recovery_email} ...")
        fx_ok, fx_detail = create_other_mailbox_account(
            recovery_email,
            recovery_password,
            full_config,
            log=log,
            force_ui=True,
        )
        if not fx_ok:
            log(f"Foxmail 添加失败({fx_detail})，仍尝试 IMAP 收信")
        else:
            log(f"Foxmail: {fx_detail}")
    else:
        log(f"IMAP 后台收验证码（跳过 Foxmail 新建）: {recovery_email}")

    try:
        maybe_bring_to_front(page)
    except Exception:
        pass

    # 等待进入验证码输入页
    if not is_code_verify_page(page):
        deadline = time.time() + 25
        while time.time() < deadline:
            if is_code_verify_page(page):
                break
            safe_wait(page, 400)

    # ── 步骤 3：IMAP 收取验证码 ──
    effective_timeout = max(code_timeout_sec, create_wait_sec)
    log(f"IMAP 收取验证码: {recovery_email}")
    code, imap_detail = fetch_microsoft_verification_code(
        recovery_email,
        recovery_password,
        imap_config,
        not_before=send_time,
        timeout_sec=effective_timeout,
        poll_sec=poll_sec,
        log=log,
    )
    if not code:
        return False, imap_detail

    # ── 步骤 4：回填验证码 ──
    try:
        maybe_bring_to_front(page)
    except Exception:
        pass

    if not submit_verification_code(page, code, log=log):
        return False, f"已收到验证码 {code}，但未能填入页面"

    safe_wait(page, 2000)
    text = page_text_lower(page)
    if "outlook.live.com/mail" in page.url.lower() or "outlook.office.com/mail" in page.url.lower():
        return True, "辅助邮箱绑定完成，已进入邮箱"
    if any(x in text for x in ["success", "verified", "完了", "成功", "added"]):
        return True, "辅助邮箱绑定完成"
    if is_recovery_bind_page(page) or is_code_verify_page(page):
        return False, "验证码已提交，但页面仍未完成绑定"
    return True, "验证码已提交"


def prepare_recovery_for_manual_code(
    page: Page,
    recovery_email: str,
    log: LogFn = _noop_log,
) -> tuple[bool, str]:
    """
    只做：填辅助邮箱 → 点次へ → 等到验证码页。
    不收信、不回填；由控制台人工输入验证码后再 submit_verification_code。
    """
    if not recovery_email:
        return False, "未提供辅助邮箱"

    try:
        maybe_bring_to_front(page)
    except Exception:
        pass

    if is_phone_verify_page(page):
        return False, "需要电话认证（当前不是辅助邮箱验证码页）"

    before_url = ""
    try:
        before_url = page.url or ""
    except Exception:
        pass

    # 仍在辅助邮箱输入页时，绝不能当成已到验证码页
    if _is_still_recovery_email_page(page) or is_recovery_bind_page(page):
        if is_code_verify_page(page):
            pass  # 理论上不会同时成立
        elif not submit_recovery_email(page, recovery_email, log=log):
            if _recovery_flow_advanced(page, before_url):
                log("填表回报失败但页面已前进，继续等待验证码页")
            else:
                return False, "未能填写辅助邮箱或点击次へ"
        log(f"已提交辅助邮箱 {recovery_email}，等待验证码页…")
        safe_wait(page, 1200)
    elif is_code_verify_page(page):
        log("已在验证码输入页")
        return True, f"机子辅助邮箱 {recovery_email} 已到验证码页，请在控制台输入验证码"
    elif wait_for_recovery_input(page, timeout_sec=15):
        if is_phone_verify_page(page):
            return False, "需要电话认证（当前不是辅助邮箱验证码页）"
        if _is_still_recovery_email_page(page) or is_recovery_bind_page(page):
            if not submit_recovery_email(page, recovery_email, log=log):
                if _recovery_flow_advanced(page, before_url):
                    log("填表回报失败但页面已前进，继续等待验证码页")
                else:
                    return False, "未能填写辅助邮箱或点击次へ"
            log(f"已提交辅助邮箱 {recovery_email}，等待验证码页…")
            safe_wait(page, 1200)
        elif is_code_verify_page(page):
            return True, f"机子辅助邮箱 {recovery_email} 已到验证码页，请在控制台输入验证码"
        else:
            return False, "未检测到绑定辅助邮箱页面"
    else:
        if _recovery_flow_advanced(page, before_url):
            return True, (
                f"已填辅助邮箱 {recovery_email}（页面已前进，请查收验证码或在控制台输入）"
            )
        return False, "未检测到绑定辅助邮箱页面"

    deadline = time.time() + 45
    while time.time() < deadline:
        if is_phone_verify_page(page):
            return False, "需要电话认证（提交后进入电话认证，非邮件验证码）"
        if is_code_verify_page(page):
            return True, (
                f"已填辅助邮箱 {recovery_email}，请查收验证码后在控制台该机子号后输入"
            )
        if _recovery_flow_advanced(page, before_url):
            if is_code_verify_page(page):
                return True, (
                    f"已填辅助邮箱 {recovery_email}，请查收验证码后在控制台该机子号后输入"
                )
            url = (page.url or "").lower()
            if "outlook.live.com/mail" in url or "outlook.office.com/mail" in url:
                return True, f"辅助邮箱 {recovery_email} 已提交，页面已进入邮箱"
            return True, (
                f"已填辅助邮箱 {recovery_email}，请查收验证码后在控制台该机子号后输入"
            )
        if _is_still_recovery_email_page(page):
            submit_recovery_email(page, recovery_email, log=log)
            safe_wait(page, 900)
            continue
        safe_wait(page, 450)

    if is_phone_verify_page(page):
        return False, "需要电话认证（非辅助邮箱验证码页）"
    if is_code_verify_page(page) and not _is_still_recovery_email_page(page):
        return True, f"已填辅助邮箱 {recovery_email}，请输入验证码"
    if _recovery_flow_advanced(page, before_url):
        return True, (
            f"已填辅助邮箱 {recovery_email}（页面已前进，请查收验证码或在控制台输入）"
        )
    if _is_still_recovery_email_page(page) or is_recovery_bind_page(page):
        return False, f"未能离开辅助邮箱输入页（邮箱: {recovery_email}），请点「打开」查看"
    return False, (
        f"已提交辅助邮箱 {recovery_email}，但未确认进入验证码页，请点「打开」查看后重试"
    )


def try_bind_recovery_on_page(
    page: Page,
    recovery_email: str,
    recovery_password: str,
    imap_config: dict[str, Any],
    config: dict[str, Any],
    log: LogFn = _noop_log,
) -> tuple[bool, str] | None:
    """若当前页是绑定辅助邮箱流程则处理，否则返回 None。"""
    try:
        maybe_bring_to_front(page)
    except Exception:
        pass
    safe_wait(page, 500)

    if is_phone_verify_page(page):
        log("当前为电话认证页，不是辅助邮箱验证码页")
        return False, "需要电话认证（非辅助邮箱验证码，未应领取辅助邮箱）"

    if not is_recovery_bind_page(page) and not is_code_verify_page(page):
        return None

    # 默认人工验证码（Foxmail/IMAP 常不可用）；config.recovery_manual_code=false 才走自动收信
    manual = bool((config or {}).get("recovery_manual_code", True))
    if manual:
        return prepare_recovery_for_manual_code(page, recovery_email, log=log)

    return bind_recovery_email(
        page,
        recovery_email,
        recovery_password,
        imap_config,
        full_config=config,
        code_timeout_sec=int(config.get("recovery_code_timeout_sec", 180)),
        poll_sec=int(config.get("recovery_code_poll_sec", 8)),
        create_wait_sec=int(config.get("recovery_mail_create_wait_sec", 30)),
        log=log,
    )
