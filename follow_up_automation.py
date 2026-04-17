import argparse
import os
import re
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, cast
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from bs4 import BeautifulSoup
from dotenv import load_dotenv
from notion_client import Client

from app import (
    FIELD_ALIASES,
    _extract_linkedin_url_from_html,
    _find_matching_export_for_url,
    _read_html_source,
    extract_profile_from_html,
    normalize_field_key,
    normalize_linkedin_url,
    normalize_notion_database_id,
)

@dataclass
class PendingContact:
    page_id: str
    name: str
    linkedin_url: str
    status: str


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value


def _parse_csv_values(raw: str) -> List[str]:
    return [piece.strip() for piece in raw.split(",") if piece.strip()]


def _env_bool(name: str, default: bool = False) -> bool:
    fallback = "true" if default else "false"
    raw = os.getenv(name, fallback).strip().lower()
    return raw in {"1", "true", "yes", "y", "on"}


def _norm(value: str) -> str:
    return normalize_field_key(value)


def _property_name_or_none(props: Dict[str, Any], expected_name: str) -> Optional[str]:
    aliases = {_norm(expected_name)}
    for alias in FIELD_ALIASES.get(expected_name, []):
        aliases.add(_norm(alias))
    for prop_name in props.keys():
        if _norm(prop_name) in aliases:
            return prop_name
    return None


def _extract_plain_text(prop_value: Dict[str, Any]) -> str:
    prop_type = prop_value.get("type")
    if prop_type == "title":
        return "".join(piece.get("plain_text", "") for piece in prop_value.get("title", []))
    if prop_type == "rich_text":
        return "".join(piece.get("plain_text", "") for piece in prop_value.get("rich_text", []))
    if prop_type == "url":
        return (prop_value.get("url") or "").strip()
    if prop_type == "email":
        return (prop_value.get("email") or "").strip()
    if prop_type == "phone_number":
        return (prop_value.get("phone_number") or "").strip()
    if prop_type == "select":
        select = prop_value.get("select")
        return (select or {}).get("name", "").strip() if isinstance(select, dict) else ""
    if prop_type == "status":
        status = prop_value.get("status")
        return (status or {}).get("name", "").strip() if isinstance(status, dict) else ""
    if prop_type == "formula":
        formula = prop_value.get("formula") or {}
        if formula.get("type") == "string":
            return (formula.get("string") or "").strip()
    return ""


def _first_title_property_name(props: Dict[str, Any]) -> Optional[str]:
    for name, prop in props.items():
        if prop.get("type") == "title":
            return name
    return None


def _normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def _linkedin_slug_from_url(url: str) -> str:
    path = urlsplit(url).path.strip("/")
    parts = path.split("/")
    if len(parts) >= 2 and parts[0] == "in":
        return parts[1].lower()
    return ""


def _html_matches_profile_url(html: str, profile_url: str) -> bool:
    target = normalize_linkedin_url(profile_url)
    extracted_url = _extract_linkedin_url_from_html(html)
    if extracted_url:
        try:
            return normalize_linkedin_url(extracted_url) == target
        except Exception:
            pass

    target_slug = _linkedin_slug_from_url(target)
    if not target_slug:
        return True

    return target_slug in html.lower()


def _safe_export_filename(profile_url: str) -> str:
    slug = _linkedin_slug_from_url(profile_url) or "linkedin-profile"
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "-", slug).strip("-") or "linkedin-profile"
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{slug}-{stamp}.html"


def _auto_capture_linkedin_export_with_mode(
    profile_url: str,
    load_delay_sec: float = 8.0,
    activate_window: bool = False,
) -> str:
    if sys.platform != "darwin":
        return ""

    escaped_url = _escape_applescript_string(profile_url)
    activate_line = "  activate\n" if activate_window else ""
    script = (
        'tell application "Safari"\n'
        "  if not running then launch\n"
        f"{activate_line}"
        "  if (count of windows) = 0 then\n"
        "    make new document\n"
        "  end if\n"
        f'  set URL of front document to "{escaped_url}"\n'
        f"  delay {load_delay_sec}\n"
        "  set pageSource to source of front document\n"
        "end tell\n"
        "return pageSource\n"
    )

    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            check=False,
            capture_output=True,
            text=True,
            timeout=max(45, int(load_delay_sec + 25)),
        )
    except Exception:
        return ""

    if result.returncode != 0:
        return ""

    html = (result.stdout or "").strip()
    if len(html) < 300:
        return ""

    downloads_dir = Path.home() / "Downloads"
    downloads_dir.mkdir(parents=True, exist_ok=True)
    output_path = downloads_dir / _safe_export_filename(profile_url)
    try:
        output_path.write_text(html, encoding="utf-8")
    except Exception:
        return ""

    return str(output_path)


def _fetch_profile_html(
    profile_url: str,
    profile_load_delay_sec: float = 8.0,
    safari_activate_window: bool = False,
) -> str:
    direct_error: Optional[Exception] = None
    blocked_999 = False

    try:
        direct_html = _read_html_source(profile_url)
        if _html_matches_profile_url(direct_html, profile_url):
            return direct_html
        print("⚠️ Direct profile fetch returned a different profile. Trying local export fallback.")
    except Exception as err:
        direct_error = err
        blocked_999 = "(999)" in str(err)
        if not blocked_999:
            print(f"⚠️ Direct profile fetch failed ({err}). Trying local export fallback.")

    fallback_path = _find_matching_export_for_url(profile_url)
    if fallback_path:
        fallback_html = _read_html_source(fallback_path)
        if _html_matches_profile_url(fallback_html, profile_url):
            if blocked_999:
                print(f"⚠️ LinkedIn blocked access (999). Using local export: {fallback_path}")
            return fallback_html
        print(f"⚠️ Ignoring mismatched local export: {fallback_path}")

    captured_path = _auto_capture_linkedin_export_with_mode(
        profile_url,
        load_delay_sec=profile_load_delay_sec,
        activate_window=safari_activate_window,
    )
    if captured_path:
        captured_html = _read_html_source(captured_path)
        if _html_matches_profile_url(captured_html, profile_url):
            label = "Auto-captured local export"
            if blocked_999:
                print(f"⚠️ LinkedIn blocked access (999). {label}: {captured_path}")
            else:
                print(f"ℹ️ {label}: {captured_path}")
            return captured_html
        print(f"⚠️ Ignoring mismatched auto-captured export: {captured_path}")

    if direct_error and not blocked_999:
        raise direct_error
    raise RuntimeError(
        "Could not load a matching profile HTML for this URL. "
        "Open the exact profile in Safari while logged in, then run again."
    )


def _connection_window(text_nodes: List[str], full_name: str) -> List[str]:
    if not text_nodes:
        return []
    if not full_name:
        return text_nodes[:120]

    idx = None
    for i, value in enumerate(text_nodes[:1000]):
        if value.strip() == full_name:
            idx = i
            break
    if idx is None:
        return text_nodes[:120]
    return text_nodes[idx : idx + 140]


def _topcard_action_markers(soup: BeautifulSoup) -> Dict[str, bool]:
    markers = {
        "has_message": False,
        "has_pending": False,
        "has_connect": False,
    }

    containers = soup.select('[componentkey*="Topcard"]')
    if not containers:
        containers = [soup]

    pending_patterns = [
        "pending",
        "invitation sent",
        "withdraw invitation",
        "invite sent",
        "cancel invitation",
    ]

    for container in containers[:6]:
        for element in container.find_all(["a", "button"], limit=200):
            text = _normalize_space(element.get_text(" ", strip=True)).lower()
            aria = _normalize_space(str(element.get("aria-label") or "")).lower()
            href = str(element.get("href") or "").strip().lower()
            label = " ".join(piece for piece in [text, aria] if piece).strip()

            if "/messaging/compose/" in href or re.search(r"\bmessage\b", label):
                markers["has_message"] = True
            if any(pattern in label for pattern in pending_patterns):
                markers["has_pending"] = True
            if re.search(r"\bconnect\b", label) or re.search(r"\binvite\b", label):
                markers["has_connect"] = True

    return markers


def _invite_status_with_safari(
    profile_url: str,
    load_delay_sec: float = 8.0,
    activate_window: bool = False,
) -> Optional[bool]:
    if sys.platform != "darwin":
        return None

    escaped_url = _escape_applescript_string(profile_url)
    status_js = (
        "(() => {"
        " const norm = (s) => (s || '').replace(/\\s+/g, ' ').trim().toLowerCase();"
        " const visible = (el) => !!el && el.getClientRects().length > 0 && window.getComputedStyle(el).visibility !== 'hidden' && window.getComputedStyle(el).display !== 'none';"
        " const controls = Array.from(document.querySelectorAll('a[href], button, [role=\"button\"]')).filter((el) => visible(el)).map((el) => {"
        "   const r = el.getBoundingClientRect();"
        "   return {"
        "     text: norm(el.innerText || el.textContent),"
        "     aria: norm(el.getAttribute('aria-label') || ''),"
        "     href: (el.getAttribute('href') || '').toLowerCase(),"
        "     x: r.x, y: r.y"
        "   };"
        " });"
        " const inActionZone = controls.filter((c) => c.x < (window.innerWidth * 0.75) && c.y > 40 && c.y < Math.min(window.innerHeight, 650));"
        " const active = inActionZone.length ? inActionZone : controls;"
        " const label = (c) => `${c.text} ${c.aria}`.trim();"
        " const hasPending = active.some((c) => {"
        "   const t = label(c);"
        "   return /\\bpending\\b/.test(t) || t.includes('invitation sent') || t.includes('withdraw invitation') || t.includes('invite sent') || t.includes('cancel invitation');"
        " });"
        " const hasConnect = active.some((c) => {"
        "   const t = label(c);"
        "   return /\\bconnect\\b/.test(t) || /\\binvite\\b/.test(t);"
        " });"
        " const hasMessage = active.some((c) => c.href.includes('/messaging/compose/') || /\\bmessage\\b/.test(label(c)));"
        " if (hasPending || (hasConnect && !hasMessage)) return 'PENDING';"
        " if (hasMessage && !hasConnect) return 'ACCEPTED';"
        " return 'UNKNOWN';"
        "})()"
    )
    escaped_js = _escape_applescript_string(status_js)
    activate_line = "  activate\n" if activate_window else ""
    script = (
        'tell application "Safari"\n'
        "  if not running then launch\n"
        f"{activate_line}"
        "  if (count of windows) = 0 then\n"
        "    make new document\n"
        "  end if\n"
        f'  set URL of front document to "{escaped_url}"\n'
        f"  delay {load_delay_sec}\n"
        f'  set statusResult to do JavaScript "{escaped_js}" in front document\n'
        "  return statusResult\n"
        "end tell\n"
    )

    result = subprocess.run(
        ["osascript", "-e", script],
        check=False,
        capture_output=True,
        text=True,
        timeout=max(50, int(load_delay_sec + 25)),
    )
    if result.returncode != 0:
        return None

    status = (result.stdout or "").strip().upper()
    if status == "ACCEPTED":
        return True
    if status == "PENDING":
        return False
    return None


def invite_is_accepted(
    profile_url: str,
    profile_load_delay_sec: float = 8.0,
    safari_activate_window: bool = False,
) -> bool:
    safari_status = _invite_status_with_safari(
        profile_url,
        load_delay_sec=profile_load_delay_sec,
        activate_window=safari_activate_window,
    )
    if safari_status is not None:
        return safari_status

    html = _fetch_profile_html(
        profile_url,
        profile_load_delay_sec=profile_load_delay_sec,
        safari_activate_window=safari_activate_window,
    )
    profile = extract_profile_from_html(html, profile_url)
    full_name = (profile.get("full_name") or "").strip()

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "template"]):
        tag.decompose()

    action_markers = _topcard_action_markers(soup)
    if action_markers["has_pending"]:
        return False
    if action_markers["has_message"] and not action_markers["has_connect"]:
        return True
    if action_markers["has_connect"]:
        return False

    text_nodes = [_normalize_space(t) for t in soup.stripped_strings]
    text_nodes = [t for t in text_nodes if t]
    window = _connection_window(text_nodes, full_name)
    lowered = [value.lower() for value in window]

    pending_markers = [
        "pending",
        "invitation sent",
        "withdraw invitation",
        "invite sent",
    ]
    if any(any(marker in value for marker in pending_markers) for value in lowered):
        return False

    has_message = any(re.search(r"\bmessage\b", value) for value in lowered)
    has_connect = any(re.search(r"\bconnect\b", value) or re.search(r"\binvite\b", value) for value in lowered)

    if has_message and not has_connect:
        return True
    return False


def _read_template(path: str) -> str:
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read()

    default_template = (
        "Hi {first_name},\n\n"
        "Great to connect with you here on LinkedIn.\n"
        "I wanted to follow up and say thanks for accepting my invitation.\n\n"
        "- Liam"
    )
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(default_template)
    print(f"ℹ️ Created message template at {path}")
    return default_template


def _render_message(template: str, contact: PendingContact) -> str:
    first_name = contact.name.split()[0] if contact.name else ""
    context = {
        "name": contact.name,
        "first_name": first_name,
        "linkedin_url": contact.linkedin_url,
    }
    try:
        return template.format(**context)
    except KeyError as err:
        missing = str(err).strip("'")
        raise RuntimeError(
            f"Template placeholder {{{missing}}} is not supported. "
            "Use {name}, {first_name}, or {linkedin_url}."
        ) from err


def _escape_applescript_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")


def _set_or_replace_query_param(url: str, key: str, value: str) -> str:
    parsed = urlsplit(url)
    pairs = [(k, v) for k, v in parse_qsl(parsed.query, keep_blank_values=True) if k.lower() != key.lower()]
    pairs.append((key, value))
    new_query = urlencode(pairs, doseq=True)
    return urlunsplit((parsed.scheme, parsed.netloc, parsed.path, new_query, parsed.fragment))


def _auto_send_linkedin_message_safari(
    profile_url: str,
    message: str,
    profile_load_delay_sec: float = 8.0,
    compose_load_delay_sec: float = 4.0,
    activate_window: bool = False,
) -> bool:
    if sys.platform != "darwin":
        raise RuntimeError("Automatic send is currently supported only on macOS (Safari).")

    find_compose_url_js = (
        "(() => {"
        " const norm = (s) => (s || '').replace(/\\s+/g, ' ').trim().toLowerCase();"
        " const visible = (el) => !!el && el.getClientRects().length > 0 && window.getComputedStyle(el).visibility !== 'hidden' && window.getComputedStyle(el).display !== 'none';"
        " const width = window.innerWidth || 1200;"
        " const height = window.innerHeight || 900;"
        " const links = Array.from(document.querySelectorAll('a[href*=\"/messaging/compose/\"]'));"
        " const candidates = links.filter((el) => visible(el) && norm(el.innerText || el.textContent) === 'message').map((el) => {"
        "   const r = el.getBoundingClientRect();"
        "   const href = el.href || '';"
        "   const aria = (el.getAttribute('aria-label') || '').toLowerCase();"
        "   let score = 0;"
        "   if (r.x < width * 0.65) score += 2;"
        "   if (r.y > 60 && r.y < height * 0.9) score += 2;"
        "   if (href.includes('screenContext=NON_SELF_PROFILE_VIEW')) score += 1;"
        "   if (href.includes('recipient=')) score += 1;"
        "   if (aria.includes('premium')) score -= 1;"
        "   return { href, score, x: r.x, y: r.y };"
        " });"
        " if (!candidates.length) return '';"
        " candidates.sort((a, b) => b.score - a.score || a.y - b.y || a.x - b.x);"
        " return candidates[0].href || '';"
        "})()"
    )
    click_send_js = (
        "(() => {"
        " const visible = (el) => !!el && el.getClientRects().length > 0 && window.getComputedStyle(el).visibility !== 'hidden' && window.getComputedStyle(el).display !== 'none';"
        " let box = document.querySelector('div.msg-form__contenteditable[role=\"textbox\"][contenteditable=\"true\"]');"
        " if (!box || !visible(box)) box = document.querySelector('[contenteditable=\"true\"][role=\"textbox\"]');"
        " if (!box || !visible(box)) box = document.querySelector('[contenteditable=\"true\"][aria-label*=\"message\" i]');"
        " if (!box) return 'ERR_NO_TEXTBOX';"
        " const root = box.closest('form.msg-form') || box.closest('form') || box.closest('.msg-form') || box.parentElement;"
        " const isSendBtn = (btn) => {"
        "   const txt = (btn.innerText || btn.textContent || '').trim().toLowerCase();"
        "   const aria = (btn.getAttribute('aria-label') || '').toLowerCase();"
        "   return btn.classList.contains('msg-form__send-button') || txt === 'send' || aria.includes('send');"
        " };"
        " const candidates = Array.from((root || document).querySelectorAll('button.msg-form__send-button, button[aria-label*=\"Send\" i], button')).filter((btn) => visible(btn) && isSendBtn(btn));"
        " if (!candidates.length) return 'ERR_NO_SEND_BUTTON';"
        " const sendBtn = candidates.find((btn) => btn.classList.contains('msg-form__send-button') && !btn.disabled) || candidates.find((btn) => !btn.disabled) || candidates[0];"
        " if (!sendBtn) return 'ERR_NO_SEND_BUTTON';"
        " if (sendBtn.disabled) return 'ERR_SEND_DISABLED';"
        " sendBtn.click();"
        " return 'OK_SENT';"
        "})()"
    )

    escaped_url = _escape_applescript_string(profile_url)
    escaped_find_compose_url_js = _escape_applescript_string(find_compose_url_js)
    activate_line = "  activate\n" if activate_window else ""
    find_script = (
        'tell application "Safari"\n'
        "  if not running then launch\n"
        f"{activate_line}"
        "  if (count of windows) = 0 then\n"
        "    make new document\n"
        "  end if\n"
        f'  set URL of front document to "{escaped_url}"\n'
        f"  delay {profile_load_delay_sec}\n"
        f'  set composeUrl to do JavaScript "{escaped_find_compose_url_js}" in front document\n'
        "  return composeUrl\n"
        "end tell\n"
    )

    find_result = subprocess.run(
        ["osascript", "-e", find_script],
        check=False,
        capture_output=True,
        text=True,
        timeout=max(45, int(profile_load_delay_sec + 20)),
    )
    if find_result.returncode != 0:
        stderr = (find_result.stderr or "").strip()
        raise RuntimeError(f"Safari automation failed: {stderr or 'unknown osascript error'}")

    compose_url = (find_result.stdout or "").strip()
    if not compose_url:
        raise RuntimeError("Auto-send flow result: ERR_NO_COMPOSE_URL")

    compose_url_with_body = _set_or_replace_query_param(compose_url, "body", message)
    escaped_compose_url = _escape_applescript_string(compose_url_with_body)
    escaped_click_send_js = _escape_applescript_string(click_send_js)
    send_script = (
        'tell application "Safari"\n'
        "  if not running then launch\n"
        f"{activate_line}"
        f'  set URL of front document to "{escaped_compose_url}"\n'
        f"  delay {compose_load_delay_sec}\n"
        f'  set sendResult to do JavaScript "{escaped_click_send_js}" in front document\n'
        "  if sendResult starts with \"OK_SENT\" then\n"
        "    return sendResult\n"
        "  end if\n"
        "  delay 1\n"
        f'  set sendResult to do JavaScript "{escaped_click_send_js}" in front document\n'
        "  if sendResult starts with \"OK_SENT\" then\n"
        "    return sendResult\n"
        "  end if\n"
        "  delay 1\n"
        f'  set sendResult to do JavaScript "{escaped_click_send_js}" in front document\n'
        "  return sendResult\n"
        "end tell\n"
    )

    send_result = subprocess.run(
        ["osascript", "-e", send_script],
        check=False,
        capture_output=True,
        text=True,
        timeout=max(70, int(compose_load_delay_sec + 35)),
    )
    if send_result.returncode != 0:
        stderr = (send_result.stderr or "").strip()
        raise RuntimeError(f"Safari automation failed: {stderr or 'unknown osascript error'}")

    send_output = (send_result.stdout or "").strip()
    if not send_output.startswith("OK_SENT"):
        raise RuntimeError(f"Auto-send flow result: {compose_url}|{send_output or 'empty result'}")
    return True


def _copy_to_clipboard(text: str) -> bool:
    if sys.platform == "darwin":
        try:
            subprocess.run(["pbcopy"], input=text.encode("utf-8"), check=False)
            return True
        except Exception:
            return False
    return False


def _open_profile_in_browser(url: str) -> None:
    if sys.platform == "darwin":
        subprocess.run(["open", url], check=False)
        return
    webbrowser.open(url)


class NotionFollowupWorkflow:
    def __init__(self, api_key: str, contacts_database_id: str):
        self.client = Client(auth=api_key)
        self.contacts_database_id = contacts_database_id

        try:
            contacts_db = cast(Dict[str, Any], self.client.databases.retrieve(database_id=contacts_database_id))
        except Exception as err:
            raise RuntimeError(
                "Could not open Contacts database in Notion. "
                "Check NOTION_CONTACTS_DATABASE_ID and integration sharing."
            ) from err

        data_sources = contacts_db.get("data_sources") or []
        if not data_sources:
            raise RuntimeError("No data sources found on this Contacts database.")
        self.contacts_data_source_id = data_sources[0]["id"]

        ds = cast(
            Dict[str, Any],
            self.client.data_sources.retrieve(data_source_id=self.contacts_data_source_id),
        )
        props = ds.get("properties")
        if not isinstance(props, dict):
            raise RuntimeError("Could not read Contacts database properties.")
        self.props = props

        self.name_prop = _property_name_or_none(self.props, "Name") or _first_title_property_name(self.props)
        self.linkedin_prop = _property_name_or_none(self.props, "LinkedIn")
        self.status_prop = _property_name_or_none(self.props, "Status")
        if not self.name_prop:
            raise RuntimeError("Could not find Name/title property on Contacts database.")
        if not self.linkedin_prop:
            raise RuntimeError("Could not find LinkedIn property on Contacts database.")
        if not self.status_prop:
            raise RuntimeError("Could not find Status property on Contacts database.")

    def _query_all_pages(self) -> List[Dict[str, Any]]:
        results: List[Dict[str, Any]] = []
        cursor: Optional[str] = None
        while True:
            kwargs: Dict[str, Any] = {
                "data_source_id": self.contacts_data_source_id,
                "page_size": 100,
            }
            if cursor:
                kwargs["start_cursor"] = cursor
            response = cast(Dict[str, Any], self.client.data_sources.query(**kwargs))
            results.extend(cast(List[Dict[str, Any]], response.get("results") or []))
            if not response.get("has_more"):
                break
            cursor = cast(Optional[str], response.get("next_cursor"))
        return results

    def pending_contacts(self, pending_statuses: Set[str]) -> List[PendingContact]:
        pages = self._query_all_pages()
        pending: List[PendingContact] = []

        for page in pages:
            page_props = cast(Dict[str, Any], page.get("properties") or {})
            status_raw = _extract_plain_text(cast(Dict[str, Any], page_props.get(self.status_prop) or {}))
            if _norm(status_raw) not in pending_statuses:
                continue

            linkedin_raw = _extract_plain_text(cast(Dict[str, Any], page_props.get(self.linkedin_prop) or {}))
            if not linkedin_raw:
                continue
            try:
                linkedin_url = normalize_linkedin_url(linkedin_raw)
            except Exception:
                continue

            name = _extract_plain_text(cast(Dict[str, Any], page_props.get(self.name_prop) or {})).strip() or "Unknown"
            pending.append(
                PendingContact(
                    page_id=page["id"],
                    name=name,
                    linkedin_url=linkedin_url,
                    status=status_raw,
                )
            )

        return pending

    def update_status(self, contact: PendingContact, new_status: str) -> None:
        prop_def = cast(Dict[str, Any], self.props[self.status_prop])
        prop_type = prop_def.get("type")
        if prop_type == "status":
            payload = {self.status_prop: {"status": {"name": new_status}}}
        elif prop_type == "select":
            payload = {self.status_prop: {"select": {"name": new_status}}}
        else:
            raise RuntimeError(
                f"Status property type '{prop_type}' is unsupported for updates (needs status/select)."
            )

        self.client.pages.update(page_id=contact.page_id, properties=payload)


def run(args: argparse.Namespace) -> None:
    notion_api_key = required_env("NOTION_API_KEY")
    contacts_db = normalize_notion_database_id(required_env("NOTION_CONTACTS_DATABASE_ID"))

    pending_status_values = _parse_csv_values(
        os.getenv("PENDING_CONNECTION_STATUSES", "Request Connection,Connection Requested,Pending")
    )
    pending_statuses = {_norm(value) for value in pending_status_values}

    invite_accepted_status = os.getenv("INVITE_ACCEPTED_STATUS", "Invite Accepted").strip() or "Invite Accepted"
    initial_reachout_status = (
        os.getenv("INITIAL_REACHOUT_STATUS", "Initial Reachout Initiated").strip()
        or "Initial Reachout Initiated"
    )
    auto_send_messages = _env_bool("AUTO_SEND_LINKEDIN_MESSAGES", default=True)
    if args.manual_send:
        auto_send_messages = False
    if auto_send_messages and sys.platform != "darwin":
        print("⚠️ Auto-send is only supported on macOS Safari in this version. Falling back to manual mode.")
        auto_send_messages = False
    safari_activate_window = _env_bool("SAFARI_ACTIVATE_WINDOW", default=False)
    profile_load_delay_sec = float(os.getenv("SAFARI_PROFILE_LOAD_DELAY_SEC", "8"))
    compose_load_delay_sec = float(
        os.getenv("SAFARI_COMPOSE_LOAD_DELAY_SEC", os.getenv("SAFARI_OVERLAY_DELAY_SEC", "4"))
    )

    only_url = (args.only_url or "").strip()
    only_url = normalize_linkedin_url(only_url) if only_url else ""

    workflow = NotionFollowupWorkflow(notion_api_key, contacts_db)
    pending = workflow.pending_contacts(pending_statuses)

    if only_url:
        pending = [contact for contact in pending if contact.linkedin_url == only_url]

    if not pending:
        print("No matching pending contacts found.")
        return

    print(f"Found {len(pending)} pending contact(s) to check.")

    template_path = args.template or os.getenv("LINKEDIN_MESSAGE_TEMPLATE_PATH", "message_template.txt")
    template_text = _read_template(template_path)

    accepted: List[PendingContact] = []
    approved: List[PendingContact] = []
    sent: List[PendingContact] = []
    approved_message_by_page_id: Dict[str, str] = {}
    verify_failed = 0
    still_pending = 0

    print("\nPhase 1/3: Checking invite statuses and updating Invite Accepted")
    for idx, contact in enumerate(pending, start=1):
        print(f"\n[{idx}/{len(pending)}] Checking invite status: {contact.name} ({contact.linkedin_url})")
        try:
            accepted_now = invite_is_accepted(
                contact.linkedin_url,
                profile_load_delay_sec=profile_load_delay_sec,
                safari_activate_window=safari_activate_window,
            )
        except Exception as err:
            verify_failed += 1
            print(f"⚠️ Could not verify {contact.name}: {err}")
            continue

        if not accepted_now:
            still_pending += 1
            print("  -> Still pending (no Notion change)")
            continue

        print("  -> Invite accepted")
        accepted.append(contact)

        if args.dry_run:
            print(f"  -> Dry run: would mark Notion status '{invite_accepted_status}'")
        else:
            try:
                workflow.update_status(contact, invite_accepted_status)
                print(f"  -> Marked Notion status: {invite_accepted_status}")
            except Exception as err:
                print(f"⚠️ Could not mark '{invite_accepted_status}' for {contact.name}: {err}")

    if accepted:
        print(f"\nPhase 1 complete. Accepted contacts: {len(accepted)}")
    else:
        print("\nPhase 1 complete. No accepted contacts found.")

    if accepted:
        print("\nPhase 2/3: Review messages (no sending yet)")
    for idx, contact in enumerate(accepted, start=1):
        message = _render_message(template_text, contact)
        print(f"\n[{idx}/{len(accepted)}] Prepared message for {contact.name}:")
        print("-----")
        print(message)
        print("-----")

        approval = input("Approve this message for later send? [y/N]: ").strip().lower()
        if approval != "y":
            print("  -> Not approved.")
            continue
        approved.append(contact)
        approved_message_by_page_id[contact.page_id] = message
        print("  -> Approved for send queue.")

    if approved:
        print(f"\nPhase 2 complete. Approved messages: {len(approved)}")
    elif accepted:
        print("\nPhase 2 complete. No messages were approved.")

    if args.dry_run:
        print("\nSummary:")
        print(f"- Pending checked: {len(pending)}")
        print(f"- Still pending: {still_pending}")
        print(f"- Invite accepted: {len(accepted)}")
        print(f"- Verification failures: {verify_failed}")
        print(f"- Messages approved: {len(approved)}")
        print("- Messages sent: 0 (dry run)")
        if approved:
            print("\nDry run contacts approved for send queue:")
            for contact in approved:
                print(f"- {contact.name}")
        return

    if approved:
        print("\nPhase 3/3: Sending approved messages")
    for idx, contact in enumerate(approved, start=1):
        message = approved_message_by_page_id.get(contact.page_id, "")
        if not message:
            print(f"⚠️ Missing approved message text for {contact.name}; skipping.")
            continue

        print(f"\n[{idx}/{len(approved)}] Sending message to {contact.name} ({contact.linkedin_url})")
        if auto_send_messages:
            try:
                was_sent = _auto_send_linkedin_message_safari(
                    contact.linkedin_url,
                    message,
                    profile_load_delay_sec=profile_load_delay_sec,
                    compose_load_delay_sec=compose_load_delay_sec,
                    activate_window=safari_activate_window,
                )
            except Exception as err:
                print(f"⚠️ Auto-send failed for {contact.name}: {err}")
                continue
            if not was_sent:
                print(f"⚠️ Auto-send did not complete for {contact.name} (missing Message or Send control).")
                continue
            print("Message sent automatically.")
        else:
            if not args.no_browser:
                _open_profile_in_browser(contact.linkedin_url)
            copied = _copy_to_clipboard(message)
            if copied:
                print("Message copied to clipboard. Please paste and send in browser.")
            else:
                print("Could not copy to clipboard automatically. Copy from terminal output.")

            confirmation = input("After you manually send in browser, type 'sent' to confirm: ").strip().lower()
            if confirmation != "sent":
                print("Not marked as sent.")
                continue

        sent.append(contact)
        try:
            workflow.update_status(contact, initial_reachout_status)
            print(f"  -> Marked Notion status: {initial_reachout_status}")
        except Exception as err:
            print(f"⚠️ Could not mark '{initial_reachout_status}' for {contact.name}: {err}")

    print("\nSummary:")
    print(f"- Pending checked: {len(pending)}")
    print(f"- Still pending: {still_pending}")
    print(f"- Invite accepted: {len(accepted)}")
    print(f"- Verification failures: {verify_failed}")
    print(f"- Messages approved: {len(approved)}")
    print(f"- Messages sent: {len(sent)}")

    if not sent:
        print("No messages confirmed as sent.")
        return

    print("\nMessages sent and confirmed:")
    for contact in sent:
        print(f"- {contact.name}")


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser(
        description=(
            "Check pending LinkedIn invites from Notion contacts, mark accepted invites, "
            "then run a queued 2-step messaging flow: approve messages first, then batch-send approved ones."
        )
    )
    parser.add_argument(
        "--only-url",
        default="",
        help=(
            "Optional single LinkedIn URL filter. "
            "When omitted, all pending contacts are processed."
        ),
    )
    parser.add_argument(
        "--template",
        default="",
        help=(
            "Path to message template file using {first_name}, {name}, {linkedin_url}. "
            "Defaults to LINKEDIN_MESSAGE_TEMPLATE_PATH from .env, then message_template.txt."
        ),
    )
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Do not auto-open profile pages in browser (manual-send mode only).",
    )
    parser.add_argument(
        "--manual-send",
        action="store_true",
        help="Disable Safari auto-send and require manual send confirmation.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do everything except Notion status updates.",
    )
    args = parser.parse_args()
    run(args)


if __name__ == "__main__":
    main()
