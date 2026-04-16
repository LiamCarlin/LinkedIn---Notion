import argparse
import email
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, cast
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from notion_client import Client


EXPECTED_FIELDS = [
    "Name",
    "Company",
    "Email",
    "LinkedIn",
    "Role",
    "Last Connected",
    "Next follow up",
    "Notes",
    "Status",
]

FIELD_ALIASES = {
    "Name": ["name"],
    "Company": ["company", "compay"],
    "Email": ["email"],
    "LinkedIn": ["linkedin"],
    "Role": ["role"],
    "Last Connected": ["lastconnected", "lastcontacted"],
    "Next follow up": ["nextfollowup", "nextfollow-up"],
    "Notes": ["notes"],
    "Status": ["status"],
}


@dataclass
class Contact:
    name: str
    company: str
    email: str
    linkedin: str
    role: str
    status: str


@dataclass
class ContactReview:
    contact: Contact


def normalize_linkedin_url(url: str) -> str:
    url = url.strip()
    if not url:
        raise ValueError("Empty URL")

    if not re.match(r"^https?://", url, re.IGNORECASE):
        url = f"https://{url}"

    parsed = urlparse(url)
    if "linkedin.com" not in parsed.netloc.lower():
        raise ValueError("This is not a LinkedIn URL")

    path = parsed.path.rstrip("/")
    if not path.startswith("/in/"):
        raise ValueError("Please paste a LinkedIn personal profile URL (linkedin.com/in/...) ")

    return f"https://www.linkedin.com{path}/"


def normalize_notion_database_id(raw_value: str) -> str:
    value = (raw_value or "").strip()
    if not value:
        raise RuntimeError("Missing Notion database ID")

    # Accept full Notion URLs, IDs with hyphens, and accidental query strings like ?v=...
    value = value.split("?", 1)[0].split("#", 1)[0]

    url_match = re.search(r"([0-9a-fA-F]{32})(?![0-9a-fA-F])", value)
    if url_match:
        return url_match.group(1)

    cleaned = value.replace("-", "")
    if re.fullmatch(r"[0-9a-fA-F]{32}", cleaned):
        return cleaned

    raise RuntimeError(
        "NOTION_CONTACTS_DATABASE_ID is invalid. Use the 32-character database ID (no ?v=...)."
    )


def normalize_field_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (value or "").strip().lower())


def is_local_source(value: str) -> bool:
    lowered = (value or "").strip().lower()
    return os.path.exists(value) and lowered.endswith((".html", ".htm", ".mhtml", ".mht"))


def _name_from_slug(profile_url: str) -> str:
    path = urlparse(profile_url).path.strip("/")
    parts = path.split("/")
    if len(parts) < 2:
        return ""
    slug = parts[1]
    cleaned = re.sub(r"[-_]+", " ", slug).strip()
    tokens = cleaned.split()
    # LinkedIn slugs often end with a numeric suffix; drop trailing numeric-only tokens.
    while tokens and re.fullmatch(r"\d+", tokens[-1]):
        tokens.pop()
    return " ".join(p.capitalize() for p in tokens)


def _split_role_company(headline: str) -> Dict[str, str]:
    text = (headline or "").strip()
    if not text:
        return {"role": "", "company": "", "tagline": ""}

    match = re.search(
        r"^(?P<role>.+?)\s+at\s+(?P<company>[^|]+?)(?:\s*\|\s*(?P<tagline>.+))?$",
        text,
        re.IGNORECASE,
    )
    if match:
        return {
            "role": match.group("role").strip(" -|"),
            "company": match.group("company").strip(" -|"),
            "tagline": (match.group("tagline") or "").strip(),
        }

    match = re.search(
        r"^(?P<role>.+?)\s+@\s+(?P<company>[^|·]+?)(?:\s*\|\s*(?P<tagline>.+))?$",
        text,
        re.IGNORECASE,
    )
    if match:
        return {
            "role": text,
            "company": match.group("company").strip(" -|"),
            "tagline": (match.group("tagline") or "").strip(),
        }

    # Example: "SVP - Industrial Design I SharkNinja"
    match = re.search(r"^(?P<role>.+?)\s+I\s+(?P<company>[^|·]+)$", text)
    if match:
        return {
            "role": text,
            "company": match.group("company").strip(" -|"),
            "tagline": "",
        }

    # Example: "Head of Design | SharkNinja" (avoid matching long taglines after '|')
    match = re.search(r"^(?P<role>[^|]+?)\s*\|\s*(?P<company>[^|,]{2,60})$", text)
    if match:
        return {
            "role": text,
            "company": match.group("company").strip(" -|"),
            "tagline": "",
        }

    return {"role": text, "company": "", "tagline": ""}


def _clean_full_name(value: str) -> str:
    name = (value or "").strip()
    if not name:
        return ""

    name = re.sub(r"\s*\|\s*LinkedIn\s*$", "", name, flags=re.IGNORECASE).strip()
    # Remove common non-name suffixes when LinkedIn title/content includes extra descriptors.
    name = re.split(r"\s+[-·|]\s+", name)[0].strip()
    if " at " in name.lower() or " location:" in name.lower() or "connections" in name.lower():
        name = re.split(r"\s+at\s+", name, flags=re.IGNORECASE)[0].strip()
    return name


def _looks_like_profile_headline(candidate: str) -> bool:
    text = (candidate or "").strip()
    lower = text.lower()
    if len(text) < 8 or len(text) > 140:
        return False
    has_signal = any(marker in lower for marker in [" at ", " | ", " i "])
    role_keywords = ["manager", "director", "vp", "svp", "head", "chief", "lead", "design"]
    if not has_signal and not any(keyword in lower for keyword in role_keywords):
        return False
    blocked_fragments = [
        "view ",
        " linkedin",
        "connections",
        "location:",
        "followers",
        "we are looking",
        "positions open",
        "http://",
        "https://",
    ]
    return not any(fragment in lower for fragment in blocked_fragments)


def _looks_like_company(candidate: str, full_name: str, role: str) -> bool:
    text = (candidate or "").strip()
    lower = text.lower()
    if not text or len(text) > 70:
        return False
    if text.startswith("·"):
        return False
    if re.search(r"\b[1-9]\d*(?:st|nd|rd|th)\b", lower):
        return False
    if not re.search(r"[A-Za-z]", text):
        return False
    exact_blocked = {
        "more",
        "message",
        "contact info",
        "english",
        "français",
        "ad options",
    }
    if lower in exact_blocked:
        return False
    blocked_fragments = [
        "linkedin",
        "location",
        "connections",
        "followers",
        "followed by",
        "metropolitan region",
        "try sales navigator",
        " at ",
        "http://",
        "https://",
    ]
    if any(fragment in lower for fragment in blocked_fragments):
        return False
    if lower in {(full_name or "").lower(), (role or "").lower()}:
        return False
    return True


def _looks_like_role(candidate: str, full_name: str) -> bool:
    text = (candidate or "").strip()
    lower = text.lower()
    if not text or len(text) < 6 or len(text) > 140:
        return False
    if text == full_name:
        return False
    blocked = [
        "notifications",
        "for business",
        "try sales navigator",
        "message",
        "connect",
        "contact info",
        "mutual connection",
        "followers",
        "connections",
        "location:",
        "view ",
        " linkedin",
    ]
    if any(item in lower for item in blocked):
        return False
    if text.startswith("·"):
        return False
    if not re.search(r"[A-Za-z]", text):
        return False
    role_markers = [" at ", "|", " I ", "design", "manager", "director", "vp", "svp", "head", "chief"]
    return any(marker.lower() in lower for marker in role_markers)


def _looks_like_full_name(candidate: str) -> bool:
    text = (candidate or "").strip()
    if len(text) < 3 or len(text) > 80:
        return False
    if re.search(r"\d", text):
        return False
    if any(ch in text for ch in "@|/"):
        return False
    if "," in text:
        return False
    tokens = text.split()
    if len(tokens) < 2 or len(tokens) > 5:
        return False
    if not all(re.fullmatch(r"[A-Za-z][A-Za-z'`.-]*", token) for token in tokens):
        return False
    return True


def _profile_text_candidates_near_name(soup: BeautifulSoup, full_name: str) -> List[str]:
    if not full_name:
        return []

    headings = soup.find_all(["h1", "h2", "h3"])
    name_heading = None
    for heading in headings:
        heading_text = _clean_full_name(heading.get_text(" ", strip=True))
        if heading_text == full_name:
            name_heading = heading
            break

    if not name_heading:
        return []

    nearby: List[str] = []
    for element in name_heading.find_all_next(["p", "span", "div"], limit=20):
        text = re.sub(r"\s+", " ", element.get_text(" ", strip=True)).strip()
        if not text:
            continue
        if text == full_name:
            continue
        if len(text) > 160:
            continue
        nearby.append(text)
    return nearby


def _extract_structured_profile_fields(soup: BeautifulSoup, full_name: str) -> Dict[str, str]:
    if not full_name:
        return {"headline": "", "company": ""}

    anchors = []
    for tag in soup.find_all(["h1", "h2", "h3", "p", "span"]):
        tag_text = _clean_full_name(tag.get_text(" ", strip=True))
        if tag_text == full_name:
            anchors.append(tag)

    if not anchors:
        return {"headline": "", "company": ""}

    best = {"headline": "", "company": ""}
    best_score = -1

    for anchor in anchors[:8]:
        headline = ""
        company = ""
        nearby_p_texts: List[str] = []
        score = 0

        for p_tag in anchor.find_all_next("p", limit=25):
            text = re.sub(r"\s+", " ", p_tag.get_text(" ", strip=True)).strip()
            if not text or text == full_name or len(text) > 160:
                continue
            lower = text.lower()
            # Stop before profile subsections where school/history strings are common.
            if any(marker in lower for marker in ["contact info", "followers", "connections"]):
                break
            nearby_p_texts.append(text)

            classes = p_tag.get("class") or []
            if not isinstance(classes, list):
                classes = [str(classes)]
            class_set = {str(item) for item in classes}

            # Observed LinkedIn export structure: this class commonly marks role/headline.
            if "c8a8c952" in class_set and _looks_like_role(text, full_name):
                headline = text
                score += 2
                split = _split_role_company(text)
                split_company = (split.get("company") or "").strip()
                if split_company and _looks_like_company(split_company, full_name, headline):
                    company = split_company
                    score += 3

            # Observed LinkedIn export structure: this class commonly marks company row.
            if "a91650dc" in class_set and _looks_like_company(text, full_name, headline):
                company = text
                score += 3

            # Fallback class seen in some exports for text rows that may include company.
            if not company and "_4feb9671" in class_set and _looks_like_company(text, full_name, headline):
                company = text
                score += 1

            if headline and company and score >= 4:
                break

        # Generic fallback if class markers are absent in a given export.
        if not headline or not company:
            for text in nearby_p_texts:
                if not headline and _looks_like_role(text, full_name):
                    headline = text
                    score += 1
                if not company and _looks_like_company(text, full_name, headline):
                    company = text
                    score += 1
                if headline and company:
                    break

        if score > best_score:
            best = {"headline": headline, "company": company}
            best_score = score

    return best


def _extract_from_profile_card(text_nodes: List[str], full_name: str) -> Dict[str, str]:
    if not full_name:
        return {"headline": "", "company": ""}

    name_idx = None
    for idx, value in enumerate(text_nodes[:120]):
        if value.strip() == full_name:
            name_idx = idx
            break

    if name_idx is None:
        return {"headline": "", "company": ""}

    window = text_nodes[name_idx + 1 : name_idx + 14]
    headline = ""
    company = ""
    company_from_headline = False

    for value in window:
        if not headline and _looks_like_role(value, full_name):
            headline = value
            split = _split_role_company(value)
            if split.get("company") and not company:
                company = split["company"]
                company_from_headline = True
            continue

        if _looks_like_company(value, full_name, headline):
            if not company or company_from_headline:
                company = value
                company_from_headline = False
            if headline and company:
                break

    return {"headline": headline, "company": company}


def _company_from_profile_summary(text: str) -> str:
    value = (text or "").strip()
    if not value:
        return ""

    # Example:
    # "Experience: SharkNinja · Education: ..."
    match = re.search(r"Experience:\s*([^·|]+)", value, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    match = re.search(r"\bat\s+([^|·]+)", value, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    return ""


def _clean_role(value: str) -> str:
    role = (value or "").strip()
    if not role:
        return ""

    # LinkedIn sometimes provides a long profile summary instead of role.
    if role.lower().startswith("experience:") and "location:" in role.lower():
        return ""
    return role


def _decode_mhtml_html(path: str) -> str:
    with open(path, "rb") as fh:
        message = email.message_from_binary_file(fh)

    html_parts: List[str] = []
    for part in message.walk():
        content_type = part.get_content_type()
        if content_type != "text/html":
            continue

        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        if not isinstance(payload, (bytes, bytearray)):
            continue

        charset = part.get_content_charset() or "utf-8"
        try:
            html_parts.append(bytes(payload).decode(charset, errors="replace"))
        except Exception:
            html_parts.append(bytes(payload).decode("utf-8", errors="replace"))

    if html_parts:
        return "\n".join(html_parts)

    raise RuntimeError(f"Could not extract HTML from MHTML file: {path}")


def _source_url_from_mhtml(path: str) -> str:
    with open(path, "rb") as fh:
        message = email.message_from_binary_file(fh)

    for header_name in ("Snapshot-Content-Location", "Content-Location"):
        value = message.get(header_name)
        if value:
            return value.strip()
    return ""


def _source_url_from_html(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    og_url = soup.find("meta", attrs={"property": "og:url"})
    if og_url:
        content = og_url.get("content")
        if isinstance(content, list):
            content = " ".join(str(x) for x in content)
        if isinstance(content, str) and content.strip():
            return content.strip()

    canonical = soup.find("link", attrs={"rel": re.compile(r"\bcanonical\b", re.I)})
    if canonical:
        href = canonical.get("href")
        if isinstance(href, list):
            href = " ".join(str(x) for x in href)
        if isinstance(href, str) and href.strip():
            return href.strip()

    return ""


def _extract_linkedin_url_from_html(html: str) -> str:
    pattern = re.compile(r"https?://(?:www\.)?linkedin\.com/in/[A-Za-z0-9\-_%]+/?", re.IGNORECASE)
    match = pattern.search(html)
    if not match:
        return ""

    try:
        return normalize_linkedin_url(match.group(0))
    except Exception:
        return ""


def _read_html_source(source: str) -> str:
    if is_local_source(source):
        if source.lower().endswith((".mhtml", ".mht")):
            return _decode_mhtml_html(source)
        with open(source, "r", encoding="utf-8", errors="replace") as fh:
            return fh.read()

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    }
    response = requests.get(source, headers=headers, timeout=30)
    if response.status_code >= 400:
        if response.status_code == 999:
            raise RuntimeError(
                "LinkedIn blocked automated access (999). I can still save this contact if you provide missing fields manually or use a saved HTML/MHTML export."
            )
        raise RuntimeError(f"Could not read LinkedIn profile page ({response.status_code}).")
    return response.text


def _linkedin_slug_from_url(url: str) -> str:
    path = urlparse(url).path.strip("/")
    parts = path.split("/")
    if len(parts) >= 2 and parts[0] == "in":
        return parts[1].lower()
    return ""


def _find_matching_export_for_url(profile_url: str) -> str:
    slug = _linkedin_slug_from_url(profile_url)
    if not slug:
        return ""

    downloads = Path.home() / "Downloads"
    if not downloads.exists():
        return ""

    candidates: List[Path] = []
    for pattern in ("*.mhtml", "*.mht", "*.html", "*.htm"):
        candidates.extend(downloads.glob(pattern))

    # Most recent first.
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    for path in candidates[:200]:
        try:
            if path.suffix.lower() in {".mhtml", ".mht"}:
                source_url = _source_url_from_mhtml(str(path)).lower()
                if slug in source_url:
                    return str(path)
            else:
                text = path.read_text(encoding="utf-8", errors="replace")
                source_url = _source_url_from_html(text).lower()
                if slug in source_url:
                    return str(path)
        except Exception:
            continue

    return ""


def _safe_export_filename(profile_url: str) -> str:
    slug = _linkedin_slug_from_url(profile_url) or "linkedin-profile"
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "-", slug).strip("-") or "linkedin-profile"
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{slug}-{stamp}.html"


def _auto_capture_html_with_safari(profile_url: str) -> str:
    if sys.platform != "darwin":
        return ""

    escaped_url = profile_url.replace("\\", "\\\\").replace('"', '\\"')
    script = (
        f'set targetUrl to "{escaped_url}"\n'
        'tell application "Safari"\n'
        "  activate\n"
        "  if (count of windows) = 0 then\n"
        "    make new document\n"
        "  end if\n"
        "  set URL of front document to targetUrl\n"
        "  delay 8\n"
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
            timeout=45,
        )
    except Exception:
        return ""

    if result.returncode != 0:
        return ""

    html = (result.stdout or "").strip()
    if len(html) < 300:
        return ""

    exports_dir = Path.home() / "Downloads"
    exports_dir.mkdir(parents=True, exist_ok=True)
    output_path = exports_dir / _safe_export_filename(profile_url)

    try:
        output_path.write_text(html, encoding="utf-8")
    except Exception:
        return ""

    return str(output_path)


def _auto_capture_linkedin_export(profile_url: str) -> str:
    return _auto_capture_html_with_safari(profile_url)


def extract_profile_from_html(html: str, profile_url: str = "") -> Dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")
    text_nodes = [re.sub(r"\s+", " ", t).strip() for t in soup.stripped_strings]
    text_nodes = [t for t in text_nodes if t]
    resolved_profile_url = profile_url or _extract_linkedin_url_from_html(html)

    full_name = ""
    headline = ""
    company = ""
    email = ""

    # Try JSON-LD first
    for tag in soup.find_all("script", attrs={"type": "application/ld+json"}):
        try:
            data = json.loads(tag.get_text(strip=True))
        except Exception:
            continue

        candidates = data if isinstance(data, list) else [data]
        for item in candidates:
            if not isinstance(item, dict):
                continue
            if item.get("@type") == "Person":
                full_name = (item.get("name") or full_name).strip()
                headline = (item.get("jobTitle") or headline).strip()
                works_for = item.get("worksFor")
                if isinstance(works_for, dict):
                    company = (works_for.get("name") or company).strip()

    if not full_name:
        title_tag = soup.find("title")
        if title_tag and title_tag.get_text(strip=True):
            title_text = title_tag.get_text(strip=True)
            full_name = _clean_full_name(title_text)

    if not full_name:
        og_title = soup.find("meta", attrs={"property": "og:title"})
        if og_title:
            content = og_title.get("content")
            if isinstance(content, list):
                content = " ".join(str(x) for x in content)
            if isinstance(content, str) and content.strip():
                full_name = _clean_full_name(content.split("|")[0].strip())

    if not full_name:
        for heading in soup.find_all(["h1", "h2", "h3"]):
            candidate = _clean_full_name(heading.get_text(" ", strip=True))
            if _looks_like_full_name(candidate):
                full_name = candidate
                break

    if not headline:
        og_desc = soup.find("meta", attrs={"property": "og:description"})
        if og_desc:
            content = og_desc.get("content")
            if isinstance(content, list):
                content = " ".join(str(x) for x in content)
            if isinstance(content, str) and content.strip():
                headline = content.strip()

    # Search visible text nodes for the most informative line, such as:
    # "VP, Industrial Design at SharkNinja | Design & Product Development Leader"
    if not headline:
        for candidate in text_nodes:
            if _looks_like_profile_headline(candidate):
                headline = candidate
                break

    if not full_name:
        full_name = _name_from_slug(resolved_profile_url)
    full_name = _clean_full_name(full_name)

    if full_name:
        structured_fields = _extract_structured_profile_fields(soup, full_name)
        if not headline and structured_fields.get("headline"):
            headline = structured_fields["headline"]
        if not company and structured_fields.get("company"):
            company = structured_fields["company"]

        nearby_text = _profile_text_candidates_near_name(soup, full_name)
        for candidate in nearby_text:
            if not headline and _looks_like_role(candidate, full_name):
                headline = candidate
            if not company and _looks_like_company(candidate, full_name, headline):
                company = candidate
            if headline and company:
                break

    top_card = _extract_from_profile_card(text_nodes, full_name)
    if not headline and top_card.get("headline"):
        headline = top_card["headline"]
    if not company and top_card.get("company"):
        company = top_card["company"]

    split = _split_role_company(headline)
    role = split["role"]
    if not company:
        company = split["company"]

    # If the company is still missing, look near the discovered headline in visible text.
    if not company:
        for idx, candidate in enumerate(text_nodes):
            if candidate != headline:
                continue
            for nearby in text_nodes[idx + 1 : idx + 6]:
                if _looks_like_company(nearby, full_name, role):
                    company = nearby
                    break
            if company:
                break

    mailto_match = re.search(r"mailto:([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})", html, re.IGNORECASE)
    if mailto_match:
        email = mailto_match.group(1).strip()

    return {
        "full_name": full_name,
        "headline": role,
        "company": company,
        "email": email,
        "linkedin_url": resolved_profile_url,
    }


def to_contact(profile: Dict[str, Any], linkedin_url: str, default_status: str) -> Contact:
    full_name = (profile.get("full_name") or "").strip()
    role = _clean_role((profile.get("occupation") or profile.get("headline") or "").strip())
    company = (profile.get("company") or "").strip()
    resolved_linkedin_url = (linkedin_url or profile.get("linkedin_url") or "").strip()

    if not company:
        company = _company_from_profile_summary((profile.get("headline") or "").strip())

    if company and not role:
        # Keep role editable in CLI but avoid writing noisy summary text by default.
        role = ""

    return Contact(
        name=full_name,
        company=company,
        email=(profile.get("email") or "").strip(),
        linkedin=resolved_linkedin_url,
        role=role,
        status=default_status,
    )


class NotionContactWriter:
    def __init__(self, api_key: str, contacts_database_id: str):
        self.client = Client(auth=api_key)
        self.contacts_database_id = contacts_database_id
        try:
            self.contacts_db = self.client.databases.retrieve(database_id=contacts_database_id)
        except Exception as err:
            raise RuntimeError(
                "Could not open Contacts database in Notion. "
                "Check NOTION_CONTACTS_DATABASE_ID and make sure your integration is shared with the database. "
                f"Details: {err}"
            ) from err

        if not isinstance(self.contacts_db, dict) or self.contacts_db.get("object") != "database":
            raise RuntimeError(
                "NOTION_CONTACTS_DATABASE_ID did not resolve to a Notion database. "
                "Use the database's 32-char ID (not a page ID/view URL) and share it with the integration."
            )

        data_sources = self.contacts_db.get("data_sources") or []
        if not data_sources:
            raise RuntimeError(
                "No data sources were found for the Contacts database. Open the database in Notion and ensure it has a data source."
            )

        self.contacts_data_source_id = data_sources[0]["id"]

        try:
            self.contacts_data_source = cast(Dict[str, Any], self.client.data_sources.retrieve(
                data_source_id=self.contacts_data_source_id
            ))
        except Exception as err:
            raise RuntimeError(
                "Could not read database properties from Notion. Confirm integration access to the Contacts database."
            ) from err

        props = cast(Dict[str, Any], self.contacts_data_source).get("properties")
        if not isinstance(props, dict):
            raise RuntimeError(
                "Could not read database properties from Notion. "
                "Confirm integration access to the Contacts database."
            )
        self.contacts_props = props

    def _title_prop_name(self, database: Dict[str, Any]) -> Optional[str]:
        for prop_name, prop_def in database.get("properties", {}).items():
            if prop_def.get("type") == "title":
                return prop_name
        return None

    def _find_or_create_company_page(self, relation_data_source_id: str, company_name: str) -> Optional[str]:
        if not company_name:
            return None

        try:
            company_ds = cast(Dict[str, Any], self.client.data_sources.retrieve(data_source_id=relation_data_source_id))
        except Exception:
            # If integration cannot access related Companies DB, skip relation linking silently.
            return None

        title_prop = self._title_prop_name(company_ds)
        if not title_prop:
            return None

        query = cast(Dict[str, Any], self.client.data_sources.query(
            data_source_id=relation_data_source_id,
            filter={
                "property": title_prop,
                "title": {
                    "equals": company_name,
                },
            },
            page_size=1,
        ))

        if query.get("results"):
            return query["results"][0]["id"]

        created = cast(Dict[str, Any], self.client.pages.create(
            parent={"data_source_id": relation_data_source_id},
            properties={
                title_prop: {
                    "title": [{"text": {"content": company_name}}],
                }
            },
        ))
        return created["id"]

    def _set_text_like(self, prop_type: str, value: str) -> Dict[str, Any]:
        if prop_type == "title":
            return {"title": [{"text": {"content": value}}]} if value else {"title": []}
        if prop_type == "rich_text":
            return {"rich_text": [{"text": {"content": value}}]} if value else {"rich_text": []}
        if prop_type == "email":
            return {"email": value or None}
        if prop_type == "url":
            return {"url": value or None}
        if prop_type == "phone_number":
            return {"phone_number": value or None}
        if prop_type == "select":
            return {"select": {"name": value}} if value else {"select": None}
        if prop_type == "status":
            return {"status": {"name": value}} if value else {"status": None}
        return {}

    def _property_name_or_none(self, expected_name: str) -> Optional[str]:
        expected_norm = normalize_field_key(expected_name)
        aliases = {expected_norm}
        for alias in FIELD_ALIASES.get(expected_name, []):
            aliases.add(normalize_field_key(alias))

        for prop_name in self.contacts_props.keys():
            prop_norm = normalize_field_key(prop_name)
            if prop_norm in aliases:
                return prop_name
        return None

    def create_contact_page(
        self,
        contact: Contact,
        auto_set_last_connected: bool,
        last_connected: str,
        next_follow_up: str,
        notes: str,
    ) -> Dict[str, Any]:
        properties: Dict[str, Any] = {}

        values = {
            "Name": contact.name,
            "Company": contact.company,
            "Email": contact.email,
            "LinkedIn": contact.linkedin,
            "Role": contact.role,
            "Status": contact.status,
            "Notes": notes,
            "Next follow up": next_follow_up,
        }

        for expected in EXPECTED_FIELDS:
            prop_name = self._property_name_or_none(expected)
            if not prop_name:
                continue

            prop_def = self.contacts_props[prop_name]
            prop_type = prop_def["type"]

            if expected == "Company" and prop_type == "relation":
                relation_data_source_id = prop_def["relation"].get("data_source_id") or prop_def["relation"].get("database_id")
                company_page_id = self._find_or_create_company_page(
                    relation_data_source_id,
                    contact.company,
                )
                if company_page_id:
                    properties[prop_name] = {"relation": [{"id": company_page_id}]}
                continue

            if expected == "Last Connected":
                if prop_type == "date":
                    if last_connected:
                        properties[prop_name] = {"date": {"start": last_connected}}
                    elif auto_set_last_connected:
                        properties[prop_name] = {"date": {"start": date.today().isoformat()}}
                continue

            if expected == "Next follow up":
                if prop_type == "date":
                    if next_follow_up:
                        properties[prop_name] = {"date": {"start": next_follow_up}}
                    continue

            value = values.get(expected, "")
            payload = self._set_text_like(prop_type, value)
            if payload:
                properties[prop_name] = payload

        # Fallback: if Name field wasn't found by name, use the database title property
        if not any(
            self.contacts_props[p]["type"] == "title" and p in properties
            for p in self.contacts_props
        ):
            for prop_name, prop_def in self.contacts_props.items():
                if prop_def.get("type") == "title":
                    properties[prop_name] = {
                        "title": [{"text": {"content": contact.name or "Unknown"}}]
                    }
                    break

        page = cast(Dict[str, Any], self.client.pages.create(
            parent={"data_source_id": self.contacts_data_source_id},
            properties=properties,
        ))
        return page


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value


def prompt_with_default(label: str, current: str) -> str:
    suffix = f" [{current}]" if current else ""
    entered = input(f"{label}{suffix}: ").strip()
    return entered if entered else current


def review_contact_fields(contact: Contact, auto_set_last_connected: bool) -> ContactReview:
    print("ℹ️ Review fields (press Enter to keep suggested value).")
    reviewed = Contact(
        name=prompt_with_default("Name", contact.name),
        company=prompt_with_default("Company", contact.company),
        email=prompt_with_default("Email", contact.email),
        linkedin=prompt_with_default("LinkedIn", contact.linkedin),
        role=prompt_with_default("Role", contact.role),
        status=contact.status,
    )
    return ContactReview(
        contact=reviewed,
    )


def import_one_profile(url: str) -> None:
    notion_api_key = required_env("NOTION_API_KEY")
    notion_contacts_db = normalize_notion_database_id(required_env("NOTION_CONTACTS_DATABASE_ID"))

    default_status = os.getenv("CONTACT_DEFAULT_STATUS", "Connected").strip() or "Connected"
    auto_set_last_connected = (
        os.getenv("AUTO_SET_LAST_CONNECTED", "true").strip().lower() in {"1", "true", "yes", "y"}
    )

    linkedin_url = ""
    try:
        if is_local_source(url):
            html = _read_html_source(url)
            source_url = (
                _source_url_from_mhtml(url)
                if url.lower().endswith((".mhtml", ".mht"))
                else _source_url_from_html(html)
            )
            if not source_url:
                source_url = _extract_linkedin_url_from_html(html)
            linkedin_url = normalize_linkedin_url(source_url) if source_url else ""
            profile = extract_profile_from_html(html, linkedin_url)
        else:
            linkedin_url = normalize_linkedin_url(url)
            try:
                html = _read_html_source(linkedin_url)
                profile = extract_profile_from_html(html, linkedin_url)
            except Exception as err:
                if "(999)" in str(err):
                    fallback_path = _find_matching_export_for_url(linkedin_url)
                    used_auto_capture = False
                    if not fallback_path:
                        fallback_path = _auto_capture_linkedin_export(linkedin_url)
                        used_auto_capture = bool(fallback_path)
                    if fallback_path:
                        if used_auto_capture:
                            print(f"⚠️ LinkedIn blocked access (999). Auto-captured local export: {fallback_path}")
                        else:
                            print(f"⚠️ LinkedIn blocked access (999). Using local export: {fallback_path}")
                        html = _read_html_source(fallback_path)
                        source_url = (
                            _source_url_from_mhtml(fallback_path)
                            if fallback_path.lower().endswith((".mhtml", ".mht"))
                            else _source_url_from_html(html)
                        )
                        if not source_url:
                            source_url = _extract_linkedin_url_from_html(html)
                        linkedin_url = normalize_linkedin_url(source_url) if source_url else linkedin_url
                        profile = extract_profile_from_html(html, linkedin_url)
                    else:
                        raise RuntimeError(
                            "LinkedIn blocked automated access (999), and automatic local export failed. "
                            "Please open the profile in Safari while logged in to LinkedIn and run again, "
                            "or pass a saved HTML/MHTML file path directly."
                        ) from err
                else:
                    raise
    except Exception as err:
        print(f"⚠️ {err}")
        profile = {
            "full_name": _name_from_slug(linkedin_url),
            "headline": "",
            "company": "",
            "email": "",
            "linkedin_url": linkedin_url,
        }

    contact = to_contact(profile, linkedin_url, default_status)
    review = review_contact_fields(contact, auto_set_last_connected)

    writer = NotionContactWriter(notion_api_key, notion_contacts_db)
    page = writer.create_contact_page(
        review.contact,
        auto_set_last_connected,
        date.today().isoformat(),
        "",
        "",
    )

    print("✅ Added contact to Notion")
    print(f"Name: {review.contact.name or 'Unknown'}")
    print(f"Company: {review.contact.company or 'Unknown'}")
    print(f"Role: {review.contact.role or 'Unknown'}")
    print(f"Notion page ID: {page['id']}")


def main() -> None:
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Paste a LinkedIn profile URL or a saved LinkedIn HTML/MHTML file and create a Notion contact automatically."
    )
    parser.add_argument("source", nargs="?", help="LinkedIn profile URL or local .html/.mhtml file")
    args = parser.parse_args()

    if args.source:
        import_one_profile(args.source)
        return

    print("LinkedIn -> Notion importer")
    print("Paste LinkedIn URLs or local HTML/MHTML file paths. Press Enter on an empty line to exit.\n")

    while True:
        url = input("LinkedIn URL: ").strip()
        if not url:
            print("Done.")
            break

        try:
            import_one_profile(url)
        except Exception as err:
            print(f"❌ {err}")
        print()


if __name__ == "__main__":
    main()
