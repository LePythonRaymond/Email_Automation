import streamlit as st
import pandas as pd
import smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import os
from typing import Dict, List, Tuple, Optional
import json
from pathlib import Path

import time
import threading
import collections
import base64
import random
import html2text
import requests
from io import BytesIO
from datetime import datetime, timedelta

# Pure helpers, deliberately free of streamlit and of any import back
# into this module, so an import cycle is structurally impossible and
# both can be tested without starting a Streamlit script run.
import notion_props
import sync_state
import unsubscribe_inbox


# Page configuration
st.set_page_config(
    page_title="MERCI RAYMOND - Raymongraphe",
    page_icon="🌱",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #2E7D32;
        text-align: center;
        margin-bottom: 2rem;
    }
    .step-header {
        font-size: 1.5rem;
        color: #388E3C;
        margin: 1rem 0;
    }
    .email-preview {
        background-color: #F1F8E9;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #4CAF50;
    }
    .success-box {
        background-color: #E8F5E8;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #4CAF50;
    }
    .warning-box {
        background-color: #FFF3E0;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #FF9800;
    }
    .info-box {
        background-color: #E3F2FD;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #2196F3;
        margin: 1rem 0;
    }
    .anti-spam-box {
        background-color: #FFF8E1;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #FFC107;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Gmail SMTP Configuration (hardcoded)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# --- Deliverability hardening (PRD 2026-05) -------------------------------
# Random delay between two sends (seconds). Reduces robotic sending pattern.
MIN_DELAY_S = 1
MAX_DELAY_S = 10

# Daily cap is an unofficial team rule, not enforced in code. The team aims
# for a soft limit (e.g. ~100 sends/sender/day) — no JSON counter file.

# Inbox that receives unsubscribe requests. Must be aliased to a real mailbox
# someone reads (or auto-forwarded), otherwise the List-Unsubscribe header
# becomes a dead end and hurts reputation. Single source of truth: change
# this one line and the header, footer, and detection marker all update.
UNSUBSCRIBE_MAILTO = "desinscription@merciraymond.fr"

# Footer HTML appended to every outgoing email (visible "unsubscribe" link).
# Marker is used by _ensure_unsubscribe_footer() to avoid duplicating the
# footer on emails that already contain it (e.g. hand-edited HTML).
UNSUBSCRIBE_FOOTER_MARKER = f"mailto:{UNSUBSCRIBE_MAILTO}"
UNSUBSCRIBE_FOOTER_HTML = (
    '<div style="margin-top:24px; padding-top:12px; '
    'border-top:1px solid #eee; font-size:11px; color:#999;">\n'
    '  Vous recevez cet email de MERCI RAYMOND. '
    'Pour ne plus en recevoir, '
    f'<a href="mailto:{UNSUBSCRIBE_MAILTO}?subject=unsubscribe" '
    'style="color:#999;">cliquez ici pour vous désinscrire</a>.\n'
    '</div>'
)
# --------------------------------------------------------------------------

def _read_secret_or_env(key: str) -> str:
    """Resolve a secret: st.secrets > os.environ > .env file. Returns '' if not found."""
    try:
        return str(st.secrets[key])
    except (KeyError, AttributeError, FileNotFoundError, Exception):
        pass
    val = os.environ.get(key, "")
    if val:
        return val
    env_path = Path(__file__).parent / ".env"
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            if line.startswith(f"{key}="):
                return line.split("=", 1)[1].strip().strip('"')
    return ""


# Notion API Configuration
NOTION_API_KEY = _read_secret_or_env("NOTION_API_KEY")
OPENAI_API_KEY = _read_secret_or_env("OPENAI_API_KEY")
NOTION_DS_ID = "285d9278-02d7-808a-9395-000b04dfc654"
# ID of the Notion database (or data source) holding the suppression list.
# Plug it in via .streamlit/secrets.toml under SUPPRESSION_NOTION_DS_ID — accepts
# either the database ID or the data source ID, the code auto-detects.
SUPPRESSION_NOTION_DS_ID = _read_secret_or_env("SUPPRESSION_NOTION_DS_ID")
NOTION_API_VERSION_LEGACY = "2022-06-28"
NOTION_API_VERSION_DS = "2025-09-03"

# --- Unsubscribe mailbox sync (F4) — configuration --------------------------
# desinscription@merciraymond.fr is an ALIAS: it redirects into a real person's
# mailbox, and an alias cannot be authenticated over IMAP. So UNSUB_IMAP_USER
# names the mailbox we actually log into, and every message is filtered on
# Delivered-To so we never look at the rest of that person's mail.
#
# The password is looked up in st.secrets["users"] first, by email: the
# `password` field there is already the Gmail app password used by
# server.login() for SMTP, and a Gmail app password works for IMAP too. So in
# the common case (the alias points at one of the senders) the only new secret
# needed is UNSUB_IMAP_USER. UNSUB_IMAP_PASSWORD is the fallback for a mailbox
# that is not one of the configured senders.
UNSUB_IMAP_USER = _read_secret_or_env("UNSUB_IMAP_USER")
UNSUB_IMAP_PASSWORD = _read_secret_or_env("UNSUB_IMAP_PASSWORD")
UNSUB_IMAP_HOST = _read_secret_or_env("UNSUB_IMAP_HOST") or "imap.gmail.com"
# Override the Delivered-To filter. Left empty it is derived automatically:
# UNSUBSCRIBE_MAILTO when we log into a different mailbox (the alias case),
# nothing when we log into the unsubscribe mailbox itself (a dedicated box,
# where all the mail concerns us anyway).
UNSUB_IMAP_FILTER = _read_secret_or_env("UNSUB_IMAP_FILTER")
# Kill switches: flip a secret to stop the feature without a redeploy.
UNSUB_SYNC_ENABLED = (_read_secret_or_env("UNSUB_SYNC_ENABLED") or "1") != "0"
# Reading the senders' own mailboxes: hard bounces (5.x.x, auto-suppressed) and
# "remove me" wording in a reply body (flagged for a human, never automatic).
UNSUB_SCAN_BOUNCES = (_read_secret_or_env("UNSUB_SCAN_BOUNCES") or "1") != "0"


def to_name_case(s: str) -> str:
    """Format a name in proper case: first letter of each part capitalized, rest lower (e.g. JEAN-PIERRE -> Jean-Pierre, françois moreau -> François Moreau)."""
    if not s or not s.strip():
        return s.strip() if s else s
    parts = s.strip().split()
    result = []
    for word in parts:
        subparts = word.split("-")
        capped = []
        for p in subparts:
            if not p:
                continue
            if len(p) == 1:
                capped.append(p.upper())
            else:
                capped.append(p[0].upper() + p[1:].lower())
        result.append("-".join(capped))
    return " ".join(result)


def fetch_notion_sites() -> List[Dict[str, str]]:
    """Fetch all sites from the Notion 'site d'entretien' data source.
    Returns a list of dicts: [{"site": "...", "email": "..."}, ...]
    Handles pagination and API version fallback (data source vs legacy database).
    """
    headers_ds = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Notion-Version": NOTION_API_VERSION_DS,
        "Content-Type": "application/json",
    }
    headers_legacy = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Notion-Version": NOTION_API_VERSION_LEGACY,
        "Content-Type": "application/json",
    }

    def _extract_sites_from_results(results: list) -> List[Dict[str, str]]:
        """Extract site name and email from Notion page results.
        The 'mail rapport et autre' property is rich_text type (not email type).
        """
        sites = []
        for page in results:
            props = page.get("properties", {})
            site_name = None
            email_value = None

            # Extract title (site name) - find by type
            for prop_name, prop_data in props.items():
                if prop_data.get("type") == "title":
                    title_items = prop_data.get("title", [])
                    if title_items:
                        site_name = "".join(item.get("plain_text", "") for item in title_items)
                    break

            # Extract email from "mail rapport et autre" (rich_text type)
            mail_prop = props.get("mail rapport et autre", {})
            if mail_prop.get("type") == "rich_text":
                rt_items = mail_prop.get("rich_text", [])
                if rt_items:
                    email_value = "".join(item.get("plain_text", "") for item in rt_items).strip()

            # Fallback: try "email" type property if rich_text didn't work
            if not email_value:
                for prop_name, prop_data in props.items():
                    if prop_data.get("type") == "email" and prop_data.get("email"):
                        email_value = prop_data["email"]
                        break

            if site_name:
                sites.append({"site": site_name, "email": email_value or ""})
        return sites

    def _query_with_pagination(url: str, headers: dict) -> List[Dict[str, str]]:
        """Query a Notion endpoint with pagination and extract sites."""
        all_sites = []
        start_cursor = None
        while True:
            body = {}
            if start_cursor:
                body["start_cursor"] = start_cursor
            resp = requests.post(url, headers=headers, json=body, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            all_sites.extend(_extract_sites_from_results(data.get("results", [])))
            if data.get("has_more"):
                start_cursor = data.get("next_cursor")
            else:
                break
        return all_sites

    # Try data source endpoint first
    ds_url = f"https://api.notion.com/v1/data_sources/{NOTION_DS_ID}/query"
    try:
        return _query_with_pagination(ds_url, headers_ds)
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code == 400:
            pass  # Fall through to legacy
        else:
            raise

    # Fallback: try as legacy database
    db_url = f"https://api.notion.com/v1/databases/{NOTION_DS_ID}/query"
    try:
        return _query_with_pagination(db_url, headers_legacy)
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code == 400:
            pass  # Fall through to data source discovery
        else:
            raise

    # Last fallback: discover data source ID from database metadata
    meta_url = f"https://api.notion.com/v1/databases/{NOTION_DS_ID}"
    resp = requests.get(meta_url, headers=headers_ds, timeout=30)
    resp.raise_for_status()
    meta = resp.json()
    data_sources = meta.get("data_sources", [])
    if not data_sources:
        raise ValueError("Impossible de trouver la source de donnees Notion. Verifiez l'ID.")
    real_ds_id = data_sources[0].get("id")
    real_ds_url = f"https://api.notion.com/v1/data_sources/{real_ds_id}/query"
    return _query_with_pagination(real_ds_url, headers_ds)


class EmailAutomation:
    def __init__(self):
        # Base email template for TEXT format (includes greetings and signature)
        self.base_email_content_text = ""

        # Base email template for Gmail-style HTML format (unified with header and footer)
        self.base_email_content_html = ""

        # Gmail-style HTML template that looks like plain text - simplified for unified content
        # Use div (not p) so block elements like <ul> from convert_markdown_to_html stay valid;
        # <p>…<ul>… is invalid HTML and causes uneven spacing in email clients.
        self.html_template = (
            """<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.4; color: #202124; background: #ffffff; margin: 0; padding: 0;">
  <div style="margin: 0 0 16px 0;">
    {first_paragraph}
  </div>

  {decorative_image_section}

  <div style="margin: 0 0 16px 0;">
    {second_paragraph}
  </div>

  {logo_section}

  """ + UNSUBSCRIBE_FOOTER_HTML + """
</div>"""
        )
        # Decorative image size options: label -> CSS max-width value
        self.decorative_image_sizes = {
            "Petit (280px)": "280px",
            "Moyen (480px)": "480px",
            "Grand (640px)": "640px",
            "Pleine largeur (100%)": "100%",
        }
        # Get OpenAI API key from secrets
 #       try:
#            self.openai_api_key = st.secrets["api_key"]
  ##             self.openai_api_key = None
    ##   except:
      #      self.openai_api_key = None
       #     st.warning("⚠️ Fichier secrets.toml manquant. Personnalisation simple uniquement.")

    def detect_column_mapping(self, df: pd.DataFrame) -> Dict[str, any]:
        """
        Dynamically detect all columns and identify email column.
        Returns email column name and all available placeholders.
        """
        columns = df.columns.tolist()

        # Email detection patterns
        email_patterns = [
            r'email', r'e-mail', r'mail', r'contact.*client.*1', r'email.*1',
            r'adresse.*mail', r'contact.*mail', r'email.*address', r'electronic.*mail'
        ]

        # Full name detection patterns
        full_name_patterns = [
            r'name', r'nom', r'full.*name', r'nom.*complet', r'contact.*name',
            r'client.*name', r'utilisateur', r'user.*name', r'prenom.*nom'
        ]

        # Find email column
        email_column = None
        best_score = 0

        for col in columns:
            col_lower = col.lower().strip()
            for pattern in email_patterns:
                if re.search(pattern, col_lower):
                    # Score based on pattern match quality
                    score = len(pattern) / len(col_lower) if col_lower else 0
                    if score > best_score:
                        best_score = score
                        email_column = col

        # Find full name columns
        full_name_columns = []
        for col in columns:
            col_lower = col.lower().strip()
            for pattern in full_name_patterns:
                if re.search(pattern, col_lower) and col != email_column:
                    # Check if column contains full names (has spaces)
                    sample_values = df[col].dropna().astype(str).head(5)
                    if any(' ' in str(val) for val in sample_values):
                        full_name_columns.append(col)
                        break

        # Create available placeholders (all columns except email)
        available_placeholders = {}
        for col in columns:
            if col != email_column:
                available_placeholders[col] = col

                # Add first name and last name placeholders for full name columns
                if col in full_name_columns:
                    available_placeholders[f"{col}_first"] = f"{col}_first"
                    available_placeholders[f"{col}_last"] = f"{col}_last"

        # Columns that hold a first name or surname only (not full "Nom Prénom" in one cell)
        name_value_patterns = [
            r'prenom',
            r'prénom',
            r'first\s*name',
            r'firstname',
            r'given\s*name',
            r'forename',
            r'nom de famille',
            r'last\s*name',
            r'lastname',
            r'surname',
            r'family\s*name',
            r'surnom',
            r'nom\s*2',
            r'prenom\s*2',
            r'prénom\s*2',
        ]
        name_value_columns = []
        for col in columns:
            if col == email_column or col in full_name_columns:
                continue
            col_lower = col.lower().strip()
            if col_lower == 'nom':
                name_value_columns.append(col)
                continue
            for pattern in name_value_patterns:
                if re.search(pattern, col_lower, re.I):
                    name_value_columns.append(col)
                    break

        return {
            'email_column': email_column,
            'available_placeholders': available_placeholders,
            'full_name_columns': full_name_columns,
            'name_value_columns': name_value_columns,
            'all_columns': columns
        }

    def extract_contact_info(
        self,
        row: pd.Series,
        email_column: str,
        available_placeholders: Dict[str, str],
        full_name_columns: List[str] = None,
        name_value_columns: List[str] = None,
    ) -> Dict[str, str]:
        """Extract all contact information from a row dynamically."""
        info = {}
        name_value_columns = name_value_columns or []

        # Extract email
        if email_column and email_column in row.index:
            email_value = row[email_column]
            if pd.notna(email_value):
                info['email'] = str(email_value).strip()

        # Extract all other columns as placeholders
        for col_name in available_placeholders.keys():
            if col_name in row.index:
                value = row[col_name]
                if pd.notna(value):
                    info[col_name] = str(value).strip()
                else:
                    info[col_name] = ''

        # Extract first name and last name from full name columns (normalized to proper name case)
        if full_name_columns:
            for full_name_col in full_name_columns:
                if full_name_col in row.index and pd.notna(row[full_name_col]):
                    raw = str(row[full_name_col]).strip()
                    full_name = to_name_case(raw)
                    info[full_name_col] = full_name
                    name_parts = full_name.split()

                    # Add first name (first part)
                    if len(name_parts) > 0:
                        info[f"{full_name_col}_first"] = name_parts[0]

                    # Add last name (all parts after first, joined)
                    if len(name_parts) > 1:
                        info[f"{full_name_col}_last"] = " ".join(name_parts[1:])
                    else:
                        info[f"{full_name_col}_last"] = ""

        # Proper case for separate first-name / surname columns (Prénom, Nom, etc.)
        for col in name_value_columns:
            if col in row.index and pd.notna(row[col]):
                raw = str(row[col]).strip()
                if raw:
                    info[col] = to_name_case(raw)

        return info

    def get_valid_emails_from_df(self, df: pd.DataFrame) -> List[Dict[str, str]]:
        """Extract all valid emails from the dataframe with dynamic column detection.
        If a cell contains multiple emails (separated by newlines, semicolons, commas, or spaces),
        each email produces a separate contact entry sharing the same row data.
        """
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

        # Detect column mapping
        mapping = self.detect_column_mapping(df)
        email_column = mapping['email_column']
        available_placeholders = mapping['available_placeholders']
        full_name_columns = mapping['full_name_columns']
        name_value_columns = mapping.get('name_value_columns', [])

        valid_contacts = []

        for idx, row in df.iterrows():
            contact_info = self.extract_contact_info(
                row,
                email_column,
                available_placeholders,
                full_name_columns,
                name_value_columns=name_value_columns,
            )

            # Extract ALL valid emails from the cell (handles multiple emails per cell)
            raw_email = contact_info.get('email', '')
            found_emails = re.findall(email_pattern, raw_email) if raw_email else []

            for email_addr in found_emails:
                # Create a contact entry for each email, sharing the same row data
                contact_data = {
                    'index': idx,
                    'email': email_addr.strip()
                }

                # Add all other fields dynamically
                for key, value in contact_info.items():
                    if key != 'email':
                        contact_data[key] = value

                # Add a default contact name if none exists
                if not any(key.lower() in ['name', 'nom', 'contact', 'contact_name'] for key in contact_data.keys()):
                    contact_data['contact_name'] = 'Contact'

                valid_contacts.append(contact_data)

        # Remove duplicates by email address, keeping first occurrence
        unique_contacts = []
        seen_emails = set()
        duplicates_removed = 0

        for contact in valid_contacts:
            email = contact['email']
            if email not in seen_emails:
                unique_contacts.append(contact)
                seen_emails.add(email)
            else:
                duplicates_removed += 1

        # Store duplicate count for display
        st.session_state.duplicates_removed = duplicates_removed

        # Filter out suppressed addresses so counters stay accurate and we
        # never even consider them downstream. Surface the count separately.
        kept = []
        suppressed_count = 0
        for contact in unique_contacts:
            if is_suppressed(contact.get('email', '')):
                suppressed_count += 1
            else:
                kept.append(contact)
        st.session_state.suppressed_removed = suppressed_count

        return kept

    def encode_image_to_base64(self, image_file) -> Optional[str]:
        """Convert uploaded image to base64 for embedding in HTML"""
        try:
            return base64.b64encode(image_file.getvalue()).decode()
        except:
            return None

    def compress_image(self, image_file, max_width=1600, quality=82):
        """Compress image to reduce file size for email sending"""
        try:
            from PIL import Image
            import io

            # Get original image data
            original_data = image_file.getvalue()
            original_image = Image.open(io.BytesIO(original_data))

            # Calculate new dimensions maintaining aspect ratio
            if original_image.width > max_width:
                ratio = max_width / original_image.width
                new_height = int(original_image.height * ratio)
                original_image = original_image.resize((max_width, new_height), Image.Resampling.LANCZOS)

            # Convert to RGB if necessary (for JPEG)
            if original_image.mode in ('RGBA', 'LA', 'P'):
                # Create white background for transparency
                background = Image.new('RGB', original_image.size, (255, 255, 255))
                if original_image.mode == 'P':
                    original_image = original_image.convert('RGBA')
                background.paste(original_image, mask=original_image.split()[-1] if original_image.mode == 'RGBA' else None)
                original_image = background
            elif original_image.mode != 'RGB':
                original_image = original_image.convert('RGB')

            # Save compressed image to BytesIO
            compressed_buffer = io.BytesIO()
            original_image.save(compressed_buffer, format='JPEG', quality=quality, optimize=True)
            compressed_buffer.seek(0)

            # Create new file-like object with compressed data
            class CompressedImageFile:
                def __init__(self, data, name):
                    self._data = data
                    self.name = name
                    self.type = 'image/jpeg'

                def getvalue(self):
                    return self._data

            return CompressedImageFile(compressed_buffer.getvalue(), image_file.name)

        except Exception as e:
            # If compression fails, return original file
            print(f"Image compression failed: {e}")
            return image_file

    def convert_markdown_to_html(self, text: str) -> str:
        """Convert markdown-style formatting to HTML for email.
        Supports: [text](url) links, **bold**, *italic*, - bullet lists (with nesting),
        {color:name}text{/color}. Also converts remaining \\n to <br> (outside list blocks).
        """
        # 0. Markdown links: [text](https://url) -> <a href="url" target="_blank">text</a>
        # Only http/https URLs are accepted (mailto/tel/etc. are stripped to plain text)
        # to keep clients out of "suspicious link" filters.
        def _replace_link(m):
            label = m.group(1).strip()
            url = m.group(2).strip()
            if not re.match(r'^https?://', url, re.IGNORECASE):
                return label  # drop the unsafe URL, keep the label as plain text
            # Minimal HTML-escape on the URL to stop quote-breaking in href
            safe_url = url.replace('"', '%22').replace("'", '%27')
            return (
                f'<a href="{safe_url}" target="_blank" rel="noopener noreferrer" '
                f'style="color:#1a73e8; text-decoration:underline;">{label}</a>'
            )

        text = re.sub(
            r'\[([^\]\n]+)\]\(([^)\s]+)\)',
            _replace_link,
            text,
        )

        # 1. Convert color syntax: {color:name}text{/color} -> <span style="color:...">text</span>
        # French color names mapped to professional, readable hex values
        color_map = {
            "bleu": "#1a73e8", "rouge": "#c62828", "vert": "#2e7d32",
            "orange": "#e65100", "violet": "#6a1b9a", "sarcelle": "#00838f",
            "marron": "#4e342e", "ardoise": "#37474f", "framboise": "#ad1457",
            "marine": "#1a237e",
            # English names also supported
            "blue": "#1a73e8", "red": "#c62828", "green": "#2e7d32",
            "purple": "#6a1b9a", "teal": "#00838f", "brown": "#4e342e",
            "navy": "#1a237e",
        }

        def _replace_color(m):
            color_name = m.group(1).strip().lower()
            color_value = color_map.get(color_name, m.group(1))  # fallback to raw value (hex codes)
            return f'<span style="color:{color_value}">{m.group(2)}</span>'

        text = re.sub(
            r'\{color:([^}]+)\}(.*?)\{/color\}',
            _replace_color,
            text,
            flags=re.DOTALL
        )

        # 2. Convert bullet lists with nesting support
        # "- item" = top-level, "    - item" (4+ spaces or tab) = nested
        # Blank lines between bullets stay inside the same <ul> but add extra top margin
        # to the following item, so spacing mirrors what the author typed.
        lines = text.split('\n')
        result_parts = []
        list_depth = 0
        list_buffer = []  # accumulate list HTML without newlines
        pending_spacer = False  # a blank line was seen before the next bullet

        def _flush_list():
            nonlocal list_buffer
            if list_buffer:
                result_parts.append(''.join(list_buffer))
                list_buffer = []

        def _line_is_list_item(raw: str) -> bool:
            if not raw.strip():
                return False
            if re.match(r'^(?:    |\t)\s*- (.+)$', raw):
                return True
            stripped = raw.strip()
            return bool(stripped.startswith('- ') and re.match(r'^- (.+)$', stripped))

        for i, line in enumerate(lines):
            nested_match = re.match(r'^(?:    |\t)\s*- (.+)$', line)
            top_match = re.match(r'^- (.+)$', line.strip()) if not nested_match else None

            # Blank line inside a list: keep the list open but flag the next item
            # so it gets extra top margin (honouring the author's spacing intent).
            if not line.strip() and list_depth > 0:
                next_is_bullet = False
                for j in range(i + 1, len(lines)):
                    if not lines[j].strip():
                        continue
                    next_is_bullet = _line_is_list_item(lines[j])
                    break
                if next_is_bullet:
                    pending_spacer = True
                    continue  # don't close the list

            if nested_match:
                item_text = nested_match.group(1)
                if list_depth == 0:
                    list_buffer.append('<ul style="margin: 8px 0; padding-left: 20px;">')
                    list_depth = 1
                if list_depth == 1:
                    list_buffer.append('<ul style="margin: 0; padding-left: 20px; list-style-type: circle;">')
                    list_depth = 2
                top_margin = "14px" if pending_spacer else "0"
                list_buffer.append(f'<li style="margin: {top_margin} 0 4px 0; padding: 0;">{item_text}</li>')
                pending_spacer = False
            elif top_match and line.strip().startswith('- '):
                item_text = top_match.group(1)
                if list_depth == 2:
                    list_buffer.append('</ul>')
                    list_depth = 1
                if list_depth == 0:
                    list_buffer.append('<ul style="margin: 8px 0; padding-left: 20px;">')
                    list_depth = 1
                top_margin = "14px" if pending_spacer else "0"
                list_buffer.append(f'<li style="margin: {top_margin} 0 4px 0; padding: 0;">{item_text}</li>')
                pending_spacer = False
            else:
                pending_spacer = False
                while list_depth > 0:
                    list_buffer.append('</ul>')
                    list_depth -= 1
                _flush_list()
                result_parts.append(line)

        while list_depth > 0:
            list_buffer.append('</ul>')
            list_depth -= 1
        _flush_list()

        text = '\n'.join(result_parts)

        # 3. Convert **bold text** to <strong>bold text</strong>
        text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)

        # 4. Convert *italic text* to <em>italic text</em>
        text = re.sub(r'\*(.*?)\*', r'<em>\1</em>', text)

        # 5. Convert remaining \n to <br>
        text = text.replace('\n', '<br>')

        return text

    def _render_image_block(self, cid: str, image_file, size_css: str,
                            show_placeholder: bool, label: str = "Image") -> str:
        """Render the HTML block for a single inline image (or its placeholder)."""
        img_style = (
            f"max-width: {size_css}; width: 100%; height: auto; "
            "border:0; outline:0; display: block;"
        )
        wrapper_style = (
            f"margin: 16px 0; max-width: {size_css}; width: 100%; "
            "box-sizing: border-box; border: 1px solid #e0e0e0; "
            "border-radius: 4px; overflow: hidden;"
        )
        if image_file:
            return (
                f'\n                <div style="{wrapper_style}">\n'
                f'                <img src="cid:{cid}" alt="{label}" style="{img_style}">\n'
                f'                </div>'
            )
        if show_placeholder:
            return (
                f'\n                <div style="margin: 16px 0; max-width: {size_css}; '
                'width: 100%; min-height: 180px; background: #f5f5f5; '
                'border: 2px dashed #bdbdbd; border-radius: 4px; display: flex; '
                'align-items: center; justify-content: center; color: #757575; '
                'font-size: 13px; box-sizing: border-box;">\n'
                f'                {label} ({size_css})\n'
                f'                </div>'
            )
        return ''

    def personalize_email(self, contact_data: Dict[str, str], email_content: str, use_html: bool = False,
                         logo_file=None, decorative_image_file=None, attachment_files=None, email_subject: str = "",
                         decorative_image_size: str = "100%", show_image_placeholder: bool = False,
                         decorative_image_file_2=None, decorative_image_size_2: str = "100%") -> Tuple[str, str]:
        """
        Dynamic personalization with any column placeholders from Excel data
        Returns: (personalized_content, personalized_subject)
        decorative_image_size: CSS max-width for the main image (e.g. "280px", "100%").
        decorative_image_file_2 / decorative_image_size_2: optional second inline image.
        show_image_placeholder: if True and no image, show a size box in preview.
        """
        # Safety check for email content - use appropriate template based on format
        if email_content is None:
            email_content = self.base_email_content_html if use_html else self.base_email_content_text

        # Start with the email content
        personalized = email_content

        # Replace all placeholders dynamically
        for key, value in contact_data.items():
            if key != 'email' and key != 'index':  # Skip email and index
                placeholder = f"{{{key}}}"
                # Replace with actual value or empty string if missing
                replacement_value = value if value else ""
                personalized = personalized.replace(placeholder, replacement_value)

        # Handle special case for contact name (extract first name)
        contact_name = contact_data.get('contact_name', '')
        if contact_name and len(contact_name.split()) > 1:
            first_name = contact_name.split()[0]
        else:
            first_name = "Madame/Monsieur"

        # Replace {contact_name} if it exists in the template
        personalized = personalized.replace('{contact_name}', first_name)

        # Personalize subject line with same placeholder logic
        personalized_subject = email_subject
        for key, value in contact_data.items():
            if key != 'email' and key != 'index':
                placeholder = f"{{{key}}}"
                replacement_value = value if value else ""
                personalized_subject = personalized_subject.replace(placeholder, replacement_value)
        personalized_subject = personalized_subject.replace('{contact_name}', first_name)

        if use_html:
            # Prepare logo section - small signature-style image
            logo_section = ""
            if logo_file:
                logo_section = f'<img src="cid:logo" alt="Merci Raymond" style="display:inline-block; height:24px; width:auto; border:0; outline:0; vertical-align:baseline;">'

            # Image 1 ({Image}) and image 2 ({Image2}) — independent sizes, separate CIDs.
            has_img1_placeholder = '{Image}' in personalized
            has_img2_placeholder = '{Image2}' in personalized

            img1_html = self._render_image_block(
                cid='decorative_image',
                image_file=decorative_image_file,
                size_css=decorative_image_size,
                show_placeholder=show_image_placeholder,
                label="Image",
            )
            img2_html = self._render_image_block(
                cid='decorative_image_2',
                image_file=decorative_image_file_2,
                size_css=decorative_image_size_2,
                show_placeholder=show_image_placeholder,
                label="Image 2",
            )

            # If a typed {Image}/{Image2} token has NO uploaded image, remove the
            # token AND the blank line it sat on, so it leaves neither literal text
            # nor an empty gap in the email. (When an image IS uploaded, img*_html
            # is non-empty and the token is replaced by the image further below.)
            _img_token_stripped = False
            if has_img1_placeholder and not img1_html:
                personalized = personalized.replace('{Image}', '')
                has_img1_placeholder = False
                _img_token_stripped = True
            if has_img2_placeholder and not img2_html:
                personalized = personalized.replace('{Image2}', '')
                has_img2_placeholder = False
                _img_token_stripped = True
            if _img_token_stripped:
                personalized = re.sub(r'\n{3,}', '\n\n', personalized).strip()

            # Default auto-placement: images appear stacked between paragraphs 1 and 2
            # when their placeholder is NOT used. Each image only contributes if it
            # actually has content (uploaded file or preview placeholder).
            decorative_image_section = ""
            if not has_img1_placeholder:
                decorative_image_section += img1_html
            if not has_img2_placeholder:
                decorative_image_section += img2_html

            # Split content into first and second paragraphs for Gmail-style layout
            paragraphs = personalized.split('\n\n')

            # First paragraph: everything up to the decorative image
            first_paragraph = ""
            second_paragraph = ""

            if len(paragraphs) >= 3:
                # Split after the first two paragraphs for better balance
                first_paragraph = paragraphs[0] + "\n\n" + paragraphs[1]
                second_paragraph = "\n\n".join(paragraphs[2:])
            elif len(paragraphs) >= 2:
                # Split after the first paragraph
                first_paragraph = paragraphs[0]
                second_paragraph = "\n\n".join(paragraphs[1:])
            else:
                # If we can't split naturally, put most content in first paragraph
                first_paragraph = personalized
                second_paragraph = ""

            # Clean up the paragraphs and ensure proper line breaks
            first_paragraph = first_paragraph.strip()
            second_paragraph = second_paragraph.strip()

            # Convert markdown formatting and line breaks to HTML
            first_paragraph = self.convert_markdown_to_html(first_paragraph)
            second_paragraph = self.convert_markdown_to_html(second_paragraph)

            # Substitute {Image} / {Image2} placeholders where the user typed them.
            # Always resolve the token: real image HTML when an image was uploaded,
            # otherwise strip it (img*_html is '' with no upload) so neither a
            # literal "{Image}" nor an empty box ever leaks into the email.
            if has_img1_placeholder:
                first_paragraph = first_paragraph.replace('{Image}', img1_html)
                second_paragraph = second_paragraph.replace('{Image}', img1_html)
            if has_img2_placeholder:
                first_paragraph = first_paragraph.replace('{Image2}', img2_html)
                second_paragraph = second_paragraph.replace('{Image2}', img2_html)

            # Apply Gmail-style HTML template with unified content
            personalized = self.html_template.format(
                first_paragraph=first_paragraph,
                second_paragraph=second_paragraph,
                logo_section=logo_section,
                decorative_image_section=decorative_image_section
            )

        return personalized, personalized_subject

    def personalize_email_with_ai(self, contact_data: Dict[str, str], email_content: str, use_html: bool = False,
                                 logo_file=None, decorative_image_file=None, attachment_files=None, email_subject: str = "",
                                 decorative_image_size: str = "100%", show_image_placeholder: bool = False,
                                 decorative_image_file_2=None, decorative_image_size_2: str = "100%") -> Tuple[str, str]:
        """AI personalization removed - using simple personalization instead"""
        return self.personalize_email(
            contact_data, email_content, use_html, logo_file, decorative_image_file,
            attachment_files, email_subject,
            decorative_image_size=decorative_image_size,
            show_image_placeholder=show_image_placeholder,
            decorative_image_file_2=decorative_image_file_2,
            decorative_image_size_2=decorative_image_size_2,
        )

    def verify_email_content(self, email_content: str) -> Tuple[bool, List[str]]:
        """SUPER SIMPLE verification - only check for curly brace placeholders"""
        issues = []

        # Check for remaining curly brace placeholders - ONLY THESE
        placeholder_patterns = [
            r'\{[^}]*\}',         # {placeholder}
            r'\{,?\}',            # {}, {,}
        ]

        for pattern in placeholder_patterns:
            matches = re.findall(pattern, email_content, re.IGNORECASE)
            if matches:
                issues.extend([f"Placeholder trouvé: {match}" for match in matches])

        # Basic check - email not empty
        if not email_content.strip():
            issues.append("Email vide")

        return len(issues) == 0, issues

def calculate_sending_time(num_emails: int, delay_seconds: int) -> str:
    """Calculate total sending time"""
    total_seconds = num_emails * delay_seconds
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60

    if hours > 0:
        return f"{hours}h {minutes}min"
    else:
        return f"{minutes}min"


def _users_file_path() -> Path:
    """Path to users.json in project root (same directory as this script)."""
    return Path(__file__).resolve().parent / "users.json"


def _load_users_from_secrets() -> List[Dict[str, str]]:
    """Load senders from st.secrets['users'] (.streamlit/secrets.toml). Returns [] if missing."""
    try:
        secrets_users = st.secrets["users"]
    except (KeyError, AttributeError, FileNotFoundError, Exception):
        return []
    out = []
    try:
        for _key, val in dict(secrets_users).items():
            name = str(val.get("name", "")).strip() if hasattr(val, "get") else ""
            email = str(val.get("email", "")).strip() if hasattr(val, "get") else ""
            password = str(val.get("password", "")).strip() if hasattr(val, "get") else ""
            if name and email:
                out.append({"name": name, "email": email, "password": password})
    except Exception:
        return []
    return out


def _load_users() -> List[Dict[str, str]]:
    """Load senders: from st.secrets first, then users.json (no duplicates by email).
    Returns [] if neither source is configured — the caller surfaces a friendly warning."""
    base = _load_users_from_secrets()
    path = _users_file_path()
    if not path.exists():
        return base
    try:
        with open(path, "r", encoding="utf-8") as f:
            extra = json.load(f)
    except (json.JSONDecodeError, OSError):
        return base
    if not isinstance(extra, list):
        return base
    seen = {u["email"].strip().lower() for u in base}
    out = list(base)
    for u in extra:
        if not isinstance(u, dict):
            continue
        name = (u.get("name") or "").strip()
        email = (u.get("email") or "").strip()
        password = (u.get("password") or "").strip()
        if not name or not email:
            continue
        if email.lower() in seen:
            continue
        seen.add(email.lower())
        out.append({"name": name, "email": email, "password": password})
    return out


# --- Suppression list (F3) -------------------------------------------------
# Persistent list of addresses we must never email again. Source of truth: a
# Notion database whose ID is plugged in via SUPPRESSION_NOTION_DS_ID
# (secrets.toml). The local filesystem is ephemeral on Streamlit Cloud, so no
# local JSON persistence — a small in-memory cache (TTL 5 min) keeps
# is_suppressed() fast across the send loops without hammering Notion.
#
# The schema is DISCOVERED at runtime, never hardcoded. History: the write path
# used to POST a "Raison" property that does not exist in the live database
# (whose real columns are Email/title, Date/date, Nom/rich_text). Notion rejects
# any unknown property name with HTTP 400 validation_error, so every automatic
# write failed and the list could only be filled by typing into Notion. The
# error was invisible because only the requests exception was shown, never the
# response body — which names the offending property in plain text.
#
# Two consequences of building the payload from the live schema:
#   * it works with the database exactly as it is today;
#   * the day someone adds a "Raison" column, it fills itself within the 5 min
#     TTL, with no code change and no redeploy.
# Anything the schema cannot hold (the reason, today) goes into the page BODY
# via notion_props.build_provenance_children(), which is validated against no
# schema and therefore can never be refused.

_SUPPRESSION_CACHE_TTL_S = 300   # 5 minutes
_SUPPRESSION_ERROR_TTL_S = 60    # after a failed fetch, wait this long (circuit breaker)
_SUPPRESSION_STALE_WARN_S = 900  # beyond this age, warn before sending
_NOTION_RESOLVE_TIMEOUT_S = 5    # per probe, and there are two of them
_NOTION_RESOLVE_FAIL_TTL_S = 120  # negative cache on data-source resolution

# Held in sync_state, NOT here: Streamlit re-executes this script's whole top
# level in a fresh namespace on every rerun (verified: three reruns, three
# different namespaces), so a dict defined here would be recreated on every
# click and the TTL and circuit breaker below could never hold across reruns.
# An imported module lives in sys.modules and is created once per process.
_suppression_state = sync_state.suppression

# The cache is mutated from the IMAP worker thread and read on every rerun.
# The GIL makes that non-corrupting, but the writes below are copy-on-write
# under this lock, which costs nothing at 15-200 entries and removes the
# question entirely.
_suppression_lock = sync_state.suppression_lock


def _notion_headers(version: Optional[str] = None) -> dict:
    return {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Notion-Version": version or NOTION_API_VERSION_DS,
        "Content-Type": "application/json",
    }


def _resolve_suppression_ds_id() -> Optional[str]:
    """Return a working data source ID for the suppression DB.

    Accepts either a data source ID or a database ID in SUPPRESSION_NOTION_DS_ID
    and figures out which one it is.

    The negative cache matters as much as the positive one: this used to record
    successes only, so on failure BOTH probes were replayed on every call. The
    read path's circuit breaker does not cover this function, and the sidebar
    paste loop calls it once per address — with a Notion endpoint that hangs
    rather than 404s, pasting 20 addresses meant 20 x 2 x 10 s of frozen rerun.
    """
    if not SUPPRESSION_NOTION_DS_ID or not NOTION_API_KEY:
        return None
    if _suppression_state["resolved_ds_id"]:
        return _suppression_state["resolved_ds_id"]
    failed_at = _suppression_state["resolve_failed_at"]
    if failed_at and (time.time() - failed_at) < _NOTION_RESOLVE_FAIL_TTL_S:
        return None

    headers_ds = _notion_headers()

    # Probe 1: treat the value as a data source ID.
    try:
        probe = requests.post(
            f"https://api.notion.com/v1/data_sources/{SUPPRESSION_NOTION_DS_ID}/query",
            headers=headers_ds, json={"page_size": 1},
            timeout=_NOTION_RESOLVE_TIMEOUT_S,
        )
        if probe.status_code < 400:
            _suppression_state["resolved_ds_id"] = SUPPRESSION_NOTION_DS_ID
            _suppression_state["resolve_failed_at"] = 0.0
            return SUPPRESSION_NOTION_DS_ID
    except requests.RequestException:
        pass

    # Probe 2: treat the value as a database ID and discover its data source.
    try:
        meta = requests.get(
            f"https://api.notion.com/v1/databases/{SUPPRESSION_NOTION_DS_ID}",
            headers=headers_ds, timeout=_NOTION_RESOLVE_TIMEOUT_S,
        )
        meta.raise_for_status()
        data_sources = meta.json().get("data_sources", [])
        if data_sources:
            ds_id = data_sources[0].get("id")
            if ds_id:
                _suppression_state["resolved_ds_id"] = ds_id
                _suppression_state["resolve_failed_at"] = 0.0
                return ds_id
    except requests.RequestException:
        pass

    _suppression_state["resolve_failed_at"] = time.time()
    return None


def _fetch_suppression_from_notion() -> dict:
    """Query Notion for the full suppression list. Returns {email_lc: meta}.
    Raises on network / auth / schema errors so the caller can fall back to
    the last known good cache."""
    ds_id = _resolve_suppression_ds_id()
    if not ds_id:
        # This message is the first thing the operator reads during an outage,
        # and it goes straight into the blocking banner, so it names the actual
        # cause rather than saying "not configured" about a secret that is set.
        # The June 2026 incident was exactly this: a correct ID on a database
        # that had been un-shared from the integration.
        raise RuntimeError(
            "base de désinscription injoignable — vérifiez que la base est "
            "toujours partagée avec l'intégration Notion (Notion → base → ••• "
            "→ Connexions) et que SUPPRESSION_NOTION_DS_ID est correct"
        )

    headers = _notion_headers()
    url = f"https://api.notion.com/v1/data_sources/{ds_id}/query"

    out = {}
    schema = None
    start_cursor = None
    while True:
        body = {"page_size": 100}
        if start_cursor:
            body["start_cursor"] = start_cursor
        resp = requests.post(url, headers=headers, json=body, timeout=10)
        resp.raise_for_status()
        payload = resp.json()
        results = payload.get("results", [])
        # A Notion PAGE carries every property of its database, empty ones
        # included, so the full schema comes for free with the rows we were
        # fetching anyway — no extra request on the hot path. Re-derived on
        # every refresh, so a column added in Notion is picked up within the TTL.
        if results and schema is None:
            schema = notion_props.schema_from_page(results[0])
        for page in results:
            props = page.get("properties", {})
            email_lc = None
            # The email lives in the title property (whatever its name).
            for _name, pdata in props.items():
                if pdata.get("type") == "title":
                    title_items = pdata.get("title", [])
                    if title_items:
                        email_lc = "".join(
                            it.get("plain_text", "") for it in title_items
                        ).strip().lower()
                    break
            if not email_lc:
                continue
            # Optional Date property.
            date_val = ""
            date_prop = props.get("Date") or {}
            if date_prop.get("type") == "date":
                date_val = (date_prop.get("date") or {}).get("start") or ""
            # Optional Raison property (rich_text or select).
            reason_val = "manual_unsubscribe"
            for reason_key in ("Raison", "Reason"):
                pdata = props.get(reason_key) or {}
                if pdata.get("type") == "rich_text":
                    items = pdata.get("rich_text", [])
                    if items:
                        reason_val = "".join(
                            it.get("plain_text", "") for it in items
                        ).strip() or reason_val
                    break
                if pdata.get("type") == "select":
                    sel = pdata.get("select") or {}
                    if sel.get("name"):
                        reason_val = sel["name"]
                    break
            out[email_lc] = {"date": date_val, "reason": reason_val}
        if payload.get("has_more"):
            start_cursor = payload.get("next_cursor")
        else:
            break
    if schema:
        _suppression_state["schema"] = schema
        _suppression_state["schema_at"] = time.time()
    return out


def _suppression_schema() -> dict:
    """{property_name: property_type} for the suppression database.

    Normally free: derived from the rows _load_suppression() already fetched.
    The explicit GET is only needed when the database is empty — no rows means
    no page, means no schema to read off one.
    """
    if _suppression_state["schema"]:
        return _suppression_state["schema"]
    _load_suppression()
    if _suppression_state["schema"]:
        return _suppression_state["schema"]
    ds_id = _resolve_suppression_ds_id()
    if not ds_id:
        return {}
    try:
        resp = requests.get(
            f"https://api.notion.com/v1/data_sources/{ds_id}",
            headers=_notion_headers(), timeout=10,
        )
        resp.raise_for_status()
        _suppression_state["schema"] = notion_props.schema_from_data_source(resp.json())
        _suppression_state["schema_at"] = time.time()
    except requests.RequestException:
        pass
    return _suppression_state["schema"] or {}


def _post_suppression_to_notion(
    email_lc: str,
    reason: str,
    nom: str = "",
    provenance: Optional[List[str]] = None,
    _retry: bool = True,
) -> bool:
    """Create a page in the suppression DB. Returns True on success.

    No st.* calls in here: this runs from the IMAP worker thread, which has no
    Streamlit script context. Failures are recorded in
    _suppression_state["last_write_error"] — INCLUDING Notion's response body,
    which is the part that was missing and that made the "Raison" bug invisible.
    """
    ds_id = _resolve_suppression_ds_id()
    if not ds_id:
        _suppression_state["last_write_error"] = (
            "data source Notion non résolu (secret absent, ou base non partagée "
            "avec l'intégration)"
        )
        return False

    schema = _suppression_schema()
    try:
        props = notion_props.build_suppression_properties(
            schema, email_lc, reason=reason,
            day=datetime.now().date().isoformat(), nom=nom,
        )
    except notion_props.NotionSchemaError:
        # Last resort: the title alone, under the historical name. Better a
        # bare row than a lost unsubscribe request.
        props = {"Email": {"title": [{"text": {"content": email_lc}}]}}

    body = {
        "parent": {"type": "data_source_id", "data_source_id": ds_id},
        "properties": props,
    }
    if provenance:
        body["children"] = notion_props.build_provenance_children(provenance)

    try:
        resp = requests.post(
            "https://api.notion.com/v1/pages",
            headers=_notion_headers(), json=body, timeout=15,
        )
    except requests.RequestException as exc:
        _suppression_state["last_write_error"] = f"réseau: {exc}"
        return False

    if resp.status_code == 429 and _retry:
        try:
            time.sleep(min(5.0, float(resp.headers.get("Retry-After", 1)) + 0.2))
        except (TypeError, ValueError):
            time.sleep(1.2)
        return _post_suppression_to_notion(email_lc, reason, nom, provenance,
                                           _retry=False)

    if resp.status_code >= 400:
        try:
            code = (resp.json() or {}).get("code", "")
        except ValueError:
            code = ""
        _suppression_state["last_write_error"] = (
            f"HTTP {resp.status_code} {code} — {resp.text[:300]}"
        )
        if code == "validation_error" and _retry:
            # The cached schema is stale (a column was renamed or removed):
            # drop it, re-read it, retry once.
            _suppression_state["schema"] = None
            return _post_suppression_to_notion(email_lc, reason, nom, provenance,
                                               _retry=False)
        if code in ("object_not_found", "invalid_request_url") and _retry:
            # Legacy parent — with the RESOLVED ds_id, not the raw secret. The
            # old code passed SUPPRESSION_NOTION_DS_ID here, so whenever the
            # secret held a data source id this fallback was structurally wrong.
            legacy = dict(body)
            legacy["parent"] = {"type": "database_id", "database_id": ds_id}
            legacy.pop("children", None)
            try:
                retry = requests.post(
                    "https://api.notion.com/v1/pages",
                    headers=_notion_headers(NOTION_API_VERSION_LEGACY),
                    json=legacy, timeout=15,
                )
            except requests.RequestException as exc:
                _suppression_state["last_write_error"] = f"legacy réseau: {exc}"
                return False
            if retry.status_code < 400:
                _suppression_state["last_write_error"] = ""
                return True
            _suppression_state["last_write_error"] = (
                f"legacy HTTP {retry.status_code} — {retry.text[:300]}"
            )
        return False

    _suppression_state["last_write_error"] = ""
    return True


def _load_suppression(now: Optional[float] = None) -> dict:
    """Return the suppression dict. Re-fetches from Notion when the in-memory
    cache is stale. On Notion failure, returns the last known good copy and
    backs off (circuit breaker) so a broken/unreachable Notion DB can never
    cause per-contact / per-rerun retry storms.

    `now` is injectable (seconds, time.time() scale) purely so the three states
    below are testable without patching the clock.
    """
    now = time.time() if now is None else now

    # Circuit breaker: if a fetch failed recently, serve the last known good
    # copy (or {} if we never loaded) WITHOUT touching Notion. This is what
    # stops the per-contact / per-rerun hammering when Notion is unreachable
    # (e.g. the DB gets un-shared from the integration -> repeated 404s).
    if _suppression_state["last_error"] and (
        now - _suppression_state["last_error"]
    ) < _SUPPRESSION_ERROR_TTL_S:
        return _suppression_state["data"]

    fresh_enough = (
        _suppression_state["loaded_ever"]
        and (now - _suppression_state["last_fetch"]) < _SUPPRESSION_CACHE_TTL_S
    )
    if fresh_enough:
        return _suppression_state["data"]
    if not SUPPRESSION_NOTION_DS_ID:
        # Not configured yet — return empty so the app keeps working. The gate
        # below is what makes sure nobody sends blind because of it.
        return {}
    try:
        fresh = _fetch_suppression_from_notion()
        with _suppression_lock:
            _suppression_state["data"] = fresh
        _suppression_state["last_fetch"] = now
        _suppression_state["loaded_ever"] = True
        _suppression_state["last_error"] = 0.0   # clear the breaker on success
        _suppression_state["last_error_msg"] = ""
        return fresh
    except Exception as exc:
        # Stale-while-error: keep serving the last known good copy if any, and
        # arm the breaker so we don't retry on every call for the next window.
        # The message is kept because a silent degradation is what let the
        # June 2026 outage look like a legitimately empty list.
        _suppression_state["last_error"] = now
        _suppression_state["last_error_msg"] = f"{exc.__class__.__name__}: {exc}"[:300]
        return _suppression_state["data"]


def is_suppressed(email: str) -> bool:
    if not email or not isinstance(email, str):
        return False
    return email.strip().lower() in _load_suppression()


def _remember_suppression(email_lc: str, reason: str) -> None:
    """Copy-on-write cache insert, so a rerun iterating the dict never sees it
    mutate underneath."""
    with _suppression_lock:
        fresh = dict(_suppression_state["data"])
        fresh[email_lc] = {
            "date": datetime.now().date().isoformat(),
            "reason": reason,
        }
        _suppression_state["data"] = fresh


def add_suppression(
    email: str,
    reason: str = "manual_unsubscribe",
    nom: str = "",
    provenance: Optional[List[str]] = None,
) -> bool:
    """Push one address to the Notion suppression DB. Returns True on success.

    Returns True for an address already on the list without writing anything:
    this guard is THE idempotency mechanism of the automatic sync, and it also
    stops the manual paste from creating a second Notion page for an address
    that is already there.
    """
    if not SUPPRESSION_NOTION_DS_ID:
        _suppression_state["last_write_error"] = (
            "SUPPRESSION_NOTION_DS_ID non configuré dans secrets.toml"
        )
        return False
    if not email or not email.strip():
        return False
    email_lc = email.strip().lower()
    if email_lc in _load_suppression():
        return True
    ok = _post_suppression_to_notion(email_lc, reason, nom=nom,
                                     provenance=provenance)
    if ok:
        # Update the local cache. We deliberately do NOT reset last_fetch: the
        # cache is already correct, and forcing a refetch meant 20 additions
        # cost 20 POSTs plus 20 full paginated reads — which the IMAP sync,
        # adding several addresses at once, would turn into the bottleneck.
        _remember_suppression(email_lc, reason)
        _suppression_state["last_error"] = 0.0   # Notion answered: clear breaker
    return ok


def _suppression_gate(state: dict, now: Optional[float] = None,
                      configured: Optional[bool] = None) -> Tuple[str, str]:
    """("ok" | "warn" | "block", message). Pure: state, clock and the
    "is the secret set" flag are all injected, so every branch below is
    reachable from a test without touching secrets.toml.

    The asymmetry to remember: Notion is the source of truth for filtering, so
    a Notion outage BLOCKS sending. IMAP is an ingestion convenience, so an
    IMAP outage only warns.

    Blocking the "never loaded" case is not blocking a campaign "by mistake":
    that state means we know nothing at all about the list, so is_suppressed()
    returns False for everybody and the counter reads "0 adresse(s)" exactly
    like a legitimately empty list. The block is precisely correlated with
    ignorance.
    """
    now = time.time() if now is None else now
    configured = bool(SUPPRESSION_NOTION_DS_ID) if configured is None else configured
    if not configured:
        return "warn", (
            "⚠️ `SUPPRESSION_NOTION_DS_ID` n'est pas configuré : **aucun filtrage "
            "de désinscription n'est appliqué**. Renseignez le secret, ou cochez "
            "la case ci-dessous pour envoyer quand même."
        )
    if not state.get("loaded_ever"):
        why = state.get("last_error_msg") or "cause inconnue"
        return "block", (
            "⛔ La liste de désinscription n'a **jamais pu être chargée** depuis "
            f"Notion ({why}). Envoyer maintenant écrirait à des personnes qui ont "
            "demandé le contraire. Vérifiez que la base est toujours partagée avec "
            "l'intégration « Rapport d'entretien », puis réessayez."
        )
    age = now - state.get("last_fetch", 0.0)
    if state.get("last_error"):
        why = state.get("last_error_msg") or "cause inconnue"
        return "warn", (
            f"⚠️ Notion est injoignable ({why}). Filtrage effectué sur la dernière "
            f"liste connue, vieille de {int(age // 60)} min."
        )
    if age > _SUPPRESSION_STALE_WARN_S:
        return "warn", (
            f"⚠️ La liste de désinscription date de {int(age // 60)} min. "
            "Elle sera rafraîchie au prochain chargement."
        )
    return "ok", ""


# --- Unsubscribe mailbox sync (F4) -----------------------------------------
# Turns the desinscription@ mailbox (and the senders' own mailboxes) into
# writes on the Notion suppression list. The parsing and the IMAP I/O live in
# unsubscribe_inbox.py; what follows is only the wiring: when to run, how to
# never run twice at once, and what to do with the verdicts.
#
# Streamlit Cloud has no scheduler, no worker and no Redis, so:
#   * the work always runs in a DAEMON THREAD, never on the rerun path. That is
#     the property which makes a repeat of the June 2026 grey-screen freeze
#     structurally impossible: a rerun does Thread.start() and a dict read, and
#     nothing else.
#   * the module-level lock IS the cross-session lock, because every Streamlit
#     session lives in the same Python process. It plays the exact role Redis
#     SETNX played in the sibling project.
#   * there is no cursor and no state file. The Notion list is the idempotency
#     store (add_suppression returns True without writing for an address it
#     already has), so re-scanning the same window is free.
#
# What "automatic" means here, honestly: the container sleeps when nobody uses
# the app, so a scan only happens while someone has it open. An external cron
# cannot help — an HTTP GET on the Streamlit URL serves the static shell and
# never executes the script, which needs a websocket session. But no email can
# be sent while nobody is using the app either (both SMTP loops sit inside
# st.button handlers), so the invariant that actually matters is reachable:
#
#     no send can start without a sync attempt having just happened.
#
# That is what the pre-send gate enforces.

_UNSUB_THROTTLE_S = 180          # minimum gap between two opportunistic cycles
_UNSUB_CYCLE_BUDGET_S = 25.0     # wall clock for a whole cycle, all mailboxes
_UNSUB_MAILBOX_DEADLINE_S = 20.0  # wall clock for one mailbox
_UNSUB_PRESEND_WAIT_S = 8.0      # bounded join() before a send; overrun warns
# 90 days, not 30. This window is only affordable because replay is free
# (Notion is the idempotency store, so re-seeing an old request costs one SEARCH
# and a few hundred bytes and writes nothing) and because the mailbox is quiet:
# measured 16 messages delivered to desinscription@ over 180 days. A short
# window silently drops any request older than itself -- and the app only runs
# when someone opens it, so a fortnight of holidays would lose real requests.
# Measured effect of widening it: a request from 43 days ago that had been typed
# into Notion with a one-character typo, and was therefore not protected at all,
# is now caught automatically.
_UNSUB_WINDOW_DAYS_UNSUB = 90
_UNSUB_WINDOW_DAYS_BOUNCE = 7    # bodies cost more, and a stale bounce is moot
_UNSUB_BACKOFF_S = {
    "auth": 3600,      # bad credentials need a human, not a retry
    "quota": 600,      # Gmail bandwidth throttle
    "mailbox": 3600,
    "config": 3600,
    "network": 300,
    "unknown": 300,
}
_UNSUB_JOURNAL_MAXLEN = sync_state.UNSUB_JOURNAL_MAXLEN
_UNSUB_REVIEW_MAXLEN = 20

# Same reason as _suppression_state above: this must outlive a rerun, so it
# lives in the imported sync_state module. Reset per rerun it would mean one
# IMAP cycle per user click, a throttle that never throttles, a backoff that is
# always forgotten, a review queue that vanishes, and no lock at all between
# two concurrent sessions -- i.e. precisely the storm this design prevents.
_unsub_state = sync_state.unsub
_unsub_lock = sync_state.imap_lock


def _unsub_journal(event: str, **fields) -> None:
    """One JSON line per decision, to stdout AND to an in-memory ring.

    stdout is the only persistent record we have: Streamlit Cloud captures it
    ("Manage app" -> logs) and it survives reruns and container sleep. For an
    irreversible action that is not optional. flush=True because a buffered
    line is lost when the container is recycled.
    """
    entry = {"ts": datetime.now().isoformat(timespec="seconds"), "evt": event}
    entry.update(fields)
    _unsub_state["journal"].append(entry)
    try:
        print(json.dumps(entry, ensure_ascii=False), flush=True)
    except Exception:
        pass


def _own_domains() -> frozenset:
    """Our own domains, derived from the configured senders.

    Computed rather than hardcoded so it can never drift out of sync when a
    sender is added. It is what stops the classifier from ever suppressing one
    of our own addresses.
    """
    domains = set()
    for user in _load_users():
        addr = (user.get("email") or "").strip().lower()
        if "@" in addr:
            domains.add(addr.rsplit("@", 1)[-1])
    if "@" in UNSUBSCRIBE_MAILTO:
        domains.add(UNSUBSCRIBE_MAILTO.rsplit("@", 1)[-1].lower())
    return frozenset(domains)


def _unsub_password_for(mailbox: str) -> str:
    """The Gmail app password for a mailbox: the senders table first."""
    target = (mailbox or "").strip().lower()
    for user in _load_users():
        if (user.get("email") or "").strip().lower() == target:
            if user.get("password"):
                return user["password"]
    if target == (UNSUB_IMAP_USER or "").strip().lower():
        return UNSUB_IMAP_PASSWORD
    return ""


def _unsub_backed_off(mailbox: str, now: float) -> bool:
    return now < _unsub_state["backoff"].get((mailbox or "").lower(), 0.0)


def _unsub_configs(sender_email: Optional[str] = None,
                   now: Optional[float] = None) -> List:
    """The mailboxes to scan in this cycle, already filtered by backoff.

    Round-robin on purpose: scanning all six senders at once would put six
    logins on one cycle. One per cycle, with a 180 s throttle, visits them all
    in about 18 minutes and keeps a cycle at two to four seconds. The pre-send
    path passes sender_email so it scans exactly the mailbox whose bounces
    matter for the campaign about to start.
    """
    now = time.time() if now is None else now
    if not UNSUB_SYNC_ENABLED:
        return []
    own = _own_domains()
    never = frozenset({UNSUBSCRIBE_MAILTO.strip().lower()})
    configs = []

    # Role A: the mailbox behind the desinscription@ alias.
    if UNSUB_IMAP_USER:
        box = UNSUB_IMAP_USER.strip().lower()
        password = _unsub_password_for(box)
        if password and not _unsub_backed_off(box, now):
            explicit = (UNSUB_IMAP_FILTER or "").strip().lower()
            # An alias needs the filter; a dedicated mailbox does not, since
            # everything in it concerns us.
            derived = "" if box == UNSUBSCRIBE_MAILTO.strip().lower() else \
                UNSUBSCRIBE_MAILTO.strip().lower()
            configs.append(unsubscribe_inbox.ImapConfig(
                user=box, password=password, host=UNSUB_IMAP_HOST,
                role=unsubscribe_inbox.ROLE_UNSUB,
                recipient_filter=explicit or derived,
                since_days=_UNSUB_WINDOW_DAYS_UNSUB,
                max_messages=60,
                deadline_s=_UNSUB_MAILBOX_DEADLINE_S,
                own_domains=own, never_suppress=never,
            ))

    # Role B: a sender's own mailbox. Bounces come back to the Return-Path, and
    # Reply-To is the sender, so this is the only place either can be found.
    if UNSUB_SCAN_BOUNCES:
        senders = [u for u in _load_users()
                   if u.get("email") and u.get("password")]
        target = None
        wanted = (sender_email or "").strip().lower()
        if wanted:
            target = next((u for u in senders
                           if u["email"].strip().lower() == wanted), None)
        elif senders:
            # The counter is advanced by the CALLER, once per cycle. It used to
            # be advanced here, and this function is called twice per cycle
            # (once to test whether there is work, once by the worker), so the
            # index moved by two: with six senders that visited indices 0, 2
            # and 4 forever and never swept the other three mailboxes.
            target = senders[_unsub_state["rr"] % len(senders)]
        if target:
            box = target["email"].strip().lower()
            if not _unsub_backed_off(box, now):
                for role in (unsubscribe_inbox.ROLE_BOUNCE,
                             unsubscribe_inbox.ROLE_REPLY):
                    configs.append(unsubscribe_inbox.ImapConfig(
                        user=box, password=target["password"],
                        host=UNSUB_IMAP_HOST, role=role,
                        since_days=_UNSUB_WINDOW_DAYS_BOUNCE,
                        max_messages=80, max_bodies=5,
                        deadline_s=_UNSUB_MAILBOX_DEADLINE_S,
                        own_domains=own, never_suppress=never,
                    ))
    return configs


def _unsub_advance_rr() -> None:
    """Move the round-robin on by one. Called exactly ONCE per cycle, by the
    worker, so every sender's mailbox is eventually visited."""
    senders = [u for u in _load_users() if u.get("email") and u.get("password")]
    if senders:
        _unsub_state["rr"] = (_unsub_state["rr"] + 1) % len(senders)


def _unsub_has_work(now: Optional[float] = None) -> bool:
    """Is there any mailbox we could scan right now? Side-effect free.

    Kept separate from _unsub_configs() precisely because that one used to be
    called for this test, and advanced the round-robin as a side effect.
    """
    now = time.time() if now is None else now
    if not UNSUB_SYNC_ENABLED:
        return False
    if (UNSUB_IMAP_USER and _unsub_password_for(UNSUB_IMAP_USER)
            and not _unsub_backed_off(UNSUB_IMAP_USER, now)):
        return True
    if not UNSUB_SCAN_BOUNCES:
        return False
    return any(u.get("email") and u.get("password")
               and not _unsub_backed_off(u["email"], now)
               for u in _load_users())


def _apply_report(report) -> None:
    """Push a report's verdicts to Notion. Fills report.added / .duplicates.

    The duplicate check is explicit even though add_suppression() also guards,
    because the counts are what prove idempotence to the operator: a second
    scan of the same window must read "0 added".
    """
    for cand in report.suppress:
        if cand.email in _load_suppression():
            report.duplicates += 1
            continue
        ok = add_suppression(
            cand.email, reason=cand.reason, nom=cand.display_name,
            provenance=unsubscribe_inbox.provenance_lines(cand),
        )
        if ok:
            report.added += 1
        _unsub_journal(
            "suppress", email=cand.email, reason=cand.reason, why=cand.why,
            frm=cand.from_addr, subject=cand.subject[:120], uid=cand.uid,
            mailbox=cand.mailbox, notion_ok=ok,
            notion_error=_suppression_state["last_write_error"][:200] if not ok else "",
        )
    for cand in report.review:
        _unsub_journal("review", candidate=cand.review_hint or cand.email,
                       why=cand.why, frm=cand.from_addr,
                       subject=cand.subject[:120], mailbox=cand.mailbox)
    for cand in report.deferred:
        # A 4.x.x deletes nothing. It is logged because a rising deferral rate
        # is the earliest signal that a sender's reputation is degrading.
        _unsub_journal("deferred", email=cand.email, status=cand.dsn_status,
                       why=cand.why, mailbox=cand.mailbox)


def _record_reports(reports: List) -> None:
    """Fold a cycle's reports into the shared state."""
    added = duplicates = examined = 0
    last_error = last_kind = ""
    for report in reports:
        added += report.added
        duplicates += report.duplicates
        examined += report.examined
        box = (report.mailbox or "").lower()
        if report.ok:
            _unsub_state["backoff"].pop(box, None)
            _unsub_state["errors"].pop(box, None)
        else:
            wait = _UNSUB_BACKOFF_S.get(report.error_kind, 300)
            _unsub_state["backoff"][box] = time.time() + wait
            _unsub_state["errors"][box] = f"{report.error_kind}: {report.error}"
            last_error, last_kind = report.error, report.error_kind
            _unsub_journal("imap_error", mailbox=report.mailbox,
                           role=report.role, kind=report.error_kind,
                           error=report.error[:200], backoff_s=wait)
        # Review items are replaced per mailbox+role, never accumulated: the
        # message stays in the mailbox, so an unhandled one simply reappears on
        # the next cycle. That gives us a "not yet dealt with" queue for free,
        # with no durable state anywhere.
        if report.ok:
            key = f"{report.mailbox}|{report.role}"
            _unsub_state["review"][key] = report.review[:_UNSUB_REVIEW_MAXLEN]
    _unsub_state["added_last"] = added
    _unsub_state["duplicates_last"] = duplicates
    _unsub_state["examined_last"] = examined
    _unsub_state["added_total"] += added
    _unsub_state["last_error"] = last_error
    _unsub_state["last_error_kind"] = last_kind
    if reports and any(r.ok for r in reports):
        _unsub_state["last_success"] = time.time()
    _unsub_journal("cycle", mailboxes=len(reports), examined=examined,
                   added=added, duplicates=duplicates, error=last_error[:120])


def _unsub_worker(sender_email: Optional[str] = None) -> None:
    """The whole cycle, off the rerun path. Holds _unsub_lock for its lifetime.

    No st.* call anywhere in here: a thread has no ScriptRunContext, so a
    st.write() would be dropped with a warning. Results land in _unsub_state
    and show up on the next render. That is also why the Notion write path no
    longer calls st.error().
    """
    try:
        budget_end = time.monotonic() + _UNSUB_CYCLE_BUDGET_S
        reports = []
        if not sender_email:
            _unsub_advance_rr()
        for cfg in _unsub_configs(sender_email):
            left = budget_end - time.monotonic()
            if left < 3.0:
                _unsub_journal("cycle_truncated", mailbox=cfg.user, role=cfg.role)
                break
            report = unsubscribe_inbox.scan_mailbox(
                unsubscribe_inbox.with_deadline(cfg, left))
            if report.ok:
                _apply_report(report)
            reports.append(report)
        _record_reports(reports)
    except (KeyboardInterrupt, SystemExit):
        raise
    except BaseException as exc:                      # noqa: BLE001
        _unsub_state["last_error"] = f"{exc.__class__.__name__}: {exc}"[:300]
        _unsub_state["last_error_kind"] = "unknown"
        _unsub_journal("worker_crash", error=str(exc)[:200])
    finally:
        _unsub_state["running"] = False
        try:
            _unsub_lock.release()
        except RuntimeError:
            pass


def _should_sync_now(state: dict, now: Optional[float] = None,
                     force: bool = False) -> Tuple[bool, str]:
    """(go, reason). Pure apart from the default clock, so it is testable.

    force skips the throttle but never the per-mailbox backoff: broken
    credentials must not inflict a 20 s timeout on every click of Send.
    """
    now = time.time() if now is None else now
    if not UNSUB_SYNC_ENABLED:
        return False, "relève désactivée (UNSUB_SYNC_ENABLED=0)"
    if not UNSUB_IMAP_USER and not UNSUB_SCAN_BOUNCES:
        return False, "aucune boîte configurée"
    if state.get("running"):
        return False, "déjà en cours"
    if not force:
        since = now - state.get("last_attempt", 0.0)
        if since < _UNSUB_THROTTLE_S:
            return False, f"trop tôt ({int(_UNSUB_THROTTLE_S - since)} s)"
    return True, ""


def _kick_unsub_sync(force: bool = False, wait_s: float = 0.0,
                     sender_email: Optional[str] = None) -> dict:
    """Start a cycle if allowed, and return immediately unless wait_s is given.

    Every trigger goes through here — the manual button, the end of main(), the
    auto-refreshing status fragment, the pre-send gate — so there is exactly one
    code path that can talk to IMAP, and it is throttled and locked.
    """
    now = time.time()
    go, reason = _should_sync_now(_unsub_state, now, force)
    if go:
        if not _unsub_has_work(now):
            reason = "toutes les boîtes sont en attente (backoff) ou sans mot de passe"
        elif _unsub_lock.acquire(blocking=False):
            # Armed BEFORE the work: a crash still arms the throttle, so a
            # failing cycle cannot be retried on every rerun.
            _unsub_state["last_attempt"] = now
            _unsub_state["running"] = True
            thread = threading.Thread(
                target=_unsub_worker, args=(sender_email,),
                name="unsub-sync", daemon=True,
            )
            _unsub_state["thread"] = thread
            thread.start()
        else:
            reason = "déjà en cours"

    thread = _unsub_state.get("thread")
    if wait_s and thread is not None and thread.is_alive():
        thread.join(timeout=wait_s)
    alive = bool(thread is not None and thread.is_alive())
    return {
        "ok": not _unsub_state["last_error_kind"],
        "error": _unsub_state["last_error"],
        "reason": reason,
        "running": alive,
        "added": _unsub_state["added_last"],
        "duplicates": _unsub_state["duplicates_last"],
    }


def _unsub_pending_review() -> List:
    """Every candidate awaiting a human, flattened across mailboxes."""
    out = []
    for items in _unsub_state["review"].values():
        out.extend(items)
    return out[:_UNSUB_REVIEW_MAXLEN]


def _unsub_configured() -> bool:
    if not UNSUB_SYNC_ENABLED:
        return False
    if UNSUB_IMAP_USER and _unsub_password_for(UNSUB_IMAP_USER):
        return True
    return bool(UNSUB_SCAN_BOUNCES and any(
        u.get("email") and u.get("password") for u in _load_users()))


def _humanize_age(seconds: Optional[float]) -> str:
    """Rounded age, or "jamais" when the caller passes None.

    A slightly NEGATIVE value is not "never", it is "just now": the clock is
    read before the fetch it measures, so the age comes out at about -0.3 s on
    a first render. Mapping that to "jamais" printed a failure right next to a
    list that had just loaded perfectly. Callers signal a real "never" by
    passing None, not by passing a huge number.
    """
    if seconds is None:
        return "jamais"
    seconds = int(max(0.0, seconds))
    if seconds < 60:
        return f"il y a {seconds} s"
    if seconds < 3600:
        return f"il y a {seconds // 60} min"
    if seconds < 86400:
        return f"il y a {seconds // 3600} h"
    return f"il y a {seconds // 86400} j"


def _force_suppression_reload() -> None:
    """Clear every cached failure so the next read really goes to Notion."""
    _suppression_state["last_error"] = 0.0
    _suppression_state["last_error_msg"] = ""
    _suppression_state["last_write_error"] = ""
    _suppression_state["last_fetch"] = 0.0
    _suppression_state["resolve_failed_at"] = 0.0
    _suppression_state["resolved_ds_id"] = None
    _suppression_state["schema"] = None


def _render_unsub_status() -> None:
    """The sidebar status block. Call it inside `with st.sidebar:`.

    Two lines, deliberately distinct:
      * the LIST is the truth used to filter recipients, and its failure blocks
        sending;
      * the SWEEP is an ingestion convenience, and its failure only warns.
    Merging them would leave the operator unable to tell whether a send is safe.

    The bare "15 adresse(s)" this replaces is precisely what let the June 2026
    outage pass unnoticed: an unreachable Notion showed exactly the same thing
    as a legitimately empty list.
    """
    now = time.time()
    suppression = _load_suppression()
    gate, gate_msg = _suppression_gate(_suppression_state, now)

    if gate == "block":
        # Outside any expander on purpose: the state that blocks sending must
        # never require a click to be seen.
        st.error(gate_msg)
        if st.button("🔄 Réessayer la liste Notion", key="unsub_retry_notion",
                     use_container_width=True):
            _force_suppression_reload()
            st.rerun()
    elif not SUPPRESSION_NOTION_DS_ID:
        st.warning(
            "⚠️ Liste de désinscription **désactivée** : `SUPPRESSION_NOTION_DS_ID` "
            "absent de `secrets.toml`. Aucun filtrage n'est appliqué."
        )
    else:
        age = _humanize_age(
            now - _suppression_state["last_fetch"]
            if _suppression_state["last_fetch"] else None
        )
        st.caption(
            f"🚫 Liste de désinscription : **{len(suppression)}** adresses · chargée {age}"
        )
        if gate == "warn" and gate_msg:
            st.caption(gate_msg)

    if _suppression_state["last_write_error"]:
        # The message Notion itself returned. Its absence is what made the
        # "Raison" bug invisible for months.
        st.error(
            "⚠️ Notion a refusé la dernière écriture : "
            + _suppression_state["last_write_error"]
        )

    if not _unsub_configured():
        st.caption(
            "📥 Relève automatique désactivée — renseignez `UNSUB_IMAP_USER` "
            "dans `secrets.toml` pour lire la boîte de désinscription."
        )
        return

    if _unsub_state["running"]:
        st.caption("📥 Relève en cours…")
    elif _unsub_state["last_success"]:
        st.caption(
            "📥 Relève : %s · %d message(s) examiné(s) · **%d** ajoutée(s), %d déjà connue(s)"
            % (_humanize_age(now - _unsub_state["last_success"]),
               _unsub_state["examined_last"], _unsub_state["added_last"],
               _unsub_state["duplicates_last"])
        )
    else:
        st.caption("📥 Relève : jamais effectuée dans cette session.")

    for mailbox, message in list(_unsub_state["errors"].items()):
        until = _unsub_state["backoff"].get(mailbox, 0.0)
        wait = max(0, int((until - now) // 60))
        st.caption(f"⏳ {mailbox} : {message[:120]} — nouvelle tentative dans {wait} min")

    # Kick a cycle from HERE too, not only at the end of main(). A fragment
    # auto-rerun re-runs the fragment alone -- main() is not re-executed -- so
    # without this line nothing would advance while the operator sits idle,
    # which is the whole point of run_every. Non-blocking (it starts a daemon
    # thread) and throttled, so the cost is one cycle per _UNSUB_THROTTLE_S
    # however often this renders.
    try:
        _kick_unsub_sync(force=False)
    except Exception:
        pass


def _render_unsub_actions() -> None:
    """Manual sync button, review queue and journal. Inside the expander."""
    if _unsub_configured():
        if st.button("🔄 Synchroniser maintenant", key="unsub_sync_now",
                     use_container_width=True):
            with st.spinner("Lecture de la boîte de désinscription…"):
                # ignore the throttle AND the backoff: an operator pressing the
                # button explicitly is asking us to try anyway.
                _unsub_state["backoff"] = {}
                info = _kick_unsub_sync(force=True, wait_s=_UNSUB_CYCLE_BUDGET_S + 5)
            if info["error"]:
                st.error(f"Relève en échec : {info['error']}")
            st.rerun()
        st.caption(
            "La relève ne tourne que quand l'application est ouverte : "
            "Streamlit met le conteneur en veille. Une synchronisation est "
            "**forcée juste avant chaque envoi**, ce qui est le moment qui compte."
        )

    pending = _unsub_pending_review()
    if pending:
        st.warning(f"⚠️ {len(pending)} message(s) à trancher à la main")
        for cand in pending:
            target = cand.review_hint or cand.email
            st.caption(f"**{target or '(aucune adresse)'}** — {cand.why}")
            st.caption(f"« {cand.subject[:90]} » de {cand.from_addr}")
            if target and "," not in target:
                # Keyed by ADDRESS, not by position. Streamlit matches a click
                # to a widget key on the NEXT rerun, so a positional key would
                # let a sweep finishing in between shift the list and attach
                # the click to a different address -- and suppression cannot be
                # undone.
                if st.button(f"➕ Désinscrire {target}",
                             key=f"unsub_review_{cand.mailbox}_{cand.uid}_{target}",
                             use_container_width=True):
                    if add_suppression(target, reason="revue_manuelle",
                                       nom=cand.display_name,
                                       provenance=unsubscribe_inbox.provenance_lines(cand)):
                        st.success(f"{target} ajoutée.")
                        st.rerun()
                    else:
                        st.error(_suppression_state["last_write_error"] or "échec Notion")

    if _unsub_state["journal"]:
        with st.expander("Journal de la relève (mémoire)"):
            # The container sleeps and this ring is lost with it, so offer the
            # export. The durable copy lives in the Streamlit Cloud logs.
            lines = list(_unsub_state["journal"])
            for entry in reversed(lines[-15:]):
                st.caption(json.dumps(entry, ensure_ascii=False)[:220])
            st.download_button(
                "⬇️ Exporter le journal (JSON)",
                data=json.dumps(lines, ensure_ascii=False, indent=1),
                file_name="journal_desinscriptions.json",
                mime="application/json",
                key="unsub_journal_dl",
            )


# Auto-refreshing status: re-runs only itself, so a sleeping operator still sees
# the sweep advance. Capability-detected so a streamlit downgrade degrades to a
# plain (non-refreshing) call instead of an AttributeError at import time.
if hasattr(st, "fragment"):
    _unsub_status_fragment = st.fragment(run_every=_UNSUB_THROTTLE_S)(
        _render_unsub_status)
else:                                            # pragma: no cover
    _unsub_status_fragment = _render_unsub_status


def _presend_notice(key: str = "send") -> bool:
    """Render the pre-send state and return whether sending may start.

    Rendered on EVERY rerun, above the send buttons, and NOT inside a button
    handler: a click is transient, so a retry button or an override checkbox
    placed inside a handler would vanish on the very rerun the checkbox
    triggers, and could never take effect. The buttons are disabled instead.
    """
    gate, msg = _suppression_gate(_suppression_state)
    if gate == "block":
        st.error(msg)
        cols = st.columns([1, 2])
        with cols[0]:
            if st.button("🔄 Réessayer Notion", key=f"presend_retry_{key}"):
                _force_suppression_reload()
                _load_suppression()
                st.rerun()
        with cols[1]:
            override = st.checkbox(
                "Je confirme envoyer SANS liste de désinscription vérifiée",
                key=f"presend_override_{key}",
            )
        if override:
            st.warning(
                "Envoi débloqué manuellement, sans liste vérifiée. "
                "Cette décision est journalisée."
            )
        return bool(override)
    if gate == "warn" and msg:
        st.warning(msg)
    return True


def _presend_sync(sender_email: Optional[str] = None) -> None:
    """Force a sweep as the first thing a send handler does.

    This is what makes the invariant true: no send starts without a sync
    attempt having just happened. It costs a couple of seconds inside a handler
    where the operator is already waiting (the send loop itself sleeps 1 to 10 s
    per email).

    It NEVER blocks the send. The asymmetry is deliberate: Notion is the truth
    used to filter, so a Notion outage stops the send (see _presend_notice);
    IMAP is only an ingestion convenience, so an IMAP outage warns. Blocking on
    IMAP would be worse than the disease, since one wrong Gmail app password
    would freeze every campaign.

    It works without touching the send loops because they re-test
    is_suppressed() per contact, and add_suppression() has already updated the
    in-memory cache: an address discovered 200 ms ago is honoured within the
    same click, and counted in the skipped_suppressed_count already displayed.
    """
    if _suppression_gate(_suppression_state)[0] == "block":
        # We only get here when the operator ticked the override box.
        _unsub_journal("gate_override", sender=sender_email or "",
                       why=str(_suppression_state.get("last_error_msg", ""))[:200])
    if not _unsub_configured():
        return
    with st.spinner("Vérification des désinscriptions…"):
        info = _kick_unsub_sync(force=True, wait_s=_UNSUB_PRESEND_WAIT_S,
                                sender_email=sender_email)
    if info.get("added"):
        st.info(
            f"🚫 {info['added']} désinscription(s) relevée(s) à l'instant — "
            "elles sont déjà exclues de cet envoi."
        )
    if info.get("error") and not info.get("ok"):
        st.warning(
            f"⚠️ Relève des désinscriptions indisponible ({str(info['error'])[:160]}). "
            "L'envoi continue sur la base de la liste Notion, qui reste la référence."
        )
    elif info.get("running"):
        st.warning(
            f"⚠️ La relève n'a pas fini en {int(_UNSUB_PRESEND_WAIT_S)} s. "
            "L'envoi continue sur la liste Notion connue."
        )


def parse_email_list(raw: str) -> List[str]:
    """Extract individual email addresses from a free-form text blob
    (newlines / commas / semicolons / spaces). Lowercased + deduped."""
    if not raw:
        return []
    pattern = r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"
    found = re.findall(pattern, raw)
    seen = []
    seen_set = set()
    for e in found:
        low = e.strip().lower()
        if low and low not in seen_set:
            seen_set.add(low)
            seen.append(low)
    return seen


# --- Unsubscribe footer guard (F2) -----------------------------------------
# The template already injects the footer. This guard makes sure manually
# edited emails (text_area in the "invalid" review flow) also carry it.

def _ensure_unsubscribe_footer(html: str) -> str:
    """Append the unsubscribe footer if the marker is absent."""
    if not html:
        return UNSUBSCRIBE_FOOTER_HTML
    if UNSUBSCRIBE_FOOTER_MARKER in html:
        return html
    return html + "\n" + UNSUBSCRIBE_FOOTER_HTML


# --- Unified email send (refactor §6) --------------------------------------
# Single source of truth used by both send loops. Adds List-Unsubscribe /
# List-Unsubscribe-Post / Reply-To headers (F1), guarantees the footer (F2),
# and handles inline images + attachments.

def _filter_cc(cc_emails: str) -> List[str]:
    """CC addresses minus anyone on the suppression list.

    The CC path bypassed is_suppressed() entirely: send_one_email() did
    recipients.extend(...) straight from the raw string, so a suppressed
    address in copy was mailed anyway. A CC is usually a colleague, but if one
    is on the list the send is a violation. Used by BOTH the Cc header and the
    SMTP envelope, so the two can never disagree.
    """
    if not cc_emails or not cc_emails.strip():
        return []
    kept = []
    for raw in cc_emails.split(','):
        addr = raw.strip()
        if addr and not is_suppressed(addr):
            kept.append(addr)
    return kept


def _build_email_message(
    sender_email: str,
    email_data: dict,
    cc_emails: str,
    decorative_image_file,
    attachment_files,
    automation,
    decorative_image_file_2=None,
):
    """Build the MIMEMultipart('mixed') root with all headers and parts."""
    msg_root = MIMEMultipart('mixed')
    msg_root['From'] = sender_email
    msg_root['To'] = email_data['email']
    msg_root['Subject'] = email_data.get(
        'personalized_subject', 'MERCI RAYMOND - Votre service paysagiste'
    )

    _cc_kept = _filter_cc(cc_emails)
    if _cc_kept:
        # Rebuilt from the filtered list so the header cannot advertise a
        # recipient the envelope no longer carries.
        msg_root['Cc'] = ', '.join(_cc_kept)

    # F1 — Deliverability headers (RFC 8058 + RFC 2369).
    # mailto: only for now; an HTTPS one-click endpoint can be added later.
    msg_root['List-Unsubscribe'] = (
        f'<mailto:{UNSUBSCRIBE_MAILTO}?subject=unsubscribe%20{email_data["email"]}>'
    )
    msg_root['List-Unsubscribe-Post'] = 'List-Unsubscribe=One-Click'
    msg_root['Reply-To'] = sender_email

    # F2 — Make sure the footer is always there, even on hand-edited HTML.
    html_body = _ensure_unsubscribe_footer(email_data['personalized_email'])

    # F4 — Put the recipient's address in the VISIBLE footer link too.
    # UNSUBSCRIBE_FOOTER_HTML is a module constant, so it is built without any
    # address, while the List-Unsubscribe header above carries one. Result: a
    # click on Gmail's Unsubscribe button (which uses the header) tells us who
    # to unsubscribe, and a click on the visible link does not -- we fall back
    # to the sender of the incoming mail. If the email had been forwarded, that
    # is the wrong person. The closing quote anchors the replacement on the
    # whole href, and UNSUBSCRIBE_FOOTER_MARKER survives it, so
    # _ensure_unsubscribe_footer() stays idempotent.
    # Note this uses email_data['email'], which in test mode is the SENDER's
    # address, not the prospect's: a click on a test email must never be able
    # to unsubscribe a real prospect.
    html_body = html_body.replace(
        f'mailto:{UNSUBSCRIBE_MAILTO}?subject=unsubscribe"',
        f'mailto:{UNSUBSCRIBE_MAILTO}?subject=unsubscribe%20{email_data["email"]}"',
    )

    alt = MIMEMultipart('alternative')
    msg_root.attach(alt)

    plain_text = html2text.html2text(html_body)
    alt.attach(MIMEText(plain_text, 'plain', 'utf-8'))

    rel = MIMEMultipart('related')
    rel.attach(MIMEText(html_body, 'html', 'utf-8'))
    alt.attach(rel)

    # Inline decorative images. Each gets a distinct Content-ID matching the
    # cid: references emitted by personalize_email (decorative_image / decorative_image_2).
    for img_file, cid, fname in (
        (decorative_image_file,   'decorative_image',   'decorative_image.jpg'),
        (decorative_image_file_2, 'decorative_image_2', 'decorative_image_2.jpg'),
    ):
        if not img_file:
            continue
        try:
            compressed = automation.compress_image(img_file)
            mime_img = MIMEImage(compressed.getvalue())
            mime_img.add_header('Content-ID', f'<{cid}>')
            mime_img.add_header('Content-Disposition', 'inline', filename=fname)
            rel.attach(mime_img)
        except Exception as e:
            st.warning(f"⚠️ Impossible d'ajouter l'image décorative ({cid}): {e}")

    if attachment_files:
        for attachment_file in attachment_files:
            try:
                if attachment_file.type.startswith('image/'):
                    attachment = MIMEImage(attachment_file.getvalue())
                else:
                    attachment = MIMEApplication(attachment_file.getvalue())
                attachment.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=attachment_file.name,
                )
                msg_root.attach(attachment)
            except Exception as e:
                st.warning(f"⚠️ Impossible de joindre {attachment_file.name}: {e}")

    return msg_root


def send_one_email(
    server,
    sender_email: str,
    email_data: dict,
    cc_emails: str,
    automation,
    decorative_image_file=None,
    attachment_files=None,
    decorative_image_file_2=None,
) -> None:
    """Send one email over an open SMTP server. Raises on failure so the
    caller can count failures and keep going. Does NOT enforce sleep,
    suppression, or daily cap — those are the caller's responsibility."""
    msg_root = _build_email_message(
        sender_email, email_data, cc_emails,
        decorative_image_file, attachment_files, automation,
        decorative_image_file_2=decorative_image_file_2,
    )
    text = msg_root.as_string()
    recipients = [email_data['email']] + _filter_cc(cc_emails)
    server.sendmail(sender_email, recipients, text)


def _append_user_to_file(name: str, email: str, password: str) -> Optional[str]:
    """Append one user to users.json. Returns None on success, error message on failure."""
    path = _users_file_path()
    extra = []
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                extra = json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
        if not isinstance(extra, list):
            extra = []
    extra.append({"name": name.strip(), "email": email.strip(), "password": password.strip()})
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(extra, f, ensure_ascii=False, indent=2)
    except OSError as e:
        return str(e)
    return None


# --- Saved uploads (images/files) for reuse ---
UPLOADED_ASSETS_DIR = Path(__file__).resolve().parent / "uploaded_assets"
DECORATIVE_SUBDIR = "decorative"
ATTACHMENTS_SUBDIR = "attachments"


def _uploaded_assets_path(subdir: str) -> Path:
    path = UPLOADED_ASSETS_DIR / subdir
    path.mkdir(parents=True, exist_ok=True)
    return path


def _save_uploaded_file(upload, subdir: str) -> Optional[Path]:
    """Save an uploaded file (Streamlit UploadedFile) to uploaded_assets/subdir. Returns path or None on error."""
    if upload is None:
        return None
    try:
        folder = _uploaded_assets_path(subdir)
        name = (upload.name or "file").strip() or "file"
        # Sanitize and avoid overwrite: base_timestamp.ext
        base, ext = os.path.splitext(name)
        base = re.sub(r"[^\w\-.]", "_", base)[:80] or "file"
        ext = (ext or ".bin").lower()
        path = folder / f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
        path.write_bytes(upload.getvalue())
        return path
    except Exception:
        return None


def _list_saved_files(subdir: str) -> List[Tuple[str, Path]]:
    """Return list of (display_name, path) for files in uploaded_assets/subdir, newest first."""
    folder = UPLOADED_ASSETS_DIR / subdir
    if not folder.exists():
        return []
    out = []
    for p in folder.iterdir():
        if p.is_file():
            out.append((p.name, p))
    out.sort(key=lambda x: x[1].stat().st_mtime, reverse=True)
    return out


class _FileLikeFromPath:
    """Wrap a file path as a file-like object with .getvalue(), .name, .type, .read(), .seek() for downstream code."""

    def __init__(self, path: Path):
        self._path = Path(path)
        self.name = self._path.name
        ext = self._path.suffix.lower()
        self.type = {
            "jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "gif": "image/gif",
            "pdf": "application/pdf", "doc": "application/msword", "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "xls": "application/vnd.ms-excel", "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "txt": "text/plain",
        }.get(ext, "application/octet-stream")
        self._bytes = None

    def getvalue(self):
        if self._bytes is None:
            self._bytes = self._path.read_bytes()
        return self._bytes

    def read(self):
        return self.getvalue()

    def seek(self, pos):
        pass


def _load_saved_file(path: Path):
    """Return a file-like object for the given path (for use as decorative_image_file or attachment)."""
    return _FileLikeFromPath(path)


def main():
    st.markdown('<h1 class="main-header">🌱 MERCI RAYMOND - Raymographe</h1>', unsafe_allow_html=True)

    # Initialize session state
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'email_automation' not in st.session_state:
        st.session_state.email_automation = EmailAutomation()
    if 'processed_emails' not in st.session_state:
        st.session_state.processed_emails = []
    if 'edited_invalid_emails' not in st.session_state:
        st.session_state.edited_invalid_emails = {}
    if 'validated_invalid_emails' not in st.session_state:
        st.session_state.validated_invalid_emails = []
    if 'notion_sites' not in st.session_state:
        st.session_state.notion_sites = None

    # Sidebar for user selection
    st.sidebar.header("👤 Choisissez un Utilisateur")

    if st.session_state.pop("user_just_added_notice", False):
        st.sidebar.success("✅ Utilisateur ajouté avec succès — vous êtes maintenant sur ce profil.")
        try:
            st.toast("Utilisateur ajouté !", icon="✅")
        except Exception:
            pass

    USERS = _load_users()

    # Create user selection dropdown
    user_options = [user["name"] for user in USERS]
    if user_options:
        if (
            "utilisateur_select" in st.session_state
            and st.session_state.utilisateur_select not in user_options
        ):
            st.session_state.utilisateur_select = user_options[0]

        selected_user_name = st.sidebar.selectbox(
            "Utilisateur",
            options=user_options,
            help="Sélectionnez l'utilisateur pour l'envoi des emails",
            key="utilisateur_select",
        )

        # Find selected user and store credentials in session state
        selected_user = next((user for user in USERS if user["name"] == selected_user_name), None)
        if selected_user:
            sender_email = selected_user["email"]
            sender_password = selected_user["password"]
            st.session_state.sender_email = sender_email
            st.session_state.sender_password = sender_password
        else:
            sender_email = None
            sender_password = None
    else:
        st.sidebar.warning("Aucun utilisateur configuré — ajoutez un bloc `[users.X]` à `.streamlit/secrets.toml` ou utilisez le formulaire ci-dessous.")
        sender_email = None
        sender_password = None

    # --- Deliverability sidebar widget (F3 + F4) --------------------------
    # Suppression list — source of truth is the Notion DB plugged via
    # SUPPRESSION_NOTION_DS_ID. The team can append addresses by hand, and the
    # F4 sweep appends them automatically from the desinscription@ mailbox.
    # No remove path: the team rule is that suppressed = forever (Notion is
    # editable directly if rare manual corrections are needed).
    st.sidebar.divider()
    st.sidebar.subheader("📬 Délivrabilité")
    with st.sidebar:
        # A fragment re-runs ITSELF every run_every seconds without a full
        # script rerun, so the status advances while the operator sits idle,
        # and it is a normal script context (st.* is safe inside).
        # Capability-detected: pinned streamlit is 1.58, st.fragment landed in
        # 1.37, but a downgrade must degrade to a plain call, not crash.
        try:
            _unsub_status_fragment()
        except Exception:
            _render_unsub_status()
    with st.sidebar.expander("🚫 Gestion des désinscriptions"):
        _render_unsub_actions()

        # Surface the success banner from the previous run (set by the form
        # handler below, then we st.rerun() so the counter refreshes from Notion).
        # pop() shows it once and clears the flag, so it doesn't stick forever.
        _added_count = st.session_state.pop('_suppression_add_success', None)
        if _added_count:
            st.success(f"✅ {_added_count} adresse(s) ajoutée(s) à Notion.")
        _dup_count = st.session_state.pop('_suppression_add_duplicates', None)
        if _dup_count:
            st.info(f"{_dup_count} adresse(s) étaient déjà dans la liste.")

        # Use a form so clear_on_submit handles the text_area reset for us —
        # writing to st.session_state['suppression_add_input'] manually fails
        # because the widget has already been instantiated in the same run.
        with st.form("suppression_add_form", clear_on_submit=True):
            new_unsubs = st.text_area(
                "Emails à ajouter (un par ligne ou séparés par , ; espace)",
                key="suppression_add_input",
                height=100,
                placeholder="exemple@client.fr\nautre@client.fr",
            )
            submitted = st.form_submit_button("➕ Ajouter")

        if submitted:
            parsed = parse_email_list(new_unsubs)
            # Cap the batch: this loop used to call add_suppression() per
            # address with no shared resolution and no breaker, so pasting a
            # long list against a hanging Notion froze the rerun for minutes.
            _MANUAL_ADD_CAP = 50
            if len(parsed) > _MANUAL_ADD_CAP:
                st.warning(
                    f"{len(parsed)} adresses collées : seules les "
                    f"{_MANUAL_ADD_CAP} premières sont traitées. Recommencez "
                    "pour les suivantes."
                )
                parsed = parsed[:_MANUAL_ADD_CAP]
            known = _load_suppression()          # one read, not one per address
            _resolve_suppression_ds_id()         # one resolution, not one per address
            added = 0
            duplicates = 0
            failed = []
            for e in parsed:
                if e in known:
                    duplicates += 1
                elif add_suppression(e, reason="manual_unsubscribe"):
                    added += 1
                else:
                    failed.append(e)
            if failed:
                # One aggregated error carrying Notion's own words, instead of
                # one st.error per address saying only "400 Bad Request".
                st.error(
                    "%d adresse(s) refusée(s) par Notion (%s%s). Détail : %s"
                    % (len(failed), ", ".join(failed[:3]),
                       "…" if len(failed) > 3 else "",
                       _suppression_state["last_write_error"] or "cause inconnue")
                )
            if added or duplicates:
                # Stash the counts and rerun; the banner block above re-renders
                # the message on the next pass with a fresh counter.
                st.session_state['_suppression_add_success'] = added
                st.session_state['_suppression_add_duplicates'] = duplicates
                st.rerun()
            elif not failed:
                st.warning("Aucune adresse email valide détectée.")
    # ----------------------------------------------------------------------

    # Add user from UI
    with st.sidebar.expander("➕ Ajouter un utilisateur"):
        add_name = st.text_input("Nom", key="add_user_name", placeholder="Jean Dupont")
        add_email = st.text_input("Email", key="add_user_email", placeholder="jean@exemple.fr")
        add_password = st.text_input("Mot de passe (app password)", key="add_user_password", type="password", placeholder="xxxx xxxx xxxx xxxx")
        if st.button("Ajouter"):
            if not (add_name and add_email):
                st.error("Nom et email sont obligatoires.")
            else:
                err = _append_user_to_file(add_name, add_email, add_password or "")
                if err:
                    st.error(f"Erreur: {err}")
                else:
                    new_name = add_name.strip()
                    st.session_state.utilisateur_select = new_name
                    st.session_state.user_just_added_notice = True
                    for k in ("add_user_name", "add_user_email", "add_user_password"):
                        if k in st.session_state:
                            st.session_state[k] = ""
                    try:
                        st.balloons()
                    except Exception:
                        pass
                    st.rerun()

    # Main content
    tab1, tab2, tab3 = st.tabs(["📁 Upload & Preview", "🎨 Design Email", "🚀 Envoi"])

    # --- Helper: process a DataFrame (column detection, stats, store contacts) ---
    def _display_dataframe_results(df: pd.DataFrame):
        """Process a loaded DataFrame: detect columns, extract emails, show stats."""
        st.session_state.df = df

        st.success(f"Donnees chargees avec succes! {len(df)} lignes trouvees.")

        # Show preview
        st.subheader("Apercu des donnees")
        st.dataframe(df.head(10))

        # Detect column mapping and show it
        mapping = st.session_state.email_automation.detect_column_mapping(df)
        email_column = mapping['email_column']
        available_placeholders = mapping['available_placeholders']
        full_name_columns = mapping['full_name_columns']
        all_columns = mapping['all_columns']

        st.subheader("Detection automatique des colonnes")
        col1, col2 = st.columns(2)

        with col1:
            st.write("**Colonnes detectees:**")

            # Show detected email column
            if email_column:
                st.write(f"- **Email detecte:** `{email_column}`")
            else:
                st.write("- **Email:** Non detecte")
                st.warning("Aucune colonne email detectee. Veuillez verifier votre fichier.")

            # Show all available placeholders
            if available_placeholders:
                st.write("**Placeholders disponibles:**")

                # Group placeholders by type
                regular_placeholders = []
                name_placeholders = {}

                for col_name in available_placeholders.keys():
                    if col_name.endswith('_first') or col_name.endswith('_last'):
                        base_name = col_name.replace('_first', '').replace('_last', '')
                        if base_name not in name_placeholders:
                            name_placeholders[base_name] = {'first': None, 'last': None, 'full': None}

                        if col_name.endswith('_first'):
                            name_placeholders[base_name]['first'] = col_name
                        elif col_name.endswith('_last'):
                            name_placeholders[base_name]['last'] = col_name
                    else:
                        regular_placeholders.append(col_name)

                # Show regular placeholders
                for col_name in regular_placeholders:
                    st.write(f"- `{{{col_name}}}`")

                # Show name placeholders in groups
                for base_name, placeholders in name_placeholders.items():
                    if placeholders['first'] and placeholders['last']:
                        st.write(f"- **{base_name}:** `{{{base_name}}}` (nom complet), `{{{placeholders['first']}}}` (prenom), `{{{placeholders['last']}}}` (nom de famille)")
            else:
                st.write("**Placeholders:** Aucun (seulement email)")

        # Get valid emails using new system
        valid_contacts = st.session_state.email_automation.get_valid_emails_from_df(df)

        with col2:
            st.write("**Statistiques:**")
            st.metric("Total lignes", len(df))
            st.metric("Emails valides", len(valid_contacts))
            if len(df) > 0:
                st.metric("Taux email", f"{len(valid_contacts)/len(df)*100:.1f}%")

            # Show duplicate removal info
            duplicates_removed = st.session_state.get('duplicates_removed', 0)
            if duplicates_removed > 0:
                st.metric("Doublons retires", duplicates_removed)
                st.info(f"{len(valid_contacts)} emails uniques (dont {duplicates_removed} doublons retires)")
            else:
                st.info(f"{len(valid_contacts)} emails uniques")

            # Show suppression filter info
            suppressed_removed = st.session_state.get('suppressed_removed', 0)
            if suppressed_removed > 0:
                st.metric("Désinscrits ignorés", suppressed_removed)
                st.info(f"🚫 {suppressed_removed} adresse(s) retirée(s) via la liste de désinscription")

        # Show user guidance
        if email_column and available_placeholders:
            placeholder_list = ", ".join([f"`{{{col}}}`" for col in available_placeholders.keys()])
        elif email_column:
            st.info("**Email detecte!** Vous pouvez personnaliser vos emails avec les donnees de cette colonne.")
        else:
            st.error("**Aucune colonne email detectee.** Verifiez que votre fichier contient une colonne avec des adresses email.")

        # Store valid contacts for later use
        st.session_state.valid_contacts = valid_contacts

    # --- Tab 1: Upload & Preview ---
    with tab1:
        st.markdown('<h2 class="step-header">Etape 1: Source des donnees</h2>', unsafe_allow_html=True)

        data_source = st.radio(
            "Choisissez la source de donnees :",
            options=["Fichier Excel / CSV", "Base Entretien (Notion)"],
            horizontal=True,
            key="data_source_choice"
        )

        st.divider()

        if data_source == "Fichier Excel / CSV":
            uploaded_file = st.file_uploader(
                "Choisissez votre fichier Excel ou CSV",
                type=["xlsx", "xls", "csv"],
                help="Excel ou CSV - l'app detectera automatiquement les noms, emails, entreprises, etc."
            )

            if uploaded_file is not None:
                try:
                    name_lower = (uploaded_file.name or "").lower()
                    if name_lower.endswith(".csv"):
                        raw_bytes = uploaded_file.getvalue()
                        buf = BytesIO(raw_bytes)
                        try:
                            df = pd.read_csv(buf, encoding="utf-8", sep=None, engine="python")
                        except UnicodeDecodeError:
                            buf.seek(0)
                            df = pd.read_csv(buf, encoding="latin-1", sep=None, engine="python")
                    else:
                        df = pd.read_excel(uploaded_file)

                    _display_dataframe_results(df)

                except Exception as e:
                    st.error(f"Erreur lors du chargement: {e}")

        else:  # Base Entretien (Notion)
            st.markdown("**Charger les sites depuis la base Notion Entretien**")

            if st.button("Charger les sites", key="load_notion_sites"):
                with st.spinner("Chargement des sites depuis Notion..."):
                    try:
                        sites = fetch_notion_sites()
                        if sites:
                            st.session_state.notion_sites = sites
                        else:
                            st.warning("Aucun site trouve dans la base Notion.")
                            st.session_state.notion_sites = None
                    except Exception as e:
                        st.error(f"Erreur lors du chargement depuis Notion: {e}")
                        st.session_state.notion_sites = None

            # Display sites if loaded
            if st.session_state.get('notion_sites'):
                sites = st.session_state.notion_sites

                # --- Classify sites as INT, EXT, or INT/EXT ---
                def _classify_site(name: str) -> str:
                    lower = name.lower()
                    if '(int/ext)' in lower or '(ext/int)' in lower:
                        return 'INT/EXT'
                    elif '(int)' in lower:
                        return 'INT'
                    elif '(ext)' in lower:
                        return 'EXT'
                    return 'AUTRE'

                site_types = [_classify_site(s["site"]) for s in sites]

                # --- Filter radio ---
                site_filter = st.radio(
                    "Filtrer par type :",
                    options=["Tous", "Exterieur (EXT)", "Interieur (INT)"],
                    horizontal=True,
                    key="notion_site_filter",
                )

                # Determine which sites to show and their default selection
                has_mixed = any(t == 'INT/EXT' for t in site_types)

                if site_filter == "Tous":
                    # Show all, all selected
                    display_indices = list(range(len(sites)))
                    default_selected = [True] * len(sites)
                elif site_filter == "Exterieur (EXT)":
                    # Show EXT + INT/EXT (INT/EXT unchecked at top)
                    mixed_idx = [i for i, t in enumerate(site_types) if t == 'INT/EXT']
                    ext_idx = [i for i, t in enumerate(site_types) if t == 'EXT']
                    display_indices = mixed_idx + ext_idx
                    default_selected = [False] * len(mixed_idx) + [True] * len(ext_idx)
                else:  # Interieur (INT)
                    # Show INT + INT/EXT (INT/EXT unchecked at top)
                    mixed_idx = [i for i, t in enumerate(site_types) if t == 'INT/EXT']
                    int_idx = [i for i, t in enumerate(site_types) if t == 'INT']
                    display_indices = mixed_idx + int_idx
                    default_selected = [False] * len(mixed_idx) + [True] * len(int_idx)

                # Show warning about INT/EXT sites if filtering
                if site_filter != "Tous" and has_mixed:
                    st.warning("Certains sites sont a la fois INT et EXT. Ils apparaissent en haut de la liste, non-selectionnes. Cochez ceux qui sont pertinents pour cet envoi.")

                # Build editable DataFrame for displayed sites
                sites_df = pd.DataFrame({
                    "Selectionne": default_selected,
                    "Site": [sites[i]["site"] for i in display_indices],
                    "Email": [sites[i]["email"] for i in display_indices],
                })

                st.markdown(f"**{len(display_indices)} sites affiches.** Deselectionnez ceux que vous ne souhaitez pas inclure :")

                edited_sites = st.data_editor(
                    sites_df,
                    column_config={
                        "Selectionne": st.column_config.CheckboxColumn("Selectionne", default=True),
                        "Site": st.column_config.TextColumn("Site", disabled=True),
                        "Email": st.column_config.TextColumn("Email", disabled=True),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key=f"notion_sites_editor_{site_filter}",
                )

                # Count selected
                selected = edited_sites[edited_sites["Selectionne"] == True]
                selected_with_email = selected[selected["Email"].str.strip().astype(bool)]
                st.info(f"{len(selected)} sites selectionnes dont {len(selected_with_email)} avec un email valide")

                if st.button("Utiliser la selection", type="primary", key="use_notion_selection"):
                    if len(selected) == 0:
                        st.warning("Veuillez selectionner au moins un site.")
                    else:
                        # Create a DataFrame compatible with the existing pipeline
                        result_df = selected[["Site", "Email"]].copy()
                        result_df = result_df.rename(columns={"Email": "email"})
                        _display_dataframe_results(result_df)

    with tab2:
        st.markdown('<h2 class="step-header">Étape 2: Contenu et Design de l\'email</h2>', unsafe_allow_html=True)




        # Set format to HTML
        email_format = "HTML (Gmail-style)"
        st.session_state.email_format = email_format




        # Use Gmail-style HTML template
        is_html_format = True
        base_template = st.session_state.email_automation.base_email_content_html


        # Email subject line
        email_subject = st.text_input(
            "Objet de l'email:",
            value="",
            help="L'objet de l'email qui apparaîtra dans la boîte de réception. Vous pouvez utiliser des placeholders comme {contact_name}, {Company}, {Site}, etc.",
            key=f"email_subject_{email_format}",  # Unique key per format
            placeholder=""
        )

        # Store email subject in session state
        st.session_state.email_subject = email_subject

        st.divider()

        # Allow user to modify email content (unified: header + content + footer)
        email_content = st.text_area(
            "Modifiez le contenu complet de votre email (en-tête, contenu principal, signature):",
            value="",
            height=400,
            help="Créez votre email complet ici. Utilisez des placeholders comme {contact_name}, {site}, etc. Formatage: **gras**, *italique*, - puces, {color:red}couleur{/color}, [texte du lien](https://url).",
            key=f"email_content_{email_format}",  # Unique key per format
            placeholder=""
        )

        # Show formatting help
        with st.expander("💡 Aide au formatage du texte"):
            # Show dynamic placeholders if Excel file is uploaded
            if st.session_state.df is not None:
                mapping = st.session_state.email_automation.detect_column_mapping(st.session_state.df)
                available_placeholders = mapping['available_placeholders']
                full_name_columns = mapping.get('full_name_columns', [])

                if available_placeholders:
                    # Group placeholders by type
                    regular_placeholders = []
                    name_placeholders = {}

                    for col_name in available_placeholders.keys():
                        if col_name.endswith('_first') or col_name.endswith('_last'):
                            base_name = col_name.replace('_first', '').replace('_last', '')
                            if base_name not in name_placeholders:
                                name_placeholders[base_name] = {'first': None, 'last': None, 'full': None}

                            if col_name.endswith('_first'):
                                name_placeholders[base_name]['first'] = col_name
                            elif col_name.endswith('_last'):
                                name_placeholders[base_name]['last'] = col_name
                        else:
                            regular_placeholders.append(col_name)

                    # Build placeholder display
                    placeholder_text = "**Placeholders disponibles depuis votre Excel :**\n\n"

                    # Regular placeholders
                    if regular_placeholders:
                        placeholder_text += "**Colonnes normales :**\n"
                        for col in regular_placeholders:
                            placeholder_text += f"- `{{{col}}}`\n"
                        placeholder_text += "\n"

                    # Name placeholders
                    if name_placeholders:
                        placeholder_text += "**Colonnes de noms (avec options prénom/nom) :**\n"
                        for base_name, placeholders in name_placeholders.items():
                            if placeholders['first'] and placeholders['last']:
                                placeholder_text += f"- **{base_name}:** `{{{base_name}}}` (nom complet), `{{{placeholders['first']}}}` (prénom), `{{{placeholders['last']}}}` (nom de famille)\n"

                    placeholder_text += "\n**Placeholders spéciaux :**\n"
                    placeholder_text += "- `{Image}` : place la 1re image décorative ici (sinon placement automatique)\n"
                    placeholder_text += "- `{Image2}` : place la 2e image décorative ici (sinon placement automatique)"

                    st.markdown(placeholder_text)
                else:
                    st.markdown("""
                    **Placeholders spéciaux :**
                    - `{Image}` : place la 1re image décorative ici (sinon placement automatique)
                    - `{Image2}` : place la 2e image décorative ici (sinon placement automatique)
                    """)
            else:
                st.markdown("""
                *Chargez un fichier Excel pour voir les placeholders disponibles*
                """)

            # Formatting explanation and examples at the end
            st.markdown('''<div style="margin-top: 1rem;">
<p><strong>Formatage disponible :</strong></p>
<ul>
<li><code>**texte**</code> &rarr; <strong>texte en gras</strong></li>
<li><code>*texte*</code> &rarr; <em>texte en italique</em></li>
<li><code>- element</code> &rarr; liste a puces (une ligne par element)</li>
<li><code>&nbsp;&nbsp;&nbsp;&nbsp;- sous-element</code> &rarr; sous-liste (4 espaces avant le tiret)</li>
<li><code>{color:nom}texte{/color}</code> &rarr; texte en couleur</li>
<li><code>[texte](https://url)</code> &rarr; <a href="#" style="color:#1a73e8; text-decoration:underline;">lien cliquable</a> (uniquement http:// ou https://)</li>
</ul>
</div>''', unsafe_allow_html=True)
            st.markdown('''<div style="margin-top: 0.5rem;">
<p><strong>Palette de couleurs :</strong></p>
<table style="border-collapse: collapse; font-size: 13px;">
<tr>
<td style="padding: 3px 12px;"><span style="color:#1a73e8;">&#9632;</span> <code>bleu</code></td>
<td style="padding: 3px 12px;"><span style="color:#c62828;">&#9632;</span> <code>rouge</code></td>
<td style="padding: 3px 12px;"><span style="color:#2e7d32;">&#9632;</span> <code>vert</code></td>
<td style="padding: 3px 12px;"><span style="color:#e65100;">&#9632;</span> <code>orange</code></td>
<td style="padding: 3px 12px;"><span style="color:#6a1b9a;">&#9632;</span> <code>violet</code></td>
</tr>
<tr>
<td style="padding: 3px 12px;"><span style="color:#00838f;">&#9632;</span> <code>sarcelle</code></td>
<td style="padding: 3px 12px;"><span style="color:#4e342e;">&#9632;</span> <code>marron</code></td>
<td style="padding: 3px 12px;"><span style="color:#37474f;">&#9632;</span> <code>ardoise</code></td>
<td style="padding: 3px 12px;"><span style="color:#ad1457;">&#9632;</span> <code>framboise</code></td>
<td style="padding: 3px 12px;"><span style="color:#1a237e;">&#9632;</span> <code>marine</code></td>
</tr>
</table>
<p style="margin-top: 6px; font-size: 13px;">Utilisation : <code>{color:bleu}votre texte{/color}</code> &rarr; <span style="color:#1a73e8;">votre texte</span></p>
</div>''', unsafe_allow_html=True)
            st.markdown('''<div style="color: #2196F3; margin-top: 0.5rem;">
<p><strong>Exemples :</strong></p>
<ul>
<li><code>**Bonjour** *{contact_name}*</code> &rarr; <strong>Bonjour</strong> <em>Marie</em></li>
<li><code>- Premier point</code> (puis nouvelle ligne) <code>- Deuxieme point</code> &rarr; liste a puces</li>
<li><code>{color:vert}texte vert{/color}</code> &rarr; <span style="color:#2e7d32;">texte vert</span></li>
</ul>
</div>''', unsafe_allow_html=True)

        # Store custom email content with format awareness
        if email_content and email_content.strip():
            st.session_state.custom_email_content = email_content
            st.session_state.custom_email_format = email_format
        else:
            st.session_state.custom_email_content = None
            st.session_state.custom_email_format = None

        st.divider()

        # Visual elements section
        if 'saved_decorative_choice' not in st.session_state:
            st.session_state.saved_decorative_choice = None

        decorative_image_file = st.file_uploader(
            "Image décorative",
            type=['png', 'jpg', 'jpeg'],
            help="Image qui apparaîtra dans le corps de l'email HTML. Elle sera enregistrée pour réutilisation ultérieure."
        )

        # Save new upload and persist; or allow picking a saved image
        if decorative_image_file:
            saved_path = _save_uploaded_file(decorative_image_file, DECORATIVE_SUBDIR)
            st.session_state.decorative_image_file = decorative_image_file
            st.session_state.saved_decorative_choice = None
            if saved_path:
                st.success("✅ Image décorative chargée et enregistrée pour plus tard.")
            else:
                st.success("✅ Image décorative chargée")
        else:
            saved_list = _list_saved_files(DECORATIVE_SUBDIR)
            options = ["— Aucune —"] + [name for name, _ in saved_list]
            idx = 0
            if st.session_state.saved_decorative_choice:
                for i, (name, _) in enumerate(saved_list):
                    if name == st.session_state.saved_decorative_choice:
                        idx = i + 1
                        break
            chosen = st.selectbox(
                "Ou utiliser une image enregistrée",
                options=options,
                index=idx,
                key="saved_decorative_select"
            )
            if chosen and chosen != "— Aucune —":
                for name, path in saved_list:
                    if name == chosen:
                        st.session_state.decorative_image_file = _load_saved_file(path)
                        st.session_state.saved_decorative_choice = name
                        break
            else:
                st.session_state.saved_decorative_choice = None
                st.session_state.decorative_image_file = None

        # Ensure session state key exists
        if 'decorative_image_file' not in st.session_state:
            st.session_state.decorative_image_file = None

        # Size of the main decorative image
        size_options = list(st.session_state.email_automation.decorative_image_sizes.keys())
        default_size_index = size_options.index("Pleine largeur (100%)") if "Pleine largeur (100%)" in size_options else 0
        if 'decorative_image_size' not in st.session_state:
            st.session_state.decorative_image_size = size_options[default_size_index]
        decorative_image_size_label = st.selectbox(
            "Taille de l'image décorative",
            options=size_options,
            index=size_options.index(st.session_state.decorative_image_size) if st.session_state.decorative_image_size in size_options else default_size_index,
            help="Choisissez la largeur maximale de l'image dans l'email. L'aperçu affiche un cadre de cette taille si aucune image n'est chargée.",
            key="decorative_image_size_select"
        )
        st.session_state.decorative_image_size = decorative_image_size_label
        decorative_image_size_css = st.session_state.email_automation.decorative_image_sizes.get(decorative_image_size_label, "100%")

        # --- Optional 2nd inline image (placeholder {Image2}) -------------
        if 'saved_decorative_choice_2' not in st.session_state:
            st.session_state.saved_decorative_choice_2 = None
        if 'decorative_image_file_2' not in st.session_state:
            st.session_state.decorative_image_file_2 = None

        with st.expander("➕ Ajouter une 2e image (optionnel)"):
            st.caption("Place le placeholder `{Image2}` dans ton texte pour la positionner précisément. Sinon elle s'affiche sous la première.")
            decorative_image_file_2 = st.file_uploader(
                "Image décorative n°2",
                type=['png', 'jpg', 'jpeg'],
                help="2e image inline. Enregistrée pour réutilisation comme la première.",
                key="decorative_image_file_2_uploader",
            )

            if decorative_image_file_2:
                saved_path_2 = _save_uploaded_file(decorative_image_file_2, DECORATIVE_SUBDIR)
                st.session_state.decorative_image_file_2 = decorative_image_file_2
                st.session_state.saved_decorative_choice_2 = None
                if saved_path_2:
                    st.success("✅ 2e image chargée et enregistrée pour plus tard.")
                else:
                    st.success("✅ 2e image chargée")
            else:
                saved_list_2 = _list_saved_files(DECORATIVE_SUBDIR)
                options_2 = ["— Aucune —"] + [name for name, _ in saved_list_2]
                idx_2 = 0
                if st.session_state.saved_decorative_choice_2:
                    for i, (name, _) in enumerate(saved_list_2):
                        if name == st.session_state.saved_decorative_choice_2:
                            idx_2 = i + 1
                            break
                chosen_2 = st.selectbox(
                    "Ou utiliser une image enregistrée",
                    options=options_2,
                    index=idx_2,
                    key="saved_decorative_select_2",
                )
                if chosen_2 and chosen_2 != "— Aucune —":
                    for name, path in saved_list_2:
                        if name == chosen_2:
                            st.session_state.decorative_image_file_2 = _load_saved_file(path)
                            st.session_state.saved_decorative_choice_2 = name
                            break
                else:
                    st.session_state.saved_decorative_choice_2 = None
                    st.session_state.decorative_image_file_2 = None

            if 'decorative_image_size_2' not in st.session_state:
                st.session_state.decorative_image_size_2 = size_options[default_size_index]
            decorative_image_size_label_2 = st.selectbox(
                "Taille de la 2e image",
                options=size_options,
                index=size_options.index(st.session_state.decorative_image_size_2) if st.session_state.decorative_image_size_2 in size_options else default_size_index,
                help="Taille indépendante de l'image n°1.",
                key="decorative_image_size_select_2",
            )
            st.session_state.decorative_image_size_2 = decorative_image_size_label_2
        decorative_image_size_css_2 = st.session_state.email_automation.decorative_image_sizes.get(
            st.session_state.get('decorative_image_size_2', size_options[default_size_index]), "100%"
        )

        attachment_files = st.file_uploader(
            "Fichiers à joindre",
            type=['pdf', 'doc', 'docx', 'xls', 'xlsx', 'png', 'jpg', 'jpeg', 'txt'],
            accept_multiple_files=True,
            help="Sélectionnez un ou plusieurs fichiers à joindre à l'email. Ils seront enregistrés pour réutilisation."
        )

        # Store files in session state; save new attachments to disk
        if 'attachment_files' not in st.session_state:
            st.session_state.attachment_files = []
        if 'saved_attachment_names' not in st.session_state:
            st.session_state.saved_attachment_names = []

        if attachment_files:
            for f in attachment_files:
                _save_uploaded_file(f, ATTACHMENTS_SUBDIR)
            st.session_state.attachment_files = attachment_files
            st.success(f"✅ {len(attachment_files)} fichier(s) joint(s) chargé(s) et enregistrés.")
            for i, file in enumerate(attachment_files):
                st.write(f"📎 {file.name} ({file.size} bytes)")

        # Option to add from saved attachments
        saved_att_list = _list_saved_files(ATTACHMENTS_SUBDIR)
        if saved_att_list:
            add_saved = st.selectbox(
                "Fichier enregistré à ajouter",
                options=["— Choisir —"] + [name for name, _ in saved_att_list],
                key="add_saved_attachment"
            )
            if st.button("Ajouter ce fichier aux pièces jointes", key="add_saved_att_btn"):
                if add_saved and add_saved != "— Choisir —":
                    for name, path in saved_att_list:
                        if name == add_saved:
                            if any(getattr(f, "name", None) == name for f in st.session_state.attachment_files):
                                st.info(f"« {name} » est déjà dans les pièces jointes.")
                            else:
                                loaded = _load_saved_file(path)
                                st.session_state.attachment_files = list(st.session_state.attachment_files) + [loaded]
                                st.success(f"✅ « {name} » ajouté aux pièces jointes.")
                            break
                else:
                    st.warning("Choisissez un fichier dans la liste ci-dessus.")

        # Preview section
        st.divider()
        st.subheader("🎨 Aperçu du design Gmail-style")

        # Show Gmail-style HTML preview
        st.write("**📧 Aperçu Gmail-style:**")

        # Use actual first contact data if available, otherwise use placeholder data
        if st.session_state.df is not None and hasattr(st.session_state, 'valid_contacts') and st.session_state.valid_contacts:
            sample_contact = st.session_state.valid_contacts[0].copy()
        else:
            # Fallback to placeholder data
            sample_contact = {
                'contact_name': 'Marie Dupont',
                'Site': 'Bureau Paris',
                'Company': 'Entreprise ABC'
            }

        sample_html, sample_subject = st.session_state.email_automation.personalize_email(
                sample_contact, email_content if email_content else "", use_html=True,
                logo_file=None,
                decorative_image_file=st.session_state.decorative_image_file,
                attachment_files=st.session_state.get('attachment_files', []),
                email_subject=email_subject if email_subject else "",
                decorative_image_size=decorative_image_size_css,
                show_image_placeholder=False,  # WYSIWYG: empty image slots take no space (match the sent email)
                decorative_image_file_2=st.session_state.get('decorative_image_file_2'),
                decorative_image_size_2=decorative_image_size_css_2,
            )
        # In preview iframe cid: doesn't work; inject image(s) as base64 so they display at correct size
        for _ss_key, _cid in (('decorative_image_file', 'decorative_image'),
                              ('decorative_image_file_2', 'decorative_image_2')):
            _preview_img = st.session_state.get(_ss_key)
            if _preview_img is None:
                continue
            try:
                _preview_img.seek(0)
                _b64 = base64.b64encode(_preview_img.read()).decode('utf-8')
                _mime = _preview_img.type if getattr(_preview_img, 'type', None) else 'image/jpeg'
                _data_url = f"data:{_mime};base64,{_b64}"
                sample_html = sample_html.replace(f'src="cid:{_cid}"', f'src="{_data_url}"')
            except Exception:
                pass
        st.markdown(f'<p style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.4; color: #202124; margin: 0 0 16px 0;"><strong>Objet:</strong> {sample_subject}</p>', unsafe_allow_html=True)
        st.iframe(sample_html, height=300)

        # Add "Valider" button for personalization - right after preview
        st.markdown("<br>", unsafe_allow_html=True)

        if st.session_state.df is not None and hasattr(st.session_state, 'valid_contacts'):
            valid_contacts = st.session_state.valid_contacts

            if valid_contacts:
                # Get email content and subject from current tab
                # Use custom content if available, otherwise use what's in the text area
                if 'custom_email_content' in st.session_state and st.session_state.custom_email_content:
                    email_content_for_processing = st.session_state.custom_email_content
                else:
                    email_content_for_processing = email_content if email_content else ""

                email_subject_for_processing = email_subject if email_subject else ""

                if st.button("Valider", type="primary"):
                    # Always use Gmail-style HTML
                    use_html_for_processing = True

                    decorative_image_file_for_processing = st.session_state.get('decorative_image_file', None)
                    decorative_image_file_2_for_processing = st.session_state.get('decorative_image_file_2', None)

                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    processed_emails = []

                    for i, contact_data in enumerate(valid_contacts):
                        # Get a display name for status (use first available field or email)
                        display_name = contact_data.get('contact_name', contact_data.get('Name', contact_data.get('email', 'Contact')))
                        status_text.text(f"Traitement: {display_name} ({i+1}/{len(valid_contacts)})")

                        # Always use simple personalization - reliable and bulletproof
                        _img_size_css = st.session_state.email_automation.decorative_image_sizes.get(
                            st.session_state.get('decorative_image_size', 'Pleine largeur (100%)'), "100%"
                        )
                        _img_size_css_2 = st.session_state.email_automation.decorative_image_sizes.get(
                            st.session_state.get('decorative_image_size_2', 'Pleine largeur (100%)'), "100%"
                        )
                        personalized, personalized_subject = st.session_state.email_automation.personalize_email(
                            contact_data, email_content_for_processing, use_html_for_processing,
                            logo_file=None, decorative_image_file=decorative_image_file_for_processing,
                            attachment_files=st.session_state.get('attachment_files', []),
                            email_subject=email_subject_for_processing,
                            decorative_image_size=_img_size_css,
                            show_image_placeholder=False,
                            decorative_image_file_2=decorative_image_file_2_for_processing,
                            decorative_image_size_2=_img_size_css_2,
                        )

                        is_valid, issues = st.session_state.email_automation.verify_email_content(personalized)

                        processed_emails.append({
                            **contact_data,
                            'personalized_email': personalized,
                            'personalized_subject': personalized_subject,
                            'is_valid': is_valid,
                            'issues': issues,
                            'use_html': use_html_for_processing
                        })

                        progress_bar.progress((i + 1) / len(valid_contacts))
                        time.sleep(0.1)  # Small delay to show progress

                    st.session_state.processed_emails = processed_emails
                    status_text.text("✅ Traitement terminé!")

                    # Show summary
                    valid_count = sum(1 for email in processed_emails if email['is_valid'])
                    st.success(f"🎉 {valid_count}/{len(processed_emails)} emails prêts à envoyer")

                    if valid_count < len(processed_emails):
                        st.warning(f"⚠️ {len(processed_emails) - valid_count} emails nécessitent une révision")
            else:
                st.warning("Aucun email valide trouvé dans le fichier.")
        else:
            st.info("Veuillez d'abord charger un fichier Excel ou CSV dans l'onglet 'Upload & Preview'.")

    with tab3:
        st.markdown('<h2 class="step-header">Étape 3: Envoi des emails</h2>', unsafe_allow_html=True)

        # CC email addresses
        st.subheader("📋 CC (optionnel)")
        cc_emails = st.text_input(
            "Adresses email en copie",
            placeholder="email1@example.com, email2@example.com",
            help="Séparez plusieurs adresses par des virgules. Ces emails recevront une copie de tous les emails envoyés."
        )

        # Test mode checkbox
        test_mode = st.checkbox(
            "Mode test",
            help="Envoyer 5 emails de test à votre propre adresse (pas aux clients)"
        )

        st.divider()

        # Progress file to track sent emails
        progress_file = Path("email_sending_progress.json")

        # Load previously sent emails if file exists
        sent_emails_set = set()
        if progress_file.exists():
            try:
                with open(progress_file, 'r', encoding='utf-8') as f:
                    progress_data = json.load(f)
                    if 'sent_emails' in progress_data:
                        sent_emails_set = set(progress_data['sent_emails'])
            except Exception as e:
                st.warning(f"⚠️ Impossible de charger le progrès précédent: {e}")

        # Get sender credentials from session state
        sender_email = st.session_state.get('sender_email', None)
        sender_password = st.session_state.get('sender_password', None)

        if st.session_state.processed_emails:
            processed_emails = st.session_state.processed_emails
            valid_emails = [email for email in processed_emails if email['is_valid']]
            invalid_emails = [email for email in processed_emails if not email['is_valid']]

            # Apply test mode filter
            if test_mode:
                # Take first 5 emails but replace recipient addresses with sender's email
                valid_emails = valid_emails[:5]
                for email_data in valid_emails:
                    if 'original_email' not in email_data:  # Only modify if not already modified
                        email_data['original_email'] = email_data['email']  # Keep original for reference
                    email_data['email'] = sender_email  # Replace with sender's email
                st.info(f"🧪 Mode test activé - 5 emails de test seront envoyés à {sender_email}")
                st.warning("⚠️ Les emails de test seront envoyés à VOTRE adresse, pas aux destinataires réels")
            else:
                # Restore original email addresses when test mode is disabled
                # First, restore emails in the current valid_emails list
                for email_data in valid_emails:
                    if 'original_email' in email_data:
                        email_data['email'] = email_data['original_email']  # Restore original email
                        del email_data['original_email']  # Clean up

                # Also restore emails in the full processed_emails list to ensure complete restoration
                for email_data in processed_emails:
                    if 'original_email' in email_data:
                        email_data['email'] = email_data['original_email']  # Restore original email
                        del email_data['original_email']  # Clean up

            # Filter out already sent emails
            remaining_emails = [email for email in valid_emails if email['email'] not in sent_emails_set]
            already_sent_count = len(valid_emails) - len(remaining_emails)

            # F3 — Drop suppressed addresses so the counters and "ready to send" UI
            # reflect what will actually be sent. Test mode preserves the real
            # recipient in 'original_email', so we check that.
            def _suppression_target(e):
                return e.get('original_email', e.get('email', ''))

            pre_suppress_count = len(remaining_emails)
            remaining_emails = [e for e in remaining_emails if not is_suppressed(_suppression_target(e))]
            suppressed_in_batch = pre_suppress_count - len(remaining_emails)
            if suppressed_in_batch > 0:
                st.info(f"🚫 {suppressed_in_batch} destinataire(s) ignoré(s) (présents dans la liste de désinscription)")

            # F4 — One gate for both send buttons below (they share this scope).
            # Rendered here rather than inside either handler so the retry
            # button and the override checkbox survive the reruns they cause.
            _send_allowed = _presend_notice()

            # Show status if some emails were already sent
            if already_sent_count > 0 and len(remaining_emails) > 0:
                st.info(f"📊 **Progrès:** {already_sent_count} emails déjà envoyés. {len(remaining_emails)} emails restants à envoyer.")
                if st.button("🗑️ Effacer le progrès et recommencer", type="secondary"):
                    if progress_file.exists():
                        progress_file.unlink()
                    st.rerun()

            # Get valid_contacts count (use remaining for display)
            valid_contacts_count = len(remaining_emails)

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Emails prêts", valid_contacts_count)
            with col2:
                st.metric("Emails avec problèmes", len(invalid_emails))
            with col3:
                if remaining_emails:
                    # Hardcoded 1 second delay
                    sending_time = calculate_sending_time(valid_contacts_count, 1)
                    st.metric("Temps d'envoi", sending_time)

            # Configuration display
            st.markdown(f"""
            **Configuration :**
            - ⏱️ Délai entre emails : 1 seconde
            - 🧪 Mode test : {'Activé (5 emails max)' if test_mode else 'Désactivé'}
            - 📧 Emails à envoyer : {valid_contacts_count}
            - ⏰ Temps total estimé : {calculate_sending_time(valid_contacts_count, 1)}
            """)

            # Add test mode information
            if test_mode:
                st.info(f"📧 Test: Les 5 emails seront envoyés à {sender_email}")

            if invalid_emails:
                st.subheader("⚠️ Emails nécessitant une révision")
                st.info(f"📝 {len(invalid_emails)} emails à corriger avant envoi")

                for idx, email_data in enumerate(invalid_emails):
                    email_key = f"{email_data['email']}_{idx}"

                    # Check if this email has been validated
                    is_validated = email_key in st.session_state.validated_invalid_emails

                    # Display toggle/expander with status
                    status_icon = "✅" if is_validated else "❌"
                    # Get display name and location with fallbacks
                    display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                    location = email_data.get('site', email_data.get('Site', email_data.get('Location', 'N/A')))
                    with st.expander(f"{status_icon} {display_name} - {location}", expanded=not is_validated):
                        st.write("**Problèmes détectés:**")
                        for issue in email_data.get('issues', []):
                            st.warning(f"⚠️ {issue}")

                        st.divider()

                        # Get current content (edited or original)
                        if email_key in st.session_state.edited_invalid_emails:
                            current_content = st.session_state.edited_invalid_emails[email_key]
                        else:
                            current_content = email_data['personalized_email']

                        # Show editable text area for Gmail-style HTML
                        st.write("**Format:** Gmail-style HTML")
                        st.info("💡 Vous pouvez modifier le code HTML Gmail-style ci-dessous")
                        edited_content = st.text_area(
                            "Contenu de l'email:",
                            value=current_content,
                            height=300,
                            key=f"edit_{email_key}"
                        )

                        # Show preview of Gmail-style HTML with checkbox
                        show_preview = st.checkbox("👁️ Afficher l'aperçu Gmail-style", key=f"preview_{email_key}")
                        if show_preview:
                            st.iframe(edited_content, height=400)

                        # Save button
                        col1, col2 = st.columns([1, 3])
                        with col1:
                            if st.button("💾 Sauvegarder", key=f"save_{email_key}", type="primary"):
                                # Verify the edited content
                                is_valid, issues = st.session_state.email_automation.verify_email_content(edited_content)

                                if is_valid:
                                    # Save edited content
                                    st.session_state.edited_invalid_emails[email_key] = edited_content

                                    # Mark as validated
                                    if email_key not in st.session_state.validated_invalid_emails:
                                        st.session_state.validated_invalid_emails.append(email_key)

                                    # Update the email data
                                    email_data['personalized_email'] = edited_content
                                    email_data['is_valid'] = True
                                    email_data['issues'] = []

                                    st.success("✅ Email sauvegardé et validé!")
                                    st.rerun()
                                else:
                                    st.error("❌ L'email contient encore des problèmes:")
                                    for issue in issues:
                                        st.warning(f"⚠️ {issue}")

                        with col2:
                            if is_validated:
                                st.success("✅ Cet email a été validé et est prêt à être envoyé")

                # Show summary
                validated_count = len(st.session_state.validated_invalid_emails)
                if validated_count > 0:
                    st.info(f"📊 Progression: {validated_count}/{len(invalid_emails)} emails corrigés")

                # Button to send validated invalid emails
                if validated_count == len(invalid_emails) and validated_count > 0:
                    st.success("🎉 Tous les emails ont été corrigés!")

                    # Prepare validated emails for sending
                    validated_emails = []
                    for idx, email_data in enumerate(invalid_emails):
                        email_key = f"{email_data['email']}_{idx}"
                        if email_key in st.session_state.validated_invalid_emails:
                            # Update with edited content
                            if email_key in st.session_state.edited_invalid_emails:
                                email_data['personalized_email'] = st.session_state.edited_invalid_emails[email_key]
                                email_data['is_valid'] = True
                            validated_emails.append(email_data)

                    st.divider()
                    st.subheader("📤 Envoyer les emails corrigés")

                    # Gmail configuration check
                    if not all([sender_email, sender_password]):
                        st.error("⚠️ Configuration Gmail incomplète. Vérifiez la barre latérale.")
                    else:
                        # Show preview of emails to send
                        with st.expander(f"Aperçu des {len(validated_emails)} emails corrigés à envoyer"):
                            for email_data in validated_emails:
                                # Get display name and location with fallbacks
                                display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                location = email_data.get('site', email_data.get('Site', email_data.get('Location', 'N/A')))
                                st.write(f"**{display_name}** ({email_data['email']}) - {location} - [Gmail-style]")

                            # Show CC information
                            if cc_emails and cc_emails.strip():
                                cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
                                st.write(f"📋 **CC:** {', '.join(cc_list)}")

                        # Send button
                        if st.button("📤 Envoyer les emails corrigés", type="primary",
                                     key="send_invalid", disabled=not _send_allowed):
                            # First statement of the handler, before any smtplib call.
                            _presend_sync(sender_email)
                            progress_bar_invalid = st.progress(0)
                            status_text_invalid = st.empty()

                            sent_count = 0
                            failed_count = 0
                            skipped_suppressed_count = 0

                            try:
                                # Setup SMTP
                                server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                                server.starttls()
                                server.login(sender_email, sender_password)

                                # Load current progress
                                current_sent_emails = set()
                                if progress_file.exists():
                                    try:
                                        with open(progress_file, 'r', encoding='utf-8') as f:
                                            progress_data = json.load(f)
                                            if 'sent_emails' in progress_data:
                                                current_sent_emails = set(progress_data['sent_emails'])
                                    except:
                                        pass

                                decorative_image_file = st.session_state.get('decorative_image_file', None)
                                decorative_image_file_2 = st.session_state.get('decorative_image_file_2', None)
                                attachment_files = st.session_state.get('attachment_files', [])

                                for i, email_data in enumerate(validated_emails):
                                    # F3 — Never email a suppressed address. Test mode preserves
                                    # the real recipient in 'original_email', so we check that.
                                    suppression_target = email_data.get('original_email', email_data['email'])
                                    if is_suppressed(suppression_target):
                                        skipped_suppressed_count += 1
                                        progress_bar_invalid.progress((i + 1) / len(validated_emails))
                                        continue

                                    try:
                                        display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                        status_text_invalid.text(f"Envoi: {display_name} ({i+1}/{len(validated_emails)})")

                                        send_one_email(
                                            server, sender_email, email_data, cc_emails,
                                            st.session_state.email_automation,
                                            decorative_image_file=decorative_image_file,
                                            attachment_files=attachment_files,
                                            decorative_image_file_2=decorative_image_file_2,
                                        )
                                        sent_count += 1

                                        current_sent_emails.add(email_data['email'])
                                        progress_data = {
                                            'sent_emails': list(current_sent_emails),
                                            'last_update': datetime.now().isoformat()
                                        }
                                        with open(progress_file, 'w', encoding='utf-8') as f:
                                            json.dump(progress_data, f, indent=2)

                                    except Exception as e:
                                        display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                        st.error(f"Erreur envoi {display_name}: {e}")
                                        failed_count += 1

                                    progress_bar_invalid.progress((i + 1) / len(validated_emails))

                                    # F4 — Random delay (1–10s) between sends; none after the last.
                                    if i < len(validated_emails) - 1:
                                        time.sleep(random.uniform(MIN_DELAY_S, MAX_DELAY_S))

                                server.quit()

                                # Final status
                                if sent_count > 0:
                                    st.markdown('<div class="success-box">', unsafe_allow_html=True)
                                    st.success(f"✅ **{sent_count} emails corrigés envoyés avec succès!** 🎉")
                                    st.markdown('</div>', unsafe_allow_html=True)

                                    # Big green tick
                                    st.markdown("""
                                    <div style="text-align: center; margin: 20px 0;">
                                        <div style="font-size: 4rem; color: #4CAF50;">✅</div>
                                        <h3 style="color: #2E7D32; margin: 10px 0;">Envoi des emails corrigés terminé!</h3>
                                    </div>
                                    """, unsafe_allow_html=True)

                                    # Clear validated emails
                                    st.session_state.validated_invalid_emails = []
                                    st.session_state.edited_invalid_emails = {}

                                if failed_count > 0:
                                    st.error(f"❌ {failed_count} emails ont échoué")

                                if skipped_suppressed_count > 0:
                                    st.info(f"🚫 {skipped_suppressed_count} destinataire(s) ignoré(s) (présents dans la liste de désinscription)")

                                status_text_invalid.text("✅ Envoi terminé!")

                            except Exception as e:
                                st.error(f"Erreur de connexion Gmail: {e}")
                                st.info("💡 Vérifiez que vous utilisez un mot de passe d'application Gmail")

            if remaining_emails:
                st.subheader("✅ Emails prêts à envoyer")

                # Gmail configuration check
                if not all([sender_email, sender_password]):
                    st.error("⚠️ Configuration Gmail incomplète. Vérifiez la barre latérale.")
                else:
                    # Show preview of emails to send
                    with st.expander(f"Aperçu des {len(remaining_emails)} emails à envoyer"):
                        for email_data in remaining_emails[:5]:  # Show first 5
                            # Get display name and location with fallbacks
                            display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                            location = email_data.get('site', email_data.get('Site', email_data.get('Location', 'N/A')))
                            st.write(f"**{display_name}** ({email_data['email']}) - {location} - [Gmail-style]")
                        if len(remaining_emails) > 5:
                            st.write(f"... et {len(remaining_emails) - 5} autres")

                        # Show CC information
                        if cc_emails and cc_emails.strip():
                            cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
                            st.write(f"📋 **CC:** {', '.join(cc_list)}")

                    # Send emails - FIXED VERSION
                    if st.button("📤 Envoyer tous les emails", type="primary",
                                 disabled=not _send_allowed):
                        # First statement of the handler, before any smtplib call.
                        _presend_sync(sender_email)
                        progress_bar = st.progress(0)
                        status_text = st.empty()

                        sent_count = 0
                        failed_count = 0
                        skipped_suppressed_count = 0

                        try:
                            # Setup SMTP with hardcoded Gmail settings
                            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                            server.starttls()
                            server.login(sender_email, sender_password)

                            # Load current progress
                            current_sent_emails = set()
                            if progress_file.exists():
                                try:
                                    with open(progress_file, 'r', encoding='utf-8') as f:
                                        progress_data = json.load(f)
                                        if 'sent_emails' in progress_data:
                                            current_sent_emails = set(progress_data['sent_emails'])
                                except:
                                    pass

                            decorative_image_file = st.session_state.get('decorative_image_file', None)
                            decorative_image_file_2 = st.session_state.get('decorative_image_file_2', None)
                            attachment_files = st.session_state.get('attachment_files', [])

                            for i, email_data in enumerate(remaining_emails):
                                # F3 — Never email a suppressed address.
                                suppression_target = email_data.get('original_email', email_data['email'])
                                if is_suppressed(suppression_target):
                                    skipped_suppressed_count += 1
                                    progress_bar.progress((i + 1) / len(remaining_emails))
                                    continue

                                try:
                                    display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                    status_text.text(f"Envoi: {display_name} ({i+1}/{len(remaining_emails)})")

                                    send_one_email(
                                        server, sender_email, email_data, cc_emails,
                                        st.session_state.email_automation,
                                        decorative_image_file=decorative_image_file,
                                        attachment_files=attachment_files,
                                        decorative_image_file_2=decorative_image_file_2,
                                    )
                                    sent_count += 1

                                    current_sent_emails.add(email_data['email'])
                                    progress_data = {
                                        'sent_emails': list(current_sent_emails),
                                        'last_update': datetime.now().isoformat()
                                    }
                                    with open(progress_file, 'w', encoding='utf-8') as f:
                                        json.dump(progress_data, f, indent=2)

                                except Exception as e:
                                    display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                    st.error(f"Erreur envoi {display_name}: {e}")
                                    failed_count += 1

                                progress_bar.progress((i + 1) / len(remaining_emails))

                                # F4 — Random delay (1–10s) between sends; none after the last.
                                if i < len(remaining_emails) - 1:
                                    time.sleep(random.uniform(MIN_DELAY_S, MAX_DELAY_S))

                            server.quit()

                            # Clear progress file if all emails sent
                            if len(current_sent_emails) >= len(valid_emails):
                                if progress_file.exists():
                                    progress_file.unlink()
                                st.success("🎉 Tous les emails ont été envoyés! Le fichier de progrès a été effacé.")

                            # Final status with prominent green tick
                            if sent_count > 0:
                                st.markdown('<div class="success-box">', unsafe_allow_html=True)
                                st.success(f"✅ **{sent_count} emails envoyés avec succès!** 🎉")
                                st.markdown('</div>', unsafe_allow_html=True)

                                # Big green tick for visual confirmation
                                st.markdown("""
                                <div style="text-align: center; margin: 20px 0;">
                                    <div style="font-size: 4rem; color: #4CAF50;">✅</div>
                                    <h3 style="color: #2E7D32; margin: 10px 0;">Envoi Terminé avec Succès!</h3>
                                </div>
                                """, unsafe_allow_html=True)

                                if test_mode:
                                    st.success(f"✅ Mode test utilisé - {sent_count} emails envoyés à {sender_email} (pas aux clients)")
                                    st.info("💡 Désactivez le mode test pour envoyer aux vrais destinataires")

                            if failed_count > 0:
                                st.error(f"❌ {failed_count} emails ont échoué")

                            if skipped_suppressed_count > 0:
                                st.info(f"🚫 {skipped_suppressed_count} destinataire(s) ignoré(s) (présents dans la liste de désinscription)")


                            status_text.text("✅ Envoi terminé!")

                            # Show remaining count if any
                            remaining_after = len(valid_emails) - len(current_sent_emails)
                            if remaining_after > 0:
                                st.info(f"📊 {remaining_after} emails restants. Rechargez la page pour continuer.")

                        except Exception as e:
                            st.error(f"Erreur de connexion Gmail: {e}")
                            st.info("💡 Vérifiez que vous utilisez un mot de passe d'application Gmail")

            elif already_sent_count > 0:
                st.success(f"✅ Tous les emails ont déjà été envoyés! ({already_sent_count} emails)")
                if st.button("🗑️ Effacer le progrès", type="secondary"):
                    if progress_file.exists():
                        progress_file.unlink()
                    st.rerun()
            else:
                st.info("Aucun email valide prêt à envoyer.")
        else:
            st.info("Veuillez d'abord traiter les emails dans l'onglet Personnalisation.")

    # F4 — Opportunistic sweep, at the very END of main(): the page is already
    # painted when the IMAP connection starts, and the result shows on the next
    # render. Non-blocking anyway (it only starts a daemon thread), throttled to
    # one cycle per _UNSUB_THROTTLE_S whatever the number of reruns, and wrapped
    # because a failed sweep must never be able to break the page.
    try:
        _kick_unsub_sync(force=False)
    except Exception:
        pass


if __name__ == "__main__":
    main()
