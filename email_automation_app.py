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
import base64
import random
import html2text
import requests
from io import BytesIO
from datetime import datetime, timedelta

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

# Notion API Configuration
# Priority: st.secrets (Streamlit Cloud) > env var > .env file
try:
    NOTION_API_KEY = st.secrets["NOTION_API_KEY"]
except (KeyError, AttributeError):
    NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
    if not NOTION_API_KEY:
        _env_path = Path(__file__).parent / ".env"
        if _env_path.exists():
            for _line in _env_path.read_text().splitlines():
                if _line.startswith("NOTION_API_KEY="):
                    NOTION_API_KEY = _line.split("=", 1)[1].strip().strip('"')
                    break
NOTION_DS_ID = "285d9278-02d7-808a-9395-000b04dfc654"
NOTION_API_VERSION_LEGACY = "2022-06-28"
NOTION_API_VERSION_DS = "2025-09-03"


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
        self.html_template = """<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.4; color: #202124; background: #ffffff; margin: 0; padding: 0;">
  <p style="margin: 0 0 16px 0;">
    {first_paragraph}
  </p>

  {decorative_image_section}

  <p style="margin: 0 0 16px 0;">
    {second_paragraph}
  </p>

  {logo_section}
</div>"""
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

        return unique_contacts

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
        Supports: **bold**, *italic*, - bullet lists (with nesting), {color:name}text{/color}
        Also converts remaining \\n to <br> (outside of list blocks).
        """
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
        lines = text.split('\n')
        result_parts = []
        list_depth = 0
        list_buffer = []  # accumulate list HTML without newlines

        def _flush_list():
            """Flush accumulated list HTML as a single block (no internal newlines)."""
            nonlocal list_buffer
            if list_buffer:
                result_parts.append(''.join(list_buffer))
                list_buffer = []

        for line in lines:
            nested_match = re.match(r'^(?:    |\t)\s*- (.+)$', line)
            top_match = re.match(r'^- (.+)$', line.strip()) if not nested_match else None

            if nested_match:
                item_text = nested_match.group(1)
                if list_depth == 0:
                    list_buffer.append('<ul style="margin: 0; padding-left: 20px;">')
                    list_depth = 1
                if list_depth == 1:
                    list_buffer.append('<ul style="margin: 0; padding-left: 20px; list-style-type: circle;">')
                    list_depth = 2
                list_buffer.append(f'<li style="margin: 0; padding: 1px 0;">{item_text}</li>')
            elif top_match and line.strip().startswith('- '):
                item_text = top_match.group(1)
                if list_depth == 2:
                    list_buffer.append('</ul>')
                    list_depth = 1
                if list_depth == 0:
                    list_buffer.append('<ul style="margin: 0; padding-left: 20px;">')
                    list_depth = 1
                list_buffer.append(f'<li style="margin: 0; padding: 1px 0;">{item_text}</li>')
            else:
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

    def personalize_email(self, contact_data: Dict[str, str], email_content: str, use_html: bool = False,
                         logo_file=None, decorative_image_file=None, attachment_files=None, email_subject: str = "",
                         decorative_image_size: str = "100%", show_image_placeholder: bool = False) -> Tuple[str, str]:
        """
        Dynamic personalization with any column placeholders from Excel data
        Returns: (personalized_content, personalized_subject)
        decorative_image_size: CSS max-width for the main image (e.g. "280px", "100%").
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

            # Check if {Image} placeholder exists in content
            has_image_placeholder = '{Image}' in personalized

            # Build image style with chosen size (max-width)
            img_style = f"max-width: {decorative_image_size}; width: 100%; height: auto; border:0; outline:0; display: block;"

            # Wrapper so the image size box is always visible (same max-width + subtle border)
            image_wrapper_style = f"margin: 16px 0; max-width: {decorative_image_size}; width: 100%; box-sizing: border-box; border: 1px solid #e0e0e0; border-radius: 4px; overflow: hidden;"
            # Prepare decorative image section - only if no {Image} placeholder
            decorative_image_section = ""
            if decorative_image_file and not has_image_placeholder:
                decorative_image_section = f'''
                <div style="{image_wrapper_style}">
                <img src="cid:decorative_image" alt="Image" style="{img_style}">
                </div>'''
            elif show_image_placeholder and not decorative_image_file and not has_image_placeholder:
                # Preview: show a box of the chosen size when no image is uploaded
                decorative_image_section = f'''
                <div style="margin: 16px 0; max-width: {decorative_image_size}; width: 100%; min-height: 180px; background: #f5f5f5; border: 2px dashed #bdbdbd; border-radius: 4px; display: flex; align-items: center; justify-content: center; color: #757575; font-size: 13px; box-sizing: border-box;">
                Image ({decorative_image_size})
                </div>'''

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

            # Replace {Image} placeholder with actual image HTML or size box (same wrapper for visible box)
            if has_image_placeholder:
                if decorative_image_file:
                    image_html = f'''
                <div style="{image_wrapper_style}">
                <img src="cid:decorative_image" alt="Image" style="{img_style}">
                </div>'''
                else:
                    image_html = f'''
                <div style="margin: 16px 0; max-width: {decorative_image_size}; width: 100%; min-height: 180px; background: #f5f5f5; border: 2px dashed #bdbdbd; border-radius: 4px; display: flex; align-items: center; justify-content: center; color: #757575; font-size: 13px; box-sizing: border-box;">
                Image ({decorative_image_size})
                </div>''' if show_image_placeholder else ''
                if image_html:
                    first_paragraph = first_paragraph.replace('{Image}', image_html)
                    second_paragraph = second_paragraph.replace('{Image}', image_html)

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
                                 decorative_image_size: str = "100%", show_image_placeholder: bool = False) -> Tuple[str, str]:
        """AI personalization removed - using simple personalization instead"""
        return self.personalize_email(contact_data, email_content, use_html, logo_file, decorative_image_file, attachment_files, email_subject,
                                      decorative_image_size=decorative_image_size, show_image_placeholder=show_image_placeholder)

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


def _load_users() -> List[Dict[str, str]]:
    """Load user list: base users plus any from users.json (no duplicates by email)."""
    base = [
        {"name": "Salomé Cremona", "email": "salome.cremona@merciraymond.fr", "password": "kosj dkza wuku hlbo"},
        {"name": "Taddeo Carpinelli", "email": "taddeo.carpinelli@merciraymond.fr", "password": "tdcg uymo tswu urvk"},
        {"name": "Guillaume H.", "email": "guillaume@merciraymond.fr", "password": "ahlv pstg ibnv elsm"},
        {"name": "Clémence Joly", "email": "clemence@merciraymond.fr", "password": "clef gwtu cbrm vsry"},
        {"name": "Hugo Meunier", "email": "hugo@merciraymond.fr", "password": "rayq gdyj vaec jmrb"},
    ]
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
        st.sidebar.warning("Aucun utilisateur configuré")
        sender_email = None
        sender_password = None

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

                # Build editable DataFrame with checkboxes
                sites_df = pd.DataFrame({
                    "Selectionne": [True] * len(sites),
                    "Site": [s["site"] for s in sites],
                    "Email": [s["email"] for s in sites],
                })

                st.markdown(f"**{len(sites)} sites trouves.** Deselectionnez ceux que vous ne souhaitez pas inclure :")

                edited_sites = st.data_editor(
                    sites_df,
                    column_config={
                        "Selectionne": st.column_config.CheckboxColumn("Selectionne", default=True),
                        "Site": st.column_config.TextColumn("Site", disabled=True),
                        "Email": st.column_config.TextColumn("Email", disabled=True),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key="notion_sites_editor",
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
            help="Créez votre email complet ici. Utilisez des placeholders comme {contact_name}, {site}, etc. Formatage: **gras**, *italique*, - puces, {color:red}couleur{/color}.",
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
                    placeholder_text += "- `{Image}` : Place l'image décorative à cet endroit (remplace le placement automatique)"

                    st.markdown(placeholder_text)
                else:
                    st.markdown("""
                    **Placeholders spéciaux :**
                    - `{Image}` : Place l'image décorative à cet endroit (remplace le placement automatique)
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
                show_image_placeholder=True
            )
        # In preview iframe cid: doesn't work; inject image as base64 so it displays at correct size
        _preview_img = st.session_state.get('decorative_image_file')
        if _preview_img is not None:
            try:
                _preview_img.seek(0)
                _b64 = base64.b64encode(_preview_img.read()).decode('utf-8')
                _mime = _preview_img.type if getattr(_preview_img, 'type', None) else 'image/jpeg'
                _data_url = f"data:{_mime};base64,{_b64}"
                sample_html = sample_html.replace('src="cid:decorative_image"', f'src="{_data_url}"')
            except Exception:
                pass
        st.markdown(f'<p style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.4; color: #202124; margin: 0 0 16px 0;"><strong>Objet:</strong> {sample_subject}</p>', unsafe_allow_html=True)
        st.components.v1.html(sample_html, height=300, scrolling=True)

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
                        personalized, personalized_subject = st.session_state.email_automation.personalize_email(
                            contact_data, email_content_for_processing, use_html_for_processing,
                            logo_file=None, decorative_image_file=decorative_image_file_for_processing,
                            attachment_files=st.session_state.get('attachment_files', []),
                            email_subject=email_subject_for_processing,
                            decorative_image_size=_img_size_css,
                            show_image_placeholder=False
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
                            st.components.v1.html(edited_content, height=400, scrolling=True)

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
                        if st.button("📤 Envoyer les emails corrigés", type="primary", key="send_invalid"):
                            progress_bar_invalid = st.progress(0)
                            status_text_invalid = st.empty()

                            sent_count = 0
                            failed_count = 0

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

                                for i, email_data in enumerate(validated_emails):
                                    try:
                                        # Get display name with fallbacks
                                        display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                        status_text_invalid.text(f"Envoi: {display_name} ({i+1}/{len(validated_emails)})")

                                        # Create email with proper MIME structure
                                        msg_root = MIMEMultipart('mixed')
                                        msg_root['From'] = sender_email
                                        msg_root['To'] = email_data['email']
                                        msg_root['Subject'] = email_data.get('personalized_subject', 'MERCI RAYMOND - Votre service paysagiste')

                                        # Add CC if specified
                                        if cc_emails and cc_emails.strip():
                                            msg_root['Cc'] = cc_emails.strip()

                                        # Create alternative container for plain text and HTML
                                        alt = MIMEMultipart('alternative')
                                        msg_root.attach(alt)

                                        # Generate plain text version from HTML
                                        plain_text = html2text.html2text(email_data['personalized_email'])
                                        alt.attach(MIMEText(plain_text, 'plain', 'utf-8'))

                                        # Create related container for HTML and inline images
                                        rel = MIMEMultipart('related')
                                        rel.attach(MIMEText(email_data['personalized_email'], 'html', 'utf-8'))
                                        alt.attach(rel)

                                        # Add decorative image as inline attachment
                                        decorative_image_file = st.session_state.get('decorative_image_file', None)
                                        if decorative_image_file:
                                            try:
                                                # Compress decorative image before attaching
                                                compressed_decorative = st.session_state.email_automation.compress_image(decorative_image_file)
                                                image_attachment = MIMEImage(compressed_decorative.getvalue())
                                                image_attachment.add_header('Content-ID', '<decorative_image>')
                                                image_attachment.add_header('Content-Disposition', 'inline', filename='decorative_image.jpg')
                                                rel.attach(image_attachment)
                                            except Exception as e:
                                                st.warning(f"⚠️ Impossible d'ajouter l'image décorative: {e}")

                                        # Add regular attachments to root level
                                        attachment_files = st.session_state.get('attachment_files', [])
                                        if attachment_files:
                                            for attachment_file in attachment_files:
                                                try:
                                                    if attachment_file.type.startswith('image/'):
                                                        attachment = MIMEImage(attachment_file.getvalue())
                                                    else:
                                                        from email.mime.application import MIMEApplication
                                                        attachment = MIMEApplication(attachment_file.getvalue())

                                                    attachment.add_header(
                                                        'Content-Disposition',
                                                        'attachment',
                                                        filename=attachment_file.name
                                                    )
                                                    msg_root.attach(attachment)
                                                except Exception as e:
                                                    st.warning(f"⚠️ Impossible de joindre {attachment_file.name}: {e}")

                                        # Send email
                                        text = msg_root.as_string()

                                        # Prepare recipient list (TO + CC)
                                        recipients = [email_data['email']]
                                        if cc_emails and cc_emails.strip():
                                            cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
                                            recipients.extend(cc_list)

                                        server.sendmail(sender_email, recipients, text)
                                        sent_count += 1

                                        # Track sent email
                                        current_sent_emails.add(email_data['email'])

                                        # Save progress after each email
                                        progress_data = {
                                            'sent_emails': list(current_sent_emails),
                                            'last_update': datetime.now().isoformat()
                                        }
                                        with open(progress_file, 'w', encoding='utf-8') as f:
                                            json.dump(progress_data, f, indent=2)

                                    except Exception as e:
                                        # Get display name with fallbacks
                                        display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                        st.error(f"Erreur envoi {display_name}: {e}")
                                        failed_count += 1

                                    progress_bar_invalid.progress((i + 1) / len(validated_emails))

                                    # Anti-spam delay - hardcoded 1 second
                                    if i < len(validated_emails) - 1:
                                        time.sleep(1)

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
                    if st.button("📤 Envoyer tous les emails", type="primary"):
                        progress_bar = st.progress(0)
                        status_text = st.empty()

                        sent_count = 0
                        failed_count = 0

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

                            for i, email_data in enumerate(remaining_emails):
                                try:
                                    # Get display name with fallbacks
                                    display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                    status_text.text(f"Envoi: {display_name} ({i+1}/{len(remaining_emails)})")

                                    # Create email with proper MIME structure
                                    msg_root = MIMEMultipart('mixed')
                                    msg_root['From'] = sender_email
                                    msg_root['To'] = email_data['email']
                                    msg_root['Subject'] = email_data.get('personalized_subject', 'MERCI RAYMOND - Votre service paysagiste')

                                    # Add CC if specified
                                    if cc_emails and cc_emails.strip():
                                        msg_root['Cc'] = cc_emails.strip()

                                    # Create alternative container for plain text and HTML
                                    alt = MIMEMultipart('alternative')
                                    msg_root.attach(alt)

                                    # Generate plain text version from HTML
                                    plain_text = html2text.html2text(email_data['personalized_email'])
                                    alt.attach(MIMEText(plain_text, 'plain', 'utf-8'))

                                    # Create related container for HTML and inline images
                                    rel = MIMEMultipart('related')
                                    rel.attach(MIMEText(email_data['personalized_email'], 'html', 'utf-8'))
                                    alt.attach(rel)

                                    # Add decorative image as inline attachment
                                    decorative_image_file = st.session_state.get('decorative_image_file', None)
                                    if decorative_image_file:
                                        try:
                                            # Compress decorative image before attaching
                                            compressed_decorative = st.session_state.email_automation.compress_image(decorative_image_file)
                                            image_attachment = MIMEImage(compressed_decorative.getvalue())
                                            image_attachment.add_header('Content-ID', '<decorative_image>')
                                            image_attachment.add_header('Content-Disposition', 'inline', filename='decorative_image.jpg')
                                            rel.attach(image_attachment)
                                        except Exception as e:
                                            st.warning(f"⚠️ Impossible d'ajouter l'image décorative: {e}")

                                    # Add regular attachments to root level
                                    attachment_files = st.session_state.get('attachment_files', [])
                                    if attachment_files:
                                        for attachment_file in attachment_files:
                                            try:
                                                # Déterminer le type MIME
                                                if attachment_file.type.startswith('image/'):
                                                    attachment = MIMEImage(attachment_file.getvalue())
                                                else:
                                                    from email.mime.application import MIMEApplication
                                                    attachment = MIMEApplication(attachment_file.getvalue())

                                                attachment.add_header(
                                                    'Content-Disposition',
                                                    'attachment',
                                                    filename=attachment_file.name
                                                )
                                                msg_root.attach(attachment)
                                            except Exception as e:
                                                st.warning(f"⚠️ Impossible de joindre {attachment_file.name}: {e}")

                                    # Send email
                                    text = msg_root.as_string()

                                    # Prepare recipient list (TO + CC)
                                    recipients = [email_data['email']]
                                    if cc_emails and cc_emails.strip():
                                        cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
                                        recipients.extend(cc_list)

                                    server.sendmail(sender_email, recipients, text)
                                    sent_count += 1

                                    # Track sent email
                                    current_sent_emails.add(email_data['email'])

                                    # Save progress after each email
                                    progress_data = {
                                        'sent_emails': list(current_sent_emails),
                                        'last_update': datetime.now().isoformat()
                                    }
                                    with open(progress_file, 'w', encoding='utf-8') as f:
                                        json.dump(progress_data, f, indent=2)

                                except Exception as e:
                                    # Get display name with fallbacks
                                    display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                    st.error(f"Erreur envoi {display_name}: {e}")
                                    failed_count += 1

                                progress_bar.progress((i + 1) / len(remaining_emails))

                                # Anti-spam delay - hardcoded 1 second
                                if i < len(remaining_emails) - 1:  # Don't delay after last email
                                    time.sleep(1)

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

if __name__ == "__main__":
    main()
