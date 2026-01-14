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

import time
import base64
import random
import html2text
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

        return {
            'email_column': email_column,
            'available_placeholders': available_placeholders,
            'full_name_columns': full_name_columns,
            'all_columns': columns
        }

    def extract_contact_info(self, row: pd.Series, email_column: str, available_placeholders: Dict[str, str], full_name_columns: List[str] = None) -> Dict[str, str]:
        """Extract all contact information from a row dynamically."""
        info = {}

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

        # Extract first name and last name from full name columns
        if full_name_columns:
            for full_name_col in full_name_columns:
                if full_name_col in row.index and pd.notna(row[full_name_col]):
                    full_name = str(row[full_name_col]).strip()
                    name_parts = full_name.split()

                    # Add first name (first part)
                    if len(name_parts) > 0:
                        info[f"{full_name_col}_first"] = name_parts[0]

                    # Add last name (all parts after first, joined)
                    if len(name_parts) > 1:
                        info[f"{full_name_col}_last"] = ' '.join(name_parts[1:])
                    else:
                        info[f"{full_name_col}_last"] = ''

        return info

    def get_valid_emails_from_df(self, df: pd.DataFrame) -> List[Dict[str, str]]:
        """Extract all valid emails from the dataframe with dynamic column detection."""
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

        # Detect column mapping
        mapping = self.detect_column_mapping(df)
        email_column = mapping['email_column']
        available_placeholders = mapping['available_placeholders']
        full_name_columns = mapping['full_name_columns']

        valid_contacts = []

        for idx, row in df.iterrows():
            contact_info = self.extract_contact_info(row, email_column, available_placeholders, full_name_columns)

            # Check if we have a valid email
            if contact_info.get('email') and re.search(email_pattern, contact_info['email']):
                # Create contact with all available data
                contact_data = {
                    'index': idx,
                    'email': contact_info['email']
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
        """Convert markdown-style formatting to HTML for email"""
        # Convert **bold text** to <strong>bold text</strong>
        text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)

        # Convert *italic text* to <em>italic text</em>
        text = re.sub(r'\*(.*?)\*', r'<em>\1</em>', text)

        return text

    def personalize_email(self, contact_data: Dict[str, str], email_content: str, use_html: bool = False,
                         logo_file=None, decorative_image_file=None, attachment_files=None, email_subject: str = "") -> Tuple[str, str]:
        """
        Dynamic personalization with any column placeholders from Excel data
        Returns: (personalized_content, personalized_subject)
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

            # Prepare decorative image section - only if no {Image} placeholder
            decorative_image_section = ""
            if decorative_image_file and not has_image_placeholder:
                decorative_image_section = f'''
                <div style="margin: 16px 0;">
                <img src="cid:decorative_image" alt="Image" style="max-width: 100%; height: auto; border:0; outline:0; display: block;">
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

            # Convert line breaks to <br> tags for HTML
            first_paragraph = first_paragraph.replace('\n', '<br>')
            second_paragraph = second_paragraph.replace('\n', '<br>')

            # Convert markdown-style bold text to HTML
            first_paragraph = self.convert_markdown_to_html(first_paragraph)
            second_paragraph = self.convert_markdown_to_html(second_paragraph)

            # Replace {Image} placeholder with actual image HTML if it exists
            if has_image_placeholder and decorative_image_file:
                image_html = f'''
                <div style="margin: 16px 0;">
                <img src="cid:decorative_image" alt="Image" style="max-width: 100%; height: auto; border:0; outline:0; display: block;">
                </div>'''
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
                                 logo_file=None, decorative_image_file=None, attachment_files=None, email_subject: str = "") -> Tuple[str, str]:
        """AI personalization removed - using simple personalization instead"""
        return self.personalize_email(contact_data, email_content, use_html, logo_file, decorative_image_file, attachment_files, email_subject)

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

    # Sidebar for user selection
    st.sidebar.header("👤 Choisissez un Utilisateur")

    # Hardcoded user credentials
    USERS = [
        {"name": "Salomé Cremona", "email": "salome.cremona@merciraymond.fr", "password": "kosj dkza wuku hlbo"},
        {"name": "Taddeo Carpinelli", "email": "taddeo.carpinelli@merciraymond.fr", "password": "tdcg uymo tswu urvk"},
        {"name": "Guillaume H.", "email": "guillaume@merciraymond.fr", "password": "ahlv pstg ibnv elsm"}# Add more users as needed
    ]

    # Create user selection dropdown
    user_options = [user["name"] for user in USERS]
    if user_options:
        selected_user_name = st.sidebar.selectbox(
            "Utilisateur",
            options=user_options,
            help="Sélectionnez l'utilisateur pour l'envoi des emails"
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

    # Main content
    tab1, tab2, tab3 = st.tabs(["📁 Upload & Preview", "🎨 Design Email", "🚀 Envoi"])

    with tab1:
        st.markdown('<h2 class="step-header">Étape 1: Upload du fichier Excel</h2>', unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Choisissez votre fichier Excel",
            type=['xlsx', 'xls'],
            help="Le fichier peut contenir n'importe quelles colonnes - l'app détectera automatiquement les noms, emails, entreprises, etc."
        )

        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file)
                st.session_state.df = df

                st.success(f"✅ Fichier chargé avec succès! {len(df)} lignes trouvées.")

                # Show preview
                st.subheader("Aperçu des données")
                st.dataframe(df.head(10))

                # Detect column mapping and show it
                mapping = st.session_state.email_automation.detect_column_mapping(df)
                email_column = mapping['email_column']
                available_placeholders = mapping['available_placeholders']
                full_name_columns = mapping['full_name_columns']
                all_columns = mapping['all_columns']

                st.subheader("🔍 Détection automatique des colonnes")
                col1, col2 = st.columns(2)

                with col1:
                    st.write("**Colonnes détectées:**")

                    # Show detected email column
                    if email_column:
                        st.write(f"- 📧 **Email détecté:** `{email_column}`")
                    else:
                        st.write("- 📧 **Email:** ❌ Non détecté")
                        st.warning("⚠️ Aucune colonne email détectée. Veuillez vérifier votre fichier Excel.")

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
                                st.write(f"- **{base_name}:** `{{{base_name}}}` (nom complet), `{{{placeholders['first']}}}` (prénom), `{{{placeholders['last']}}}` (nom de famille)")
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
                        st.metric("Doublons retirés", duplicates_removed)
                        st.info(f"📧 {len(valid_contacts)} emails uniques (dont {duplicates_removed} doublons retirés)")
                    else:
                        st.info(f"📧 {len(valid_contacts)} emails uniques")

                # Show user guidance
                if email_column and available_placeholders:
                    placeholder_list = ", ".join([f"`{{{col}}}`" for col in available_placeholders.keys()])
                elif email_column:
                    st.info("💡 **Email détecté!** Vous pouvez personnaliser vos emails avec les données de cette colonne.")
                else:
                    st.error("❌ **Aucune colonne email détectée.** Vérifiez que votre fichier contient une colonne avec des adresses email.")



                # Store valid contacts for later use
                st.session_state.valid_contacts = valid_contacts

            except Exception as e:
                st.error(f"Erreur lors du chargement: {e}")

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
            help="Créez votre email complet ici. Utilisez des placeholders comme {contact_name}, {site}, etc. Utilisez **texte en gras** pour le texte en gras et *texte en italique* pour l'italique.",
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

            # Formatting explanation and example at the end
            st.markdown('<p style="margin-top: 1rem;"><strong>Formatage:</strong> Utilisez <code>**texte**</code> pour le gras et <code>*texte*</code> pour l\'italique.</p>', unsafe_allow_html=True)
            st.markdown('<p style="color: #2196F3; margin-top: 0.5rem;"><strong>Exemple:</strong> <code>**Bonjour**: *{contact_name}*</code> => <strong>Bonjour</strong>: <em>Marie</em></p>', unsafe_allow_html=True)

        # Store custom email content with format awareness
        if email_content and email_content.strip():
            st.session_state.custom_email_content = email_content
            st.session_state.custom_email_format = email_format
        else:
            st.session_state.custom_email_content = None
            st.session_state.custom_email_format = None

        st.divider()

        # Visual elements section
        decorative_image_file = st.file_uploader(
            "Image décorative",
            type=['png', 'jpg', 'jpeg'],
            help="Image qui apparaîtra dans le corps de l'email HTML"
        )

        attachment_files = st.file_uploader(
            "Fichiers à joindre",
            type=['pdf', 'doc', 'docx', 'xls', 'xlsx', 'png', 'jpg', 'jpeg', 'txt'],
            accept_multiple_files=True,
            help="Sélectionnez un ou plusieurs fichiers à joindre à l'email (PDF, Word, Excel, images, etc.)"
        )

        # Store files in session state
        if 'decorative_image_file' not in st.session_state:
            st.session_state.decorative_image_file = None
        if 'attachment_files' not in st.session_state:
            st.session_state.attachment_files = []

        if decorative_image_file:
            st.session_state.decorative_image_file = decorative_image_file
            st.success("✅ Image décorative chargée")

        if attachment_files:
            st.session_state.attachment_files = attachment_files
            st.success(f"✅ {len(attachment_files)} fichier(s) joint(s) chargé(s)")
            for i, file in enumerate(attachment_files):
                st.write(f"📎 {file.name} ({file.size} bytes)")

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
            email_subject=email_subject if email_subject else ""
            )
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
                        personalized, personalized_subject = st.session_state.email_automation.personalize_email(
                            contact_data, email_content_for_processing, use_html_for_processing,
                            logo_file=None, decorative_image_file=decorative_image_file_for_processing,
                            attachment_files=st.session_state.get('attachment_files', []),
                            email_subject=email_subject_for_processing
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
            st.info("Veuillez d'abord charger un fichier Excel dans l'onglet 'Upload & Preview'.")

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

            # Get valid_contacts count
            valid_contacts_count = len(st.session_state.valid_contacts) if hasattr(st.session_state, 'valid_contacts') else len(valid_emails)

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Emails prêts", valid_contacts_count)
            with col2:
                st.metric("Emails avec problèmes", len(invalid_emails))
            with col3:
                if valid_emails:
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

            if valid_emails:
                st.subheader("✅ Emails prêts à envoyer")

                # Gmail configuration check
                if not all([sender_email, sender_password]):
                    st.error("⚠️ Configuration Gmail incomplète. Vérifiez la barre latérale.")
                else:
                    # Show preview of emails to send
                    with st.expander(f"Aperçu des {len(valid_emails)} emails à envoyer"):
                        for email_data in valid_emails[:5]:  # Show first 5
                            # Get display name and location with fallbacks
                            display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                            location = email_data.get('site', email_data.get('Site', email_data.get('Location', 'N/A')))
                            st.write(f"**{display_name}** ({email_data['email']}) - {location} - [Gmail-style]")
                        if len(valid_emails) > 5:
                            st.write(f"... et {len(valid_emails) - 5} autres")

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

                            for i, email_data in enumerate(valid_emails):
                                try:
                                    # Get display name with fallbacks
                                    display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                    status_text.text(f"Envoi: {display_name} ({i+1}/{len(valid_emails)})")

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

                                except Exception as e:
                                    # Get display name with fallbacks
                                    display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                                    st.error(f"Erreur envoi {display_name}: {e}")
                                    failed_count += 1

                                progress_bar.progress((i + 1) / len(valid_emails))

                                # Anti-spam delay - hardcoded 1 second
                                if i < len(valid_emails) - 1:  # Don't delay after last email
                                    time.sleep(1)

                            server.quit()

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

                        except Exception as e:
                            st.error(f"Erreur de connexion Gmail: {e}")
                            st.info("💡 Vérifiez que vous utilisez un mot de passe d'application Gmail")

            else:
                st.info("Aucun email valide prêt à envoyer.")
        else:
            st.info("Veuillez d'abord traiter les emails dans l'onglet Personnalisation.")

if __name__ == "__main__":
    main()
