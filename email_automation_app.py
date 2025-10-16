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
        self.base_email_content_text = """Bonjour {contact_name},

Lorsque l'été touche à sa fin et l'hiver arrive à pas feutrés…

Les Raymonds vous emmènent dans leur traîneau et vous proposent une large palette de sapins, décorations et animations afin de préparer l'arrivée des fêtes de fin d'année !

Créez une ambiance unique avec des sapins, robustes et élégants, disponibles de 80 à 200 cm, et décorés selon vos préférences. Découvrez également nos guirlandes sur mesure, faites de branchages et personnalisables.

Vous connaissez les Raymonds, ces décorations 100% végétales seront aussi festives que durables ! Matériaux sourcés & Réemploi en intégralité.
Remplissez la lettre au père noël jointe à ce mail avec vos désirs de couleurs et vos choix de dimensions, et faisons germer ensemble l'esprit de Noël dans votre {site} !

🎁 Pour l'occasion, nous avons le plaisir d'offrir à nos clients du pôle entretien une réduction spéciale de 10 % sur notre catalogue de Noël.

Je reste à votre entière disposition pour tout complément d'information ou pour une offre sur mesure.

En vous souhaitant une bonne journée,

L'équipe MERCI RAYMOND"""

        # Base email template for Gmail-style HTML format
        self.base_email_content_html = """Lorsque l'été touche à sa fin et l'hiver arrive à pas feutrés…

Les Raymonds vous emmènent dans leur traîneau et vous proposent une large palette de sapins, décorations et animations afin de préparer l'arrivée des fêtes de fin d'année !

Créez une ambiance unique avec des sapins, robustes et élégants, disponibles de 80 à 200 cm, et décorés selon vos préférences. Découvrez également nos guirlandes sur mesure, faites de branchages et personnalisables.

Vous connaissez les Raymonds, ces décorations 100% végétales seront aussi festives que durables ! Matériaux sourcés & Réemploi en intégralité.
Remplissez la lettre au père noël jointe à ce mail avec vos désirs de couleurs et vos choix de dimensions, et faisons germer ensemble l'esprit de Noël dans votre {site} !

🎁 **Pour l'occasion**, nous avons le plaisir d'offrir à nos clients du pôle entretien une **réduction spéciale de 10%** sur notre catalogue de Noël.

Je reste à votre entière disposition pour tout complément d'information ou pour une offre sur mesure."""

        # Gmail-style HTML template that looks like plain text
        self.html_template = """<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.4; color: #202124; background: #ffffff; margin: 0; padding: 0;">
  {header_section}

  <p style="margin: 0 0 16px 0;">
    {first_paragraph}
  </p>

  {decorative_image_section}

  <p style="margin: 0 0 16px 0;">
    {second_paragraph}
  </p>

  {footer_section}

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

        # Create available placeholders (all columns except email)
        available_placeholders = {}
        for col in columns:
            if col != email_column:
                available_placeholders[col] = col

        return {
            'email_column': email_column,
            'available_placeholders': available_placeholders,
            'all_columns': columns
        }

    def extract_contact_info(self, row: pd.Series, email_column: str, available_placeholders: Dict[str, str]) -> Dict[str, str]:
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

        return info

    def get_valid_emails_from_df(self, df: pd.DataFrame) -> List[Dict[str, str]]:
        """Extract all valid emails from the dataframe with dynamic column detection."""
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

        # Detect column mapping
        mapping = self.detect_column_mapping(df)
        email_column = mapping['email_column']
        available_placeholders = mapping['available_placeholders']

        valid_contacts = []

        for idx, row in df.iterrows():
            contact_info = self.extract_contact_info(row, email_column, available_placeholders)

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

        return valid_contacts

    def encode_image_to_base64(self, image_file) -> Optional[str]:
        """Convert uploaded image to base64 for embedding in HTML"""
        try:
            return base64.b64encode(image_file.getvalue()).decode()
        except:
            return None

    def convert_markdown_to_html(self, text: str) -> str:
        """Convert markdown-style formatting to HTML for email"""
        # Convert **bold text** to <strong>bold text</strong>
        text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)

        # Convert *italic text* to <em>italic text</em>
        text = re.sub(r'\*(.*?)\*', r'<em>\1</em>', text)

        return text

    def personalize_email(self, contact_data: Dict[str, str], email_content: str, use_html: bool = False,
                         logo_file=None, decorative_image_file=None, attachment_files=None) -> str:
        """
        Dynamic personalization with any column placeholders from Excel data
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

        if use_html:
            # Prepare logo section - small signature-style image
            logo_section = ""
            if logo_file:
                logo_section = f'<img src="cid:logo" alt="Merci Raymond" style="display:inline-block; height:24px; width:auto; border:0; outline:0; vertical-align:baseline;">'

            # Prepare decorative image section - natural Gmail-style placement
            decorative_image_section = ""
            if decorative_image_file:
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

            # Get custom header and footer from session state
            header_content = st.session_state.get('email_header', 'Bonjour {contact_name}, j\'espère que vous allez bien.')
            footer_content = st.session_state.get('email_footer', 'Bien cordialement,\nSalomé Cremona')

            # Process header and footer with placeholders
            header_processed = header_content
            footer_processed = footer_content

            # Replace placeholders in header and footer
            for key, value in contact_data.items():
                if key != 'email' and key != 'index':
                    placeholder = f"{{{key}}}"
                    replacement_value = value if value else ""
                    header_processed = header_processed.replace(placeholder, replacement_value)
                    footer_processed = footer_processed.replace(placeholder, replacement_value)

            # Replace {contact_name} in header and footer
            header_processed = header_processed.replace('{contact_name}', first_name)
            footer_processed = footer_processed.replace('{contact_name}', first_name)

            # Convert line breaks to <br> tags for HTML
            header_processed = header_processed.replace('\n', '<br>')
            footer_processed = footer_processed.replace('\n', '<br>')

            # Convert markdown-style formatting to HTML
            header_processed = self.convert_markdown_to_html(header_processed)
            footer_processed = self.convert_markdown_to_html(footer_processed)

            # Wrap header and footer in proper HTML paragraphs
            header_section = f'<p style="margin: 0 0 16px 0;">{header_processed}</p>'
            footer_section = f'<p style="margin: 0 0 16px 0;">{footer_processed}</p>'

            # Apply Gmail-style HTML template
            personalized = self.html_template.format(
                header_section=header_section,
                first_paragraph=first_paragraph,
                second_paragraph=second_paragraph,
                footer_section=footer_section,
                logo_section=logo_section,
                decorative_image_section=decorative_image_section
            )

        return personalized

    def personalize_email_with_ai(self, contact_data: Dict[str, str], email_content: str, use_html: bool = False,
                                 logo_file=None, decorative_image_file=None, attachment_files=None) -> str:
        """AI personalization removed - using simple personalization instead"""
        return self.personalize_email(contact_data, email_content, use_html, logo_file, decorative_image_file, attachment_files)

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
    st.markdown('<h1 class="main-header">🌱 MERCI RAYMOND - Automation Email</h1>', unsafe_allow_html=True)

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

    # Sidebar for Gmail configuration
    st.sidebar.header("📧 Configuration Gmail")

    # Instructions for Gmail setup
#    st.sidebar.markdown('<div class="info-box">', unsafe_allow_html=True)
    st.sidebar.markdown("""
    **📋 Instructions Gmail :**
    1. Activez l'authentification à 2 facteurs
    2. Générez un "Mot de passe d'application"
    3. Utilisez ce mot de passe ci-dessous
    """)
    st.sidebar.markdown('</div>', unsafe_allow_html=True)

    sender_email = st.sidebar.text_input(
        "Adresse Gmail",
        placeholder="votre@gmail.com ou votre@merciraymond.fr",
        help="Votre adresse Gmail ou adresse configurée avec Gmail"
    )
    sender_password = st.sidebar.text_input(
        "Mot de passe d'application Gmail",
        type="password",
        help="Généré dans les paramètres de sécurité Gmail, PAS votre mot de passe habituel"
    )

    # CC email addresses
    st.sidebar.subheader("📋 CC (optionnel)")
    cc_emails = st.sidebar.text_input(
        "Adresses email en copie",
        placeholder="email1@example.com, email2@example.com",
        help="Séparez plusieurs adresses par des virgules. Ces emails recevront une copie de tous les emails envoyés."
    )

    # Anti-spam settings
    st.sidebar.subheader("🛡️ Protection Anti-spam")

#    st.sidebar.markdown('<div class="anti-spam-box">', unsafe_allow_html=True)
    st.sidebar.markdown("""
    **⚠️ Conseils anti-spam :**
    - Échelonner les envois
    - Tester d'abord sur quelques emails
    - Personnaliser au maximum
    """)
    st.sidebar.markdown('</div>', unsafe_allow_html=True)

    # Random delay between 8-20 seconds for better anti-spam protection
    delay_between_emails = random.randint(8, 20)

    test_mode = st.sidebar.checkbox(
        "Mode test",
        help="Envoyer seulement aux 5 premiers emails pour tester"
    )

    # Show system status
    st.sidebar.success("🔧 Système simple et fiable - Personnalisation robuste")

    # Show CC status
    if cc_emails and cc_emails.strip():
        cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
        st.sidebar.info(f"📋 CC configuré: {len(cc_list)} adresse(s)")
        for cc_email in cc_list:
            st.sidebar.write(f"  • {cc_email}")
    else:
        st.sidebar.info("📋 Aucun CC configuré")

    # Main content
    tab1, tab2, tab3, tab4 = st.tabs(["📁 Upload & Preview", "🎨 Design Email", "✉️ Personnalisation", "🚀 Envoi"])

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
                        for col_name in available_placeholders.keys():
                            st.write(f"- `{{{col_name}}}`")
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

                # Show user guidance
                if email_column and available_placeholders:
                    placeholder_list = ", ".join([f"`{{{col}}}`" for col in available_placeholders.keys()])
                    st.info(f"💡 **Vous pouvez utiliser ces placeholders dans votre email:** {placeholder_list}")
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
        st.subheader("📝 Objet de l'email")
        email_subject = st.text_input(
            "Objet de l'email:",
            value="MERCI RAYMOND - Votre service paysagiste",
            help="L'objet de l'email qui apparaîtra dans la boîte de réception",
            key=f"email_subject_{email_format}"  # Unique key per format
        )

        # Store email subject in session state
        st.session_state.email_subject = email_subject

        # Show subject line info
        st.info(f"📧 **Objet configuré:** {email_subject}")

        st.divider()

        # Allow user to modify email content
        email_content = st.text_area(
            "Modifiez le contenu de votre email:",
            value=base_template,
            height=300,
            help="Utilisez {site} comme placeholder pour la personnalisation du lieu. Utilisez **texte en gras** pour le texte en gras et *texte en italique* pour l'italique.",
            key=f"email_content_{email_format}"  # Unique key per format
        )

        # Show formatting help
        with st.expander("💡 Aide au formatage du texte"):
            st.markdown("""
            **Formatage du texte disponible :**

            - **Texte en gras** : Utilisez `**texte en gras**`
            - *Texte en italique* : Utilisez `*texte en italique*`

            **Exemples :**
            - `**Offre spéciale**` → **Offre spéciale**
            - `*Réduction limitée*` → *Réduction limitée*
            - `**Important** : *Action limitée*` → **Important** : *Action limitée*
            """)

            # Show dynamic placeholders if Excel file is uploaded
            if st.session_state.df is not None:
                mapping = st.session_state.email_automation.detect_column_mapping(st.session_state.df)
                available_placeholders = mapping['available_placeholders']

                if available_placeholders:
                    placeholder_list = ", ".join([f"`{{{col}}}`" for col in available_placeholders.keys()])
                    st.markdown(f"""
                    **Placeholders disponibles depuis votre Excel :**
                    {placeholder_list}

                    **Placeholders spéciaux :**
                    - `{{contact_name}}` : Remplacé par le prénom du contact
                    """)
                else:
                    st.markdown("""
                    **Placeholders spéciaux :**
                    - `{contact_name}` : Remplacé par le prénom du contact
                    """)
            else:
                st.markdown("""
                **Placeholders disponibles :**
                - `{contact_name}` : Remplacé par le nom du contact
                - `{site}` : Remplacé par le lieu du contact

                *Chargez un fichier Excel pour voir tous les placeholders disponibles*
                """)

        # Store custom email content with format awareness
        if email_content != base_template:
            st.session_state.custom_email_content = email_content
            st.session_state.custom_email_format = email_format
        else:
            st.session_state.custom_email_content = None
            st.session_state.custom_email_format = None

        st.divider()

        # Header and Footer Customization
        st.subheader("📝 En-tête et Signature personnalisés")

        col1, col2 = st.columns(2)

        with col1:
            st.write("**📧 En-tête de l'email:**")
            header_content = st.text_area(
                "En-tête (salutation):",
                value=st.session_state.get('email_header', "Bonjour {contact_name}, j'espère que vous allez bien."),
                height=100,
                help="Utilisez {contact_name} pour le prénom du contact, ou tout autre placeholder disponible.",
                key="header_input"
            )

        with col2:
            st.write("**✍️ Signature de l'email:**")
            footer_content = st.text_area(
                "Signature (formule de politesse):",
                value=st.session_state.get('email_footer', "Bien cordialement,\nSalomé Cremona"),
                height=100,
                help="Votre signature personnalisée. Vous pouvez utiliser des placeholders comme {contact_name}.",
                key="footer_input"
            )

        # Store header and footer in session state with different keys
        if header_content:
            st.session_state.email_header = header_content
        if footer_content:
            st.session_state.email_footer = footer_content

        # Show header/footer help
        with st.expander("💡 Aide pour l'en-tête et la signature"):
            st.markdown("""
            **Placeholders disponibles pour l'en-tête et la signature :**

            - `{contact_name}` : Prénom du contact
            - `{Name}` ou `{Full Name}` : Nom complet du contact
            - `{Company}` ou `{Company Name}` : Nom de l'entreprise
            - `{Site}` ou `{Location}` : Lieu/site du contact

            **Exemples d'en-têtes :**
            - `Bonjour {contact_name}, j'espère que vous allez bien.`
            - `Cher(e) {Name},`
            - `Madame/Monsieur {contact_name},`

            **Exemples de signatures :**
            - `Bien cordialement,\nVotre nom`
            - `Cordialement,\n{contact_name} de l'équipe MERCI RAYMOND`
            - `Avec mes salutations distinguées,\nVotre équipe`
            """)

        st.divider()

        # Visual elements section
        st.subheader("🎨 Éléments visuels (optionnels)")

        logo_file = st.file_uploader(
            "Logo de votre entreprise",
            type=['png', 'jpg', 'jpeg'],
            help="Logo qui apparaîtra en en-tête de l'email HTML"
        )

        decorative_image_file = st.file_uploader(
            "Image décorative",
            type=['png', 'jpg', 'jpeg'],
            help="Image qui apparaîtra dans le corps de l'email HTML"
        )

        # Nouvelle section pour les pièces jointes
        st.subheader("📎 Pièces jointes (optionnelles)")

        attachment_files = st.file_uploader(
            "Fichiers à joindre",
            type=['pdf', 'doc', 'docx', 'xls', 'xlsx', 'png', 'jpg', 'jpeg', 'txt'],
            accept_multiple_files=True,
            help="Sélectionnez un ou plusieurs fichiers à joindre à l'email (PDF, Word, Excel, images, etc.)"
        )

        # Store files in session state
        if 'logo_file' not in st.session_state:
            st.session_state.logo_file = None
        if 'decorative_image_file' not in st.session_state:
            st.session_state.decorative_image_file = None
        if 'attachment_files' not in st.session_state:
            st.session_state.attachment_files = []

        if logo_file:
            st.session_state.logo_file = logo_file
            st.success("✅ Logo chargé")

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
        # Create sample contact data for preview
        sample_contact = {
            'contact_name': 'Marie Dupont',
            'Site': 'Bureau Paris',
            'Company': 'Entreprise ABC'
        }
        sample_html = st.session_state.email_automation.personalize_email(
                sample_contact, email_content, use_html=True,
                logo_file=st.session_state.logo_file,
            decorative_image_file=st.session_state.decorative_image_file,
            attachment_files=st.session_state.get('attachment_files', [])
            )
        st.components.v1.html(sample_html, height=500, scrolling=True)



    with tab3:
        st.markdown('<h2 class="step-header">Étape 3: Personnalisation des emails</h2>', unsafe_allow_html=True)

        if st.session_state.df is not None and hasattr(st.session_state, 'valid_contacts'):
            # Get settings from previous tab - always Gmail-style HTML
            email_format = st.session_state.get('email_format', 'HTML (Gmail-style)')
            is_html_format = True

            # Ensure we always have email content - use Gmail-style template
            if 'custom_email_content' in st.session_state and st.session_state.custom_email_content:
                # Verify the custom content matches the current format
                custom_format = st.session_state.get('custom_email_format', email_format)
                if custom_format == email_format:
                    email_content = st.session_state.custom_email_content
                else:
                    # Format changed, use base template for Gmail-style
                    email_content = st.session_state.email_automation.base_email_content_html
            else:
                # No custom content, use base template for Gmail-style
                email_content = st.session_state.email_automation.base_email_content_html

            logo_file = st.session_state.get('logo_file', None)
            decorative_image_file = st.session_state.get('decorative_image_file', None)

            # Determine email formats to generate - always Gmail-style HTML
            formats_to_generate = [("gmail-style", True)]

            valid_contacts = st.session_state.valid_contacts

            if valid_contacts:
                st.info(f"📊 {len(valid_contacts)} emails à traiter")

                # Preview section
                st.subheader("Aperçu de la personnalisation")

                preview_idx = st.selectbox(
                    "Choisir un contact pour l'aperçu:",
                    range(len(valid_contacts)),
                    format_func=lambda x: f"{valid_contacts[x].get('contact_name', valid_contacts[x].get('Name', 'Contact'))} - {valid_contacts[x].get('site', valid_contacts[x].get('Site', valid_contacts[x].get('Location', 'N/A')))}"
                )

                if preview_idx is not None:
                    selected_contact = valid_contacts[preview_idx]

                    col1, col2 = st.columns(2)

                    with col1:
                        st.write("**Informations du contact:**")
                        st.write(f"- Email: {selected_contact['email']}")

                        # Show all available data fields dynamically
                        for key, value in selected_contact.items():
                            if key not in ['email', 'index']:  # Skip email and index
                                st.write(f"- {key}: {value}")

                    with col2:
                        st.info("🔧 Personnalisation Gmail-style avec placeholders dynamiques")
                        personalization_method = "Gmail-style"

                    if st.button("Générer aperçu"):
                        st.markdown('<div class="email-preview">', unsafe_allow_html=True)

                        for format_name, use_html in formats_to_generate:
                            # Always use simple personalization - reliable and bulletproof
                            personalized_email = st.session_state.email_automation.personalize_email(
                                selected_contact, email_content, use_html,
                                logo_file, decorative_image_file,
                                attachment_files=st.session_state.get('attachment_files', [])
                            )

                            st.markdown("**📧 Aperçu Gmail-style personnalisé:**")

                            # Always show as HTML for Gmail-style
                            st.components.v1.html(personalized_email, height=500, scrolling=True)

                            # Verification
                            is_valid, issues = st.session_state.email_automation.verify_email_content(personalized_email)

                            if is_valid:
                                st.success("✅ Email Gmail-style validé - Prêt à envoyer")
                            else:
                                st.warning("⚠️ Problèmes détectés - Vérifiez le contenu")

                        st.markdown('</div>', unsafe_allow_html=True)

                # Process all emails
                if st.button("🔄 Traiter tous les emails", type="primary"):
                    # Always use Gmail-style HTML
                    use_html_for_processing = True

                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    processed_emails = []

                    for i, contact_data in enumerate(valid_contacts):
                        # Get a display name for status (use first available field or email)
                        display_name = contact_data.get('contact_name', contact_data.get('Name', contact_data.get('email', 'Contact')))
                        status_text.text(f"Traitement: {display_name} ({i+1}/{len(valid_contacts)})")

                        # Always use simple personalization - reliable and bulletproof
                        personalized = st.session_state.email_automation.personalize_email(
                            contact_data, email_content, use_html_for_processing,
                            logo_file, decorative_image_file,
                            attachment_files=st.session_state.get('attachment_files', [])
                        )

                        is_valid, issues = st.session_state.email_automation.verify_email_content(personalized)

                        processed_emails.append({
                            **contact_data,
                            'personalized_email': personalized,
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

    with tab4:
        st.markdown('<h2 class="step-header">Étape 4: Envoi des emails</h2>', unsafe_allow_html=True)

        if st.session_state.processed_emails:
            processed_emails = st.session_state.processed_emails
            valid_emails = [email for email in processed_emails if email['is_valid']]
            invalid_emails = [email for email in processed_emails if not email['is_valid']]

            # Apply test mode filter
            if test_mode:
                valid_emails = valid_emails[:5]
                st.info("🧪 Mode test activé - Envoi limité aux 5 premiers emails")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Emails prêts", len(valid_contacts))
            with col2:
                st.metric("Emails avec problèmes", len(invalid_emails))
            with col3:
                if valid_emails:
                    # Use average delay (14 seconds) for time estimation
                    avg_delay = 14
                    sending_time = calculate_sending_time(len(valid_contacts), avg_delay)
                    st.metric("Temps d'envoi", sending_time)

            # Anti-spam recommendations
            st.markdown(f"""
            **🛡️ Configuration anti-spam active :**
            - ⏱️ Délai entre emails : 8-20 secondes (aléatoire)
            - 🧪 Mode test : {'Activé (5 emails max)' if test_mode else 'Désactivé'}
            - 📧 Emails à envoyer : {len(valid_contacts)}
            - ⏰ Temps total estimé : {calculate_sending_time(len(valid_contacts), 14)}
            """)
            st.markdown('</div>', unsafe_allow_html=True)

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
                        # Subject line for invalid emails - use stored subject or default
                        stored_subject_invalid = st.session_state.get('email_subject', 'MERCI RAYMOND - Votre service paysagiste')
                        email_subject_invalid = st.text_input(
                            "Objet de l'email (emails corrigés):",
                            value=stored_subject_invalid,
                            key="subject_invalid",
                            help="L'objet défini dans l'onglet Design Email"
                        )

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

                                        # Create email
                                        msg = MIMEMultipart('alternative')
                                        msg['From'] = sender_email
                                        msg['To'] = email_data['email']
                                        msg['Subject'] = email_subject_invalid

                                        # Add CC if specified
                                        if cc_emails and cc_emails.strip():
                                            msg['Cc'] = cc_emails.strip()

                                        # Generate plain text version from HTML
                                        plain_text = html2text.html2text(email_data['personalized_email'])

                                        # Add plain text first (for spam filters and accessibility)
                                        msg.attach(MIMEText(plain_text, 'plain', 'utf-8'))

                                        # Add Gmail-style HTML body
                                        msg.attach(MIMEText(email_data['personalized_email'], 'html', 'utf-8'))

                                        # Add images as inline attachments (Gmail-style HTML)
                                        # Always add images for Gmail-style format

                                        # Add logo as inline attachment
                                        logo_file = st.session_state.get('logo_file', None)
                                        if logo_file:
                                            try:
                                                logo_attachment = MIMEImage(logo_file.getvalue())
                                                logo_attachment.add_header('Content-ID', '<logo>')
                                                logo_attachment.add_header('Content-Disposition', 'inline', filename='logo.png')
                                                msg.attach(logo_attachment)
                                            except Exception as e:
                                                st.warning(f"⚠️ Impossible d'ajouter le logo: {e}")

                                        # Add decorative image as inline attachment
                                        decorative_image_file = st.session_state.get('decorative_image_file', None)
                                        if decorative_image_file:
                                            try:
                                                image_attachment = MIMEImage(decorative_image_file.getvalue())
                                                image_attachment.add_header('Content-ID', '<decorative_image>')
                                                image_attachment.add_header('Content-Disposition', 'inline', filename='decorative_image.png')
                                                msg.attach(image_attachment)
                                            except Exception as e:
                                                st.warning(f"⚠️ Impossible d'ajouter l'image décorative: {e}")

                                        # Add regular attachments
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
                                                    msg.attach(attachment)
                                                except Exception as e:
                                                    st.warning(f"⚠️ Impossible de joindre {attachment_file.name}: {e}")

                                        # Send email
                                        text = msg.as_string()

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

                                    # Anti-spam delay (random 8-20 seconds)
                                    if i < len(validated_emails) - 1:
                                        time.sleep(random.randint(8, 20))

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
                    # Subject line - use stored subject or default
                    stored_subject = st.session_state.get('email_subject', 'MERCI RAYMOND - Votre service paysagiste')
                    email_subject = st.text_input(
                        "Objet de l'email:",
                        value=stored_subject,
                        help="L'objet défini dans l'onglet Design Email"
                    )

                    # Show preview of emails to send
                    with st.expander(f"Aperçu des {len(valid_contacts)} emails à envoyer"):
                        for email_data in valid_emails[:5]:  # Show first 5
                            # Get display name and location with fallbacks
                            display_name = email_data.get('contact_name', email_data.get('Name', email_data.get('Full Name', 'Contact')))
                            location = email_data.get('site', email_data.get('Site', email_data.get('Location', 'N/A')))
                            st.write(f"**{display_name}** ({email_data['email']}) - {location} - [Gmail-style]")
                        if len(valid_contacts) > 5:
                            st.write(f"... et {len(valid_contacts) - 5} autres")

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

                                    # Create email
                                    msg = MIMEMultipart('alternative')
                                    msg['From'] = sender_email
                                    msg['To'] = email_data['email']
                                    msg['Subject'] = email_subject

                                    # Add CC if specified
                                    if cc_emails and cc_emails.strip():
                                        msg['Cc'] = cc_emails.strip()

                                    # Generate plain text version from HTML
                                    plain_text = html2text.html2text(email_data['personalized_email'])

                                    # Add plain text first (for spam filters and accessibility)
                                    msg.attach(MIMEText(plain_text, 'plain', 'utf-8'))

                                    # Add Gmail-style HTML body
                                    msg.attach(MIMEText(email_data['personalized_email'], 'html', 'utf-8'))

                                    # Add images as inline attachments (Gmail-style HTML)
                                    # Always add images for Gmail-style format

                                    # Add logo as inline attachment
                                    logo_file = st.session_state.get('logo_file', None)
                                    if logo_file:
                                        try:
                                            logo_attachment = MIMEImage(logo_file.getvalue())
                                            logo_attachment.add_header('Content-ID', '<logo>')
                                            logo_attachment.add_header('Content-Disposition', 'inline', filename='logo.png')
                                            msg.attach(logo_attachment)
                                        except Exception as e:
                                            st.warning(f"⚠️ Impossible d'ajouter le logo: {e}")

                                    # Add decorative image as inline attachment
                                    decorative_image_file = st.session_state.get('decorative_image_file', None)
                                    if decorative_image_file:
                                        try:
                                            image_attachment = MIMEImage(decorative_image_file.getvalue())
                                            image_attachment.add_header('Content-ID', '<decorative_image>')
                                            image_attachment.add_header('Content-Disposition', 'inline', filename='decorative_image.png')
                                            msg.attach(image_attachment)
                                        except Exception as e:
                                            st.warning(f"⚠️ Impossible d'ajouter l'image décorative: {e}")

                                    # Add regular attachments
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
                                                msg.attach(attachment)
                                            except Exception as e:
                                                st.warning(f"⚠️ Impossible de joindre {attachment_file.name}: {e}")

                                    # Send email
                                    text = msg.as_string()

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

                                # Anti-spam delay (random 8-20 seconds)
                                if i < len(valid_emails) - 1:  # Don't delay after last email
                                    time.sleep(random.randint(8, 20))

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
                                    st.info("🧪 Mode test utilisé - Pensez à désactiver le mode test pour l'envoi complet")

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
