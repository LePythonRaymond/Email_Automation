#!/usr/bin/env python3
"""
Test script for MERCI RAYMOND Email Automation
"""

import pandas as pd
import re
from email_automation_app import EmailAutomation

def test_email_automation():
    print("🧪 Test de l'automation email MERCI RAYMOND")
    print("=" * 50)

    # Initialize automation
    automation = EmailAutomation()

    # Test data
    test_contacts = [
        {
            'contact': 'Emmanuel SCHREDER',
            'site': '190 Avenue Victor Hugo',
            'company': 'Hestia-im',
            'email': 'taddeo.carpinelli@merciraymond.fr'
        },
        {
            'contact': 'Daphnée TYTELMAN',
            'site': 'Adyen Exterior',
            'company': 'Adyen',
            'email': 'taddeo.carpinelli@merciraymond.fr'
        }
    ]

    print(f"📊 Test avec {len(test_contacts)} contacts")
    print()

    for i, contact in enumerate(test_contacts, 1):
        print(f"Test {i}: {contact['contact']} - {contact['site']}")
        print("-" * 40)

        # Test simple personalization
        personalized = automation.personalize_email_simple(
            contact['contact'],
            contact['site'],
            contact['company']
        )

        print("📧 Email personnalisé:")
        print(personalized[:200] + "..." if len(personalized) > 200 else personalized)
        print()

        # Test verification
        is_valid, issues = automation.verify_email_content(personalized)

        if is_valid:
            print("✅ Validation: OK")
        else:
            print("❌ Validation: Problèmes détectés")
            for issue in issues:
                print(f"  - {issue}")

        print("=" * 50)
        print()

def test_excel_loading():
    print("📁 Test de chargement Excel")
    print("=" * 30)

    try:
        df = pd.read_excel("/Users/taddeocarpinelli/Desktop/MERCI RAYMOND/Sytème mail/Liste entretiens.xlsx")

        print(f"✅ Fichier chargé: {len(df)} lignes")

        # Count emails
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

        email_count = 0
        for idx, row in df.iterrows():
            if pd.notna(row.get('Contact client 1', '')) and re.search(email_pattern, str(row['Contact client 1'])):
                email_count += 1
            elif pd.notna(row.get('Contact client 2 ', '')) and re.search(email_pattern, str(row['Contact client 2 '])):
                email_count += 1

        print(f"📧 Emails valides trouvés: {email_count}")
        print(f"📊 Taux de couverture: {email_count/len(df)*100:.1f}%")

        return True

    except Exception as e:
        print(f"❌ Erreur: {e}")
        return False

if __name__ == "__main__":
    print("🌱 MERCI RAYMOND - Test de l'application")
    print("=" * 60)
    print()

    # Test Excel loading
    excel_ok = test_excel_loading()
    print()

    if excel_ok:
        # Test email automation
        test_email_automation()

    print("🎉 Tests terminés!")
    print()
    print("Pour lancer l'application:")
    print("streamlit run email_automation_app.py")
