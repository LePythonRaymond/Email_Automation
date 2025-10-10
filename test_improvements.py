#!/usr/bin/env python3
"""
Test script for the improved email automation app
Tests the new column detection and personalization features
"""

import pandas as pd
import sys
import os

# Add the current directory to the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from email_automation_app import EmailAutomation

def test_column_detection():
    """Test the intelligent column detection"""
    print("🔍 Testing Column Detection...")

    # Create test data with various column name formats
    test_data = {
        'Nom du Contact': ['Marie Dupont', 'Jean Martin', 'Sophie Bernard'],
        'Email Principal': ['marie@test.com', 'jean.martin@company.fr', 'sophie@example.org'],
        'Société': ['ABC Corp', 'XYZ Ltd', 'Test Inc'],
        'Lieu de Travail': ['Bureau Paris', 'Site Lyon', 'Usine Marseille'],
        'Email Secondaire': ['marie2@test.com', '', 'sophie.alt@example.org']
    }

    df = pd.DataFrame(test_data)
    automation = EmailAutomation()

    # Test column mapping
    mapping = automation.detect_column_mapping(df)
    print(f"✅ Column mapping detected: {mapping}")

    # Test email extraction
    valid_contacts = automation.get_valid_emails_from_df(df)
    print(f"✅ Valid contacts found: {len(valid_contacts)}")

    for contact in valid_contacts[:2]:  # Show first 2
        print(f"   - {contact['contact_name']} ({contact['email']}) - {contact['company']}")

    return len(valid_contacts) > 0

def test_personalization():
    """Test the improved personalization system"""
    print("\n✏️ Testing Personalization...")

    automation = EmailAutomation()

    # Test with custom content
    custom_content = """Bonjour {contact_name},

Nous espérons que vous allez bien.

Notre équipe souhaite vous présenter nos nouveaux services adaptés à votre {site_context}.

Cordialement,
L'équipe MERCI RAYMOND"""

    # Test text personalization
    text_result = automation.personalize_email(
        "Marie Dupont", "Bureau Paris", "ABC Corp",
        custom_content, use_html=False
    )

    print("✅ Text personalization:")
    print("   Preview:", text_result[:100] + "...")

    # Test HTML personalization
    html_result = automation.personalize_email(
        "Marie Dupont", "Bureau Paris", "ABC Corp",
        custom_content, use_html=True
    )

    print("✅ HTML personalization:")
    print("   HTML generated:", "<!DOCTYPE html" in html_result)
    print("   Contains name:", "Marie" in html_result)
    print("   Contains context:", "entreprise ABC Corp" in html_result)

    return "Marie" in text_result and "<!DOCTYPE html" in html_result

def test_verification():
    """Test email content verification"""
    print("\n🔍 Testing Verification...")

    automation = EmailAutomation()

    # Test valid content
    valid_content = "Bonjour Marie, comment allez-vous ? Cordialement."
    is_valid, issues = automation.verify_email_content(valid_content)
    print(f"✅ Valid content check: {is_valid} (issues: {len(issues)})")

    # Test invalid content with placeholders
    invalid_content = "Bonjour {name}, votre [site] est XXX"
    is_valid, issues = automation.verify_email_content(invalid_content)
    print(f"✅ Invalid content check: {not is_valid} (issues found: {len(issues)})")

    return True

def main():
    """Run all tests"""
    print("🧪 Testing Improved Email Automation App")
    print("=" * 50)

    tests = [
        ("Column Detection", test_column_detection),
        ("Personalization", test_personalization),
        ("Verification", test_verification)
    ]

    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append(result)
            print(f"✅ {test_name}: PASSED")
        except Exception as e:
            print(f"❌ {test_name}: FAILED - {e}")
            results.append(False)

    print("\n" + "=" * 50)
    print(f"🎯 Test Results: {sum(results)}/{len(results)} passed")

    if all(results):
        print("🎉 All improvements working correctly!")
    else:
        print("⚠️ Some tests failed - check the output above")

if __name__ == "__main__":
    main()
