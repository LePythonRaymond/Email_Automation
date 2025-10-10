# 🌱 MERCI RAYMOND - Email Automation App

A powerful email automation tool with dynamic Excel column detection and customizable email templates.

## Features

- **📊 Dynamic Excel Column Detection**: Works with any Excel format that has at least an email column
- **🎨 Customizable Headers & Footers**: Personalize email greetings and signatures
- **📧 Dynamic Placeholders**: Use any Excel column as placeholders in your emails
- **🚀 Professional Email Templates**: Gmail-style HTML formatting
- **🛡️ Anti-spam Features**: Built-in sending delays and test modes

## How to Use

1. **Upload Excel File**: Any format with at least an email column
2. **Customize Email**: Design your email content with dynamic placeholders
3. **Configure Gmail**: Set up your Gmail credentials
4. **Send Emails**: Send personalized emails to all contacts

## Supported Excel Formats

- Any column names (detected automatically)
- Multiple email columns
- Any number of additional data columns
- Works with French and English column names

## Placeholders

Use any column from your Excel as placeholders:
- `{Name}`, `{Company}`, `{Site}`, `{Location}`, etc.
- Special `{contact_name}` for first name
- All placeholders are optional (graceful degradation)

## Deployment

This app is deployed on Streamlit Community Cloud and ready to use!
