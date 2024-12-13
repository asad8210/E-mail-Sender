import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pytesseract
from PIL import Image
import pdfplumber
import json
import docx

# Function to extract text from PDF
def extract_text_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text() or ''  # Handle None case
    return text

# Function to extract text from Image
def extract_text_from_image(image_file):
    image = Image.open(image_file)
    text = pytesseract.image_to_string(image)
    return text

# Function to extract text from DOCX file
def extract_text_from_docx(doc_file):
    doc = docx.Document(doc_file)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"  # Add newline for better formatting
    return text

# Function to extract Name and Email from JSON
def extract_data_from_json(json_file):
    with open(json_file, 'r') as f:
        data = json.load(f)
    return data.get('Name', ''), data.get('Email', '')

# Function to extract Name and Email from extracted text
def extract_name_and_email(text):
    name = ""
    email = ""
    lines = text.splitlines()
    for line in lines:
        if "Name" in line:
            name = line.split(":")[-1].strip()
        if "Email" in line:
            email = line.split(":")[-1].strip()
    return name, email

# Function to send personalized bulk emails
def send_bulk_emails(sender_email, sender_password, subject, body, recipients, attachment_files):
    success_count = 0
    fail_count = 0
    failed_emails = []
    
    try:
        # Connect to the SMTP server once
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)

            for name, email in recipients:
                if not name or not email:
                    fail_count += 1
                    failed_emails.append(f"Missing data for: {name if not name else email}")
                    continue

                personalized_body = body.replace("[Name]", name)
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = email
                msg['Subject'] = subject
                msg.attach(MIMEText(personalized_body, 'plain'))

                # Attach files
                for file in attachment_files:
                    attach_file(file, msg)

                # Send the email
                try:
                    server.sendmail(sender_email, email, msg.as_string())
                    success_count += 1
                except Exception as e:
                    fail_count += 1
                    failed_emails.append(f"Failed to send to {email}: {str(e)}")

    except Exception as e:
        fail_count += 1
        failed_emails.append(f"SMTP connection error: {str(e)}")

    return success_count, fail_count, failed_emails

# Function to attach files to the email
def attach_file(file, msg):
    file_name = file.name
    file_content = file.read()
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(file_content)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={file_name}')
    
    msg.attach(part)

# Streamlit App
def main():
    st.set_page_config(page_title="Bulk Email Sender", layout="centered")

    st.title("📧 Bulk Email Sender with File Extraction")

    st.markdown("""
        Upload PDF, Image, DOC, or JSON files to extract Name and Email. 
        These will be used to send personalized emails to each recipient.
    """)

    # Email Credentials
    st.subheader("Enter Your Email Credentials:")
    sender_email = st.text_input("Your Email Address:")
    sender_password = st.text_input("Your Email Password:", type="password")
    sender_name = st.text_input("Your Name:")

    # File Upload Section
    st.subheader("Upload PDF, Image, DOC, or JSON Files :")
    uploaded_files = st .file_uploader("Choose files", type=["pdf", "jpg", "png", "docx", "json"], accept_multiple_files=True)

    extracted_data = []

    if uploaded_files:
        for uploaded_file in uploaded_files:
            text = ""
            name = ""
            email = ""

            # Extract text based on file type
            if uploaded_file.name.endswith("pdf"):
                text = extract_text_from_pdf(uploaded_file)
            elif uploaded_file.name.endswith("jpg") or uploaded_file.name.endswith("png"):
                text = extract_text_from_image(uploaded_file)
            elif uploaded_file.name.endswith("docx"):
                text = extract_text_from_docx(uploaded_file)
            elif uploaded_file.name.endswith("json"):
                name, email = extract_data_from_json(uploaded_file)
                extracted_data.append((name, email))
                continue  # JSON already has name and email, no need to extract further
            
            if text:
                name, email = extract_name_and_email(text)
            
            # Store extracted data
            if name and email:
                extracted_data.append((name, email))

        # Show extracted data in list format
        if extracted_data:
            st.subheader("Extracted Data:")
            for name, email in extracted_data:
                st.write(f"Name: {name}, Email: {email}")
        else:
            st.write("No data extracted from the files.")

    # Email Content
    st.subheader("Compose Your Message:")
    subject = st.text_input("Email Subject:")
    body = st.text_area("Email Body (Use [Name] as a placeholder for each recipient's name):")

    # Upload Attachments
    st.subheader("Upload Attachments (Optional):")
    attachment_files = st.file_uploader("Choose attachments (CVs, JSON, PDFs, DOCs, Images)", 
                                        type=["pdf", "doc", "docx", "json", "jpg", "png"], 
                                        accept_multiple_files=True)

    # Send Button
    if st.button("Send Emails"):
        if sender_email and sender_password and sender_name and extracted_data and subject and body:
            success_count, fail_count, failed_emails = send_bulk_emails(sender_email, sender_password, subject, body, extracted_data, attachment_files)

            st.success(f"Emails sent successfully to {success_count} recipients!")

            if fail_count > 0:
                st.error(f"Failed to send {fail_count} emails.")
                st.write("Failed emails:")
                st.write(failed_emails)

        else:
            st.error("Please make sure all fields are filled out correctly before sending emails.")

if __name__ == "__main__":
    main()