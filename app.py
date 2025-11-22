import streamlit as st  # For building the web UI
import pandas as pd  # For handling Excel data
import smtplib  # For sending emails using SMTP
from email.mime.multipart import MIMEMultipart  # For creating multi-part emails
from email.mime.text import MIMEText  # For adding plain text / HTML parts
from email.mime.image import MIMEImage  # For embedding inline images
import ssl  # For creating secure connection

# --- App Title ---
st.title("📧 Bulk Email Sender using Excel (Personalized with Links & Inline Image)")

# --- Step 1: Upload Excel File ---
st.header("Step 1: Upload Recipient List")
uploaded_file = st.file_uploader(
    "Upload Excel File with Name, Email (and optional Company) columns",
    type=["xlsx"]
)

# Initialize recipients list
recipients = []
if uploaded_file:
    # Read Excel file into DataFrame
    df = pd.read_excel(uploaded_file)
    required_cols = {"Name", "Email"}  # Required columns

    # Validate required columns
    if not required_cols.issubset(df.columns):
        st.error("❌ Excel file must have at least 'Name' and 'Email' columns.")
    else:
        # Add empty "Company" column if not present
        if "Company" not in df.columns:
            df["Company"] = ""

        # Convert dataframe rows to dictionary list (skip rows with missing email)
        recipients = df[["Name", "Email", "Company"]].dropna(subset=["Email"]).to_dict("records")

        # Show confirmation and display the dataframe
        st.success(f"✅ Loaded {len(recipients)} recipient(s).")
        st.write(df)

# --- Step 2: Compose Email (only if recipients available) ---
if recipients:
    st.header("Step 2: Compose and Send Emails")

    # Input for sender email and password
    sender_email = st.text_input("Your Email Address")
    password = st.text_input("Your Email Password / App Password", type="password")

    # Helpful link for creating Google App Password
    st.markdown(
        """
        🔑 Need help generating a password? 
        <a href="https://myaccount.google.com/apppasswords?utm_source=chatgpt.com" target="_blank">
            👉 Click here to create/get your Google App Password
        </a>
        """,
        unsafe_allow_html=True
    )

    # Subject and message template
    subject = st.text_input("Email Subject")
    st.markdown("✍️ Use `{name}` and `{company}` for personalization. "
                "You can also insert HTML like `<a href='https://example.com'>Click Here</a>`.")
    message_template = st.text_area("Email Message (HTML supported)")

    # Optional inline image
    uploaded_image = st.file_uploader(
        "Optional: Upload Image (will appear inside email)",
        type=["png", "jpg", "jpeg"]
    )

    # --- Step 3: Send Emails ---
    # Disable button if any required field is empty
    send_disabled = not (
        sender_email.strip() and
        password.strip() and
        subject.strip() and
        message_template.strip()
    )

    if st.button("🚀 Send Emails", disabled=send_disabled):
        try:
            # Create secure SMTP connection (Gmail SMTP server on port 587)
            context = ssl.create_default_context()
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls(context=context)  # Upgrade to TLS encryption
                server.login(sender_email, password)  # Login with user credentials

                # Progress bar and status text placeholders
                progress = st.progress(0)
                status_text = st.empty()

                # --- Loop through recipients ---
                for i, rec in enumerate(recipients):
                    recipient_name = rec.get("Name", "")
                    recipient_email = rec.get("Email", "")
                    recipient_company = rec.get("Company", "")

                    # Personalize message with placeholders {name} and {company}
                    try:
                        personalized_message = message_template.format(
                            name=recipient_name,
                            company=recipient_company
                        )
                    except KeyError:
                        st.error("⚠️ Error: Only use {name} and {company} placeholders.")
                        break

                    # --- Build email message ---
                    msg = MIMEMultipart("related")  # Allows HTML + images
                    msg["From"] = sender_email
                    msg["To"] = recipient_email
                    msg["Subject"] = subject

                    # Create alternative part (plain + html)
                    msg_alternative = MIMEMultipart("alternative")
                    msg.attach(msg_alternative)

                    # Plain text fallback (for email clients not supporting HTML)
                    plain_message = f"{personalized_message}"
                    msg_alternative.attach(MIMEText(plain_message, "plain"))

                    # HTML message (with optional inline image)
                    html_message = f"""
                    <html>
                      <body>
                        {personalized_message}
                        {"<br><img src='cid:banner' style='width:600px;max-width:100%;height:auto;border-radius:10px;'>" if uploaded_image else ""}
                      </body>
                    </html>
                    """
                    msg_alternative.attach(MIMEText(html_message, "html"))

                    # Attach inline image (if provided)
                    if uploaded_image is not None:
                        uploaded_image.seek(0)  # Reset pointer to beginning
                        img = MIMEImage(uploaded_image.read())
                        img.add_header("Content-ID", "<banner>")  # cid reference for inline HTML
                        img.add_header("Content-Disposition", "inline", filename=uploaded_image.name)
                        msg.attach(img)

                    # --- Send email ---
                    try:
                        server.sendmail(sender_email, recipient_email, msg.as_string())
                        status_text.text(f"📨 Sent to {recipient_name} ({recipient_email})")
                    except Exception as send_err:
                        status_text.text(f"❌ Failed to {recipient_email}: {send_err}")

                    # Update progress bar
                    progress.progress((i + 1) / len(recipients))

            # Final success message
            st.success(f"🎉 Personalized emails sent successfully to {len(recipients)} recipients!")

        except Exception as e:
            # Catch all errors (login issues, SMTP errors, etc.)
            st.error(f"⚠️ Error: {e}")
