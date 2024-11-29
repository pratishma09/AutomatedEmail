import os
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from dotenv import load_dotenv

load_dotenv()

# File paths
send_to = "data.csv"  # CSV file containing recipient data
template_path = os.path.join(os.getcwd(), "template", "template.html")  # Path to HTML template
attachments_folder = os.path.join(os.getcwd(), "Invitations")  # Folder containing image attachments

# Email subject
subject = "You're Invited to Graduation!"

def send_email():
    host = "smtp.gmail.com"
    port = 587

    # Sender credentials
    from_mail = os.getenv("APP_EMAIL")
    app_password = os.getenv("APP_PASSWORD")

    # Read HTML template
    with open(template_path, "r", encoding="utf-8") as file:
        email_content = file.read()

    # Initialize SMTP server
    with smtplib.SMTP(host, port, timeout=10) as smtp:
        smtp.starttls()
        print(f"Logging in as {from_mail}...")
        smtp.login(from_mail, app_password)
        print("Successfully logged in!")

        # Read recipient data from CSV
        df = pd.read_csv(send_to)
        total = len(df.index)
        print(f"Loaded {total} recipients from {send_to}.")

        # Loop through recipients and send emails
        for index, row in df.iterrows():
            # Replace placeholders with data from CSV
            personalized_content = email_content.replace("{name}", row["name"].title())

            # Create email message
            msg = MIMEMultipart()
            msg.attach(MIMEText(personalized_content, "html"))
            msg["Subject"] = subject
            msg["From"] = formataddr(("DWIT College", from_mail))
            msg["To"] = formataddr((row["name"], row["email"]))

            # Add attachment (image from Invitations folder)
            image_path = os.path.join(attachments_folder, row["image"])
            if os.path.exists(image_path):
                with open(image_path, "rb") as img_file:
                    img = MIMEImage(img_file.read(), name=os.path.basename(image_path))
                    msg.attach(img)
                print(f"Attached {row['image']} for {row['email']}.")
            else:
                print(f"[x] Image {row['image']} not found for {row['email']}.")

            # Send email
            try:
                smtp.sendmail(from_mail, row["email"], msg.as_string())
                print(f"Email sent to {row['email']} ({index + 1}/{total}).")
            except smtplib.SMTPRecipientsRefused as e:
                print(f"[x] Failed to send email to {row['email']}: {e}")

if __name__ == "__main__":
    send_email()
