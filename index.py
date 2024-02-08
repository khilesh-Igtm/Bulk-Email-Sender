import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Function to read data from Excel file
def read_excel(filename):
    df = pd.read_excel(filename)
    return df

# Function to send email with attachment
def send_email(sender_email, sender_password, recipient_email, subject, body, resume_folder, resume_filename):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))  # Using 'plain' as it's just text

    # Construct path to resume
    resume_path = os.path.join(resume_folder, resume_filename)

    # Attach resume if path is not empty
    if os.path.exists(resume_path):
        with open(resume_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(resume_path)}')
        msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, sender_password)
    text = msg.as_string()
    server.sendmail(sender_email, recipient_email, text)
    server.quit()

# Main function
def main():
    # Replace these values with your own
    sender_email = 'khileshsingh5678@gmail.com'
    sender_password = 'ghux rfyv xxzf eqvh'
    
    # Define email subject
    subject = 'Application for Software Developer Role'

    # Replace this with the path to your Excel file
    excel_file = 'myemails.xlsx'

    # Read data from Excel file
    df = read_excel(excel_file)

    # Define email body content
    body = """Dear Hiring Manager,

I hope this message finds you well. I am reaching out to express my interest in the Software Developer position at your organization.
With 5 months of work experience at ITNow Inc as a Servicenow Developer and proficiency in a range of programming languages including Python, Java, JavaScript,
and PHP.
Attached is my resume for your consideration. I am excited about the opportunity to discuss how my background aligns with the needs of your team.

Thank you for your time and consideration.

Best Regards,
Khilesh Singh
Resume: https://drive.google.com/file/d/1yGQ-IzO6ngkhMzteKlfu9zytwFKFfZoH/view?usp=sharing
 """

    # Folder containing resumes (current directory)
    resume_folder = os.getcwd()

    # Iterate through each row and send email
    for index, row in df.iterrows():
        recipient_email = row['Email']
        resume_filename = str(row['Resume'])  # Convert to string

        send_email(sender_email, sender_password, recipient_email, subject, body, resume_folder, resume_filename)
        print(f"Email sent to {recipient_email}")

if __name__ == "__main__":
    main()
