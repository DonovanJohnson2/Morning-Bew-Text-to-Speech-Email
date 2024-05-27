import os
import pyttsx3
import win32com.client
import requests
from bs4 import BeautifulSoup
import smtplib
from email.message import EmailMessage

# Function to read email from Outlook
def read_email():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 is the inbox folder ID
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    message = messages.GetFirst()
    subject = message.Subject
    body = message.Body
    return subject, body

# Function to download Morning Brew newsletter
def download_newsletter():
    response = requests.get("https://www.morningbrew.com/daily/issues/latest")
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        # Extract text from the webpage as needed
        text = soup.get_text()  # Example: Get all text from the webpage
        return text
    else:
        print("Failed to fetch website content.")
        return None

# Function to convert text to speech
def text_to_speech(text):
    engine = pyttsx3.init()
    engine.save_to_file(text, 'morning_brew.mp3')
    engine.runAndWait()

# Function to send email with audio attachment
def send_email(subject, body, attachment_path):
    sender_email = "donovanjohnson53@gmail.com"
    receiver_email = "donovanjohnson53@gmail.com"
    password = "srjq rqmj rdkc dcth"

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype='audio', subtype='mp3', filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, password) 
        smtp.send_message(msg)

def main():
    # Read email from Outlook
    subject, body = read_email()

    # Download Morning Brew newsletter
    newsletter_content = download_newsletter()

    if newsletter_content:
        # Convert newsletter text to speech
        text_to_speech(newsletter_content)

        # Send email with the audio attachment
        send_email(subject, body, 'morning_brew.mp3')
    else:
        print("Failed to download Morning Brew newsletter.")

if __name__ == "__main__":
    main()