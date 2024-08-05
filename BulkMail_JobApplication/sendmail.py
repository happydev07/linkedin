import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import logging
from docx import Document
from time import sleep
import random
import winsound  # Import winsound for playing system sounds

logging.basicConfig(level=logging.INFO)

def play_sound():
    # Play Windows default sound
    winsound.MessageBeep()

def docx_to_html(file_path):
    doc = Document(file_path)
    html = ["<html><body>"]
    
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p_html = "<p style='margin: 0;'>"
            for run in element.xpath('.//w:r'):
                text = run.text.replace('\t', '&nbsp;' * 4)
                style = ''
                if run.rPr is not None:  # Explicit check for None
                    if getattr(run.rPr, 'b', None) is not None:
                        style += 'font-weight: bold;'
                    if getattr(run.rPr, 'i', None) is not None:
                        style += 'font-style: italic;'
                    if getattr(run.rPr, 'u', None) is not None:
                        style += 'text-decoration: underline;'
                if style:
                    p_html += f"<span style='{style}'>{text}</span>"
                else:
                    p_html += text
            p_html += "</p>"
            html.append(p_html)
        elif element.tag.endswith('tbl'):
            html.append("<table border='1' style='border-collapse: collapse;'>")
            for row in element.xpath('.//w:tr'):
                html.append("<tr>")
                for cell in row.xpath('.//w:tc'):
                    cell_html = "<td style='padding: 5px;'>"
                    for paragraph in cell.xpath('.//w:p'):
                        for run in paragraph.xpath('.//w:r'):
                            text = run.text.replace('\t', '&nbsp;' * 4)
                            style = ''
                            if run.rPr is not None:  # Explicit check for None
                                if getattr(run.rPr, 'b', None) is not None:
                                    style += 'font-weight: bold;'
                                if getattr(run.rPr, 'i', None) is not None:
                                    style += 'font-style: italic;'
                                if getattr(run.rPr, 'u', None) is not None:
                                    style += 'text-decoration: underline;'
                            if style:
                                cell_html += f"<span style='{style}'>{text}</span>"
                            else:
                                cell_html += text
                    cell_html += "</td>"
                    html.append(cell_html)
                html.append("</tr>")
            html.append("</table>")
    
    html.append("</body></html>")
    return ''.join(html)

def send_email(to_email, subject, body, attachment_path, counter, server, username):
    msg = MIMEMultipart('alternative')
    msg['From'] = username
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name=os.path.basename(attachment_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        msg.attach(part)

    try:
        server.sendmail(username, to_email, msg.as_string())
        logging.info(f"Email {counter} sent to: {to_email}")
        print("Mail sent successfully!")  # Display message after sending mail
    except Exception as e:
        logging.error(f"Error sending email to {to_email}: {e}")

def main():
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    username = "talkdev@gmail.com"
    password = "ttld ammg bvjk puwa"

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    try:
        server.login(username, password)
        logging.info("SMTP login successful.")
    except Exception as e:
        logging.error(f"Error during SMTP login: {e}")
        return

    email_addresses_path = r"C:\Users\devra\Desktop\BulkMail_JobApplication\email.txt"
    subject_path = r"C:\Users\devra\Desktop\BulkMail_JobApplication\subject.txt"
    body_path = r"C:\Users\devra\Desktop\BulkMail_JobApplication\cover.docx"
    attachment_path = r"C:\Users\devra\Desktop\BulkMail_JobApplication\AwsDevopsDeoraj.docx"

    with open(email_addresses_path, 'r', encoding='utf-8') as file:
        email_addresses = [line.strip() for line in file]

    with open(subject_path, 'r') as file:
        subject = file.read().strip()

    email_body = docx_to_html(body_path)

    counter = 0
    for to_email in email_addresses:
        if counter >= 300:  # Stop after sending 900 emails
            break
        counter += 1
        sleep(random.randint(1, 30))  # Sleep for a random time between 1 to 60 seconds
        send_email(to_email, subject, email_body, attachment_path, counter, server, username)

        # After every 20 emails, quit and reconnect to the SMTP server
        if counter % 20 == 0:
            server.quit()
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            try:
                server.login(username, password)
                logging.info("SMTP re-login successful.")
            except Exception as e:
                logging.error(f"Error during SMTP re-login: {e}")
                return

    server.quit()

if __name__ == "__main__":
    main()
