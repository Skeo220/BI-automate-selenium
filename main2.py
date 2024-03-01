from seleniumbase import BaseCase
import smtplib
import os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText 
from email import encoders

class TableauDownload(BaseCase):
    def download_tableau_pdf(self):
        # Open Tableau visualizations
        self.open("https://public.tableau.com/app/discover/viz-of-the-day")
        self.wait_for_ready_state_complete()

        # Click on the first element
        self.sleep(2)
        self.click('img[alt="Workbook thumbnail"]')

        # Click on download
        self.wait_for_ready_state_complete()
        self.sleep(4)
        self.click('button[data-tip="Download"]')

        # Switch to the new window
        self.sleep(1)
        self.switch_to_frame('iframe[title="Data Visualization"]')

        # Download PDF
        self.click('button[data-tb-test-id="DownloadPdf-Button"]')
        self.click('button[data-tb-test-id="export-pdf-export-Button"]')
        self.sleep(5)

class get_pbi_link(BaseCase):
    def pbi_link(self):
        self.open("https://community.fabric.microsoft.com/t5/Data-Stories-Gallery/bd-p/DataStoriesGallery")
        self.wait_for_ready_state_complete()

        # Wait and click on the first element
        self.wait_for_element('img[alt="View Video"]')
        self.click('img[alt="View Video"]')

        # Wait for the page to load
        self.wait_for_ready_state_complete()

        self.switch_to_frame('iframe[src*="app.fabric.microsoft.com/view"]')

        # Now perform actions inside the iframe
        # For example, clicking the share button
        self.wait_for_element('button[aria-label="Share"]')
        self.click('button[aria-label="Share"]')

        share_url = self.get_value('input.sharingUrlBox')
        return share_url
        

def send_email(share_url):
    load_dotenv()

    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.starttls()

    try:
        smtpObj.login(os.getenv('SENDER'), os.getenv('SENDER_PASSWORD'))
        print('Successfully logged into Gmail')
        
        folder_path = 'downloaded_files'
        num_files = len([name for name in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, name))])
        print('File count successful')

        sender_email = os.getenv('SENDER')
        receiver_email = os.getenv('RECIPIENT')
        subject = 'Hello World'

        # HTML body with hyperlink
        html_body = f"""
        <html>
            <body>
                <p>Hello,</p>
                <p>The number of files in the {folder_path} folder is: {num_files}.</p>
                <p>Here is the Power BI report link: <a href="{share_url}">View Report</a></p>
            </body>
        </html>
        """

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject

        # Attach HTML body to the email
        message.attach(MIMEText(html_body, 'html'))

        # Attach files from the downloaded_files folder
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                attachment = open(file_path, 'rb')

                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={filename}')

                message.attach(part)

                attachment.close()

        text = message.as_string()
        smtpObj.sendmail(sender_email, receiver_email, text)
        print('Email sent successfully')

    except smtplib.SMTPAuthenticationError:
        print('Error: Authentication failed')
    except Exception as e:
        print(f'An error occurred: {e}')
    finally:
        smtpObj.quit()


pass

class MainTest(TableauDownload, get_pbi_link):
    def test_tableau_pdf_download_and_email(self):
        # Download the Tableau PDF
        self.download_tableau_pdf()
        shared_url = self.pbi_link()

        # Send an email
        send_email(shared_url)

if __name__ == "__main__":
    MainTest.main(__name__, __file__)
