import os
import smtplib
from email.header import Header
from email.utils import formataddr
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


class AutoEmail:
    email_address = os.environ.get('EMAIL_USER')
    email_password = os.environ.get('EMAIL_PASS')
    # author = u'이일희 매니저 <lih@glovis.net>'

    def __init__(self, receiver=None, cc=None, subject=None, content=None, attachments=None,
                  img=None, reply_to='lih@glovis.net'):
        self.receiver = receiver
        self.cc = cc
        self.subject = subject
        self.content = content
        self.attachments = attachments
        self.img = img
        self.reply_to = reply_to
        self.msg = None
        self.html = """<html><body><p>{content}</p>""".format(content=self.content)

    def initiate(self):
        self.msg = MIMEMultipart()
        self.msg['From'] = 'lih1500252@naver.com'
        self.msg['To'] = self.receiver
        self.msg['Cc'] = self.cc
        self.msg['Reply-To'] = self.reply_to
        self.msg['Subject'] = self.subject
        self.print_init()

    def print_init(self):
        print('\n<<Email Information>>')
        print('From :', AutoEmail.email_address)
        print('To :', self.receiver)
        print ('Cc :', self.cc)
        print('Reply-To :', self.reply_to)
        print('Subject :', self.subject)
        print('Content : \n<Quote>\n', self.content, '\n<Quote>\n')

    def attach(self):
        if self.attachments is None:
            pass
        else:
            if type(self.attachments) is str:
                self.attachments = [self.attachments]
            for n, file in enumerate(self.attachments):
                print('Attachment_{} : {}'.format(n+1, file))
                with open (file, 'rb') as f:
                    file_data = MIMEApplication(f.read(), Name=os.path.basename(file))
                    self.msg.attach(file_data)

    def build_html(self):
        if self.img is None:
            self.html += """</body></html>"""
        else:
            if type(self.img) is str:
                self.img = [self.img]
            for n, img_file in enumerate(self.img):
                print ('Image_{} : {}'.format (n + 1, img_file))
                self.html += """<img src="cid:{}" alt='image' style="display: block; visibility: visible;">""".format(img_file)
                """ https://stackoverflow.com/questions/48272100/embed-an-image-in-html-for-automatic-outlook365-email-send """
                with open(img_file, 'rb') as file:
                    img = MIMEImage(file.read())
                    img.add_header('Content-ID', '<{}>'.format(img_file))
                    self.msg.attach(img)
            self.html += """</body></html>"""
            self.msg.attach(MIMEText(self.html, 'html'))

    def send(self):
        self.initiate()
        self.attach()
        self.build_html()
        # ans = input('\nAre you sure to proceed?[Y/N, Default: N] : ')
        # if ans.lower() == 'y':
        with smtplib.SMTP_SSL ('smtp.naver.com', 465) as smtp:
            smtp.login (AutoEmail.email_address, AutoEmail.email_password)
            smtp.send_message (self.msg)
            print ('Mail sent successfully!\n')
