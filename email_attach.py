__author__ = 'anatolykonkin'

class FetchEmail():

    def __init__(self, user, pwd):
        # connecting to the gmail imap server
        self.user = user
        self.pwd = pwd

    def upload(self, sender, n_dir):
        # download all attachments from the last sender's email
        import imaplib, os
        import email
        # example from https://gist.github.com/baali/2633554#file-dlattachments-py
        # example from http://stackoverflow.com/questions/7596789/downloading-mms-emails-sent-to-gmail-using-python

        self.m = imaplib.IMAP4_SSL("imap.gmail.com")
        self.m.login(self.user, self.pwd)
        r, data = self.m.select('INBOX')

        resp, items = self.m.search(None, 'from','"%s"' % (sender,)) # you could filter using the IMAP rules here (check http://www.example-code.com/csharp/imap-search-critera.asp)
        items = items[0].split() # getting the mails id

        print(items)

        #last = max(items)
        last = items[len(items)-1]
        print(last)
        resp, data = self.m.fetch(last, "(RFC822)")# fetching the mail, "`(RFC822)`" means "get the whole stuff", but you can ask for headers only, etc
        email_body = data[0][1] # getting the mail content

        mail = email.message_from_string(email_body) # parsing the mail content to get a mail object

        #Check if any attachments at all
        if mail.get_content_maintype() != 'multipart':
            print('No attachments')
        else:
            pass

        counter = 0
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            # is this part an attachment ?
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()

            # if there is no filename, we create one with a counter to avoid duplicates
            # if not filename:
            #     filename = 'part-%03d%s' % (counter, 'bin')
            counter += 1

            detach_dir = n_dir + '/'
            if not os.path.exists(detach_dir):
                os.makedirs(detach_dir)

            fp = open(detach_dir + '/' + str(counter) + '.xlsx', 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()

    def send_email(self, send_to, dir):
        gmail_user = self.user
        gmail_pwd = self.pwd

        import smtplib, os
        from email.MIMEMultipart import MIMEMultipart
        from email.MIMEBase import MIMEBase
        from email.MIMEText import MIMEText
        from email.Utils import COMMASPACE, formatdate
        from email import Encoders

        FROM = self.user
        TO = [send_to] #must be a list

        msg = MIMEMultipart()
        msg['Subject'] = 'items for 1C'
        msg['From'] = FROM
        msg['To'] = ', '.join(TO)

        msg.attach(MIMEText('See attachments'))

        # attach files from dir
        fileList = [os.path.normcase(f) for f in os.listdir(dir)]
        for fn in fileList:
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(dir + '/' + fn,"rb").read())
            Encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(fn))
            msg.attach(part)

        # send email
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.login(gmail_user, gmail_pwd)
        server.sendmail(FROM, TO, msg.as_string())
        server.close()