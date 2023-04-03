import os
import win32com.client as win32
import datetime as dt
from pretty_html_table import build_table
import glob


def file_download(file_path, index_number):
    cwd = os.chdir(file_path)
    cwd = os.getcwd()
    latest_file = sorted(glob.iglob(cwd + '\*'),
                         key=os.path.getmtime, reverse=True)
    READ = latest_file[index_number]
    return READ


class Email:
    def __init__(self, inbox_folder, subject, output_dir):
        self.inbox_folder = inbox_folder
        self.subject = subject
        self.output_dir = output_dir

    def file_download(self, index_number):
        os.chdir(self.output_dir)
        cwd = os.getcwd()
        latest_file = sorted(glob.iglob(cwd + '\*'),
                             key=os.path.getmtime, reverse=True)
        READ = latest_file[index_number]
        return READ

    def __str__(self):
        return f'Inbox folder: {self.inbox_folder}, Subject: {self.subject}, Output directory: {self.output_dir}'

    def clean_up(self):
        latest_file = sorted(glob.iglob(self.output_dir + '\*'),
                             key=os.path.getmtime, reverse=True)
        [os.remove(i) for i in latest_file]

    def personal_email_dl(self):
        outlook = win32.Dispatch('outlook.application')
        mapi = outlook.GetNamespace('MAPI')
        inbox = mapi.GetDefaultFolder(6).Folders[self.inbox_folder]
        messages = inbox.Items
        messages1 = messages.Restrict(f"[Subject] = {self.subject}")
        date_same_day = dt.datetime.now() - dt.timedelta(hours=24)
        date_same_day = messages1.Restrict(
            "[ReceivedTime] >= '" + date_same_day.strftime('%d/%m/%Y %H:%M %p')+"'")
        date_same_day.Sort("[ReceivedTime]", Descending=True)
        try:
            for message in list(date_same_day)[0:1]:
                try:
                    s = message.Sender
                    for attachment in message.Attachments:
                        attachment.SaveAsFile(os.path.join(
                            self.output_dir, attachment.FileName))
                        print(
                            f"attachment {attachment.FileName} from {s} saved")
                except Exception as e:
                    print("error when saving the attachment:" + str(e))
        except Exception as e:
            print("error when processing email messages:" + str(e))


class Email_send:

    """

    Template for sending email messages using VBA

    Save the body in a variable if required, needs to be written with HTML elements, use triple quotes for long messages

    recipient(str): email of recipient (can be a list of STR's)
    subject(str): subject of email
    attachment_path(raw-str): path of the attachment
    """

    def __init__(self, recipient, subject, attachment_path):
        self.recipient = recipient
        self.subject = subject
        self.attachment_path = attachment_path

    def __str__(self):
        return 'Recips: {}, Subject: {}, Attach: {}'.format(self.recipient, self.subject, self.attachment_path)

    def create_html_table(self, DF, tab_colour, list_col_width, font_size, text_align, e_bg_colour, padding):
        styled_table_body = build_table(DF, tab_colour, font_size=font_size,

                                        text_align=text_align, index=True, even_bg_color=e_bg_colour, padding=padding,

                                        width_dict=list_col_width)
        return styled_table_body

    def create_html_body(self, DF):
        get_email = self.recipient.split('.')
        name = get_email[0]
        body = f"""
        <html>
        <head>
        </head>
        <body>

        <span style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: black;">
            <p>Hello,<br />
            <br />
            Here is a quick table to summarize the attached.
            {DF}
            <br />
            Let me know if you have any questions. <br />
            <br/ >
            Thanks,<br />
            James <br />
            <br />
            </p>
        </span>
        </body>
        </html>

        """
        return body
    
    def text_body(self, text_body):
        body = f"""
        <html>
        <head>
        </head>
        <body>
        <span style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: black;">        
        <p>Hello,<br />
        <br />      
        {text_body}
        Thanks,<br />
        James <br />
        <br />
        </p>
        </span>
        </body>
        </html>          
        """

    def email_send_generic(self, body):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = self.recipient
        mail.Subject = self.subject
        mail.HTMLBody = body
        mail.Attachments.Add(self.attachment_path)
        mail.Send()
