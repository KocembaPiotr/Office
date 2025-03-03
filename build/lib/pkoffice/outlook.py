import win32com.client as win32
import time

ATTACHMENTS_TMP_PATH = r'C:\Temp\\'


def send_mail(address_to: str, address_cc: str, mail_subject: str, mail_body: str,
              attachments_list: list = None, address_from: str = None) -> None:
    """
    Function to send mail using Outlook without opening it.
    :param address_to: recipient address list 'John@John.com;Ann@Ann.com'
    :param address_cc: recipient in copy address list 'John@John.com;Ann@Ann.com'
    :param mail_subject: subject of mail
    :param mail_body: body of mail using HTML
    :param attachments_list: list of paths of attachments ['path1','path2']
    :param address_from: indication who will send mail
    :return: None
    """
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    if address_from is not None:
        mail.SentOnBehalfOfName = address_from
    mail.To = address_to
    mail.Cc = address_cc
    mail.Subject = mail_subject
    mail.HTMLBody = mail_body
    if attachments_list is not None:
        for attachment in attachments_list:
            mail.Attachments.Add(attachment)
    time.sleep(3)
    mail.Send()
