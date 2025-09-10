## PACKAGES ##
import os
import pandas as pd
import win32com.client as win32

## PATHS ##
MAIN = os.getcwd()
CONFIG = os.path.join(MAIN, "config")
ATTACH = os.path.join(MAIN, "attachments")

###############
## Functions ##
###############
def get_config(file_path):
    data = pd.read_excel(file_path)
    res = data.to_dict(orient = 'records')
    return res

def read_text(file_path):
    with open(file_path, "r") as f:
        content = f.read()
    return content

def fill_text(base_text, fill_dict):
    mod_text = base_text
    for k, v in fill_dict.items():
        mod_text = mod_text.replace(f'[{k}]', v)
    return mod_text

def create_outlook_draft(to, cc, subject, body, attach = None):
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0 represents MailItem

        ## Meta parameters
        mail.To = to
        mail.Cc = cc
        mail.Subject = subject
        ## Editing body shenanigans
        compat_body = body.replace('\n', '<br>')
        mail.Display()
        mail.HTMLBody = mail.HTMLBody.replace('&nbsp;', compat_body, 1)
        ## Attachments
        if attach:
            for attachment_path in attach:
                mail.Attachments.Add(attachment_path)

        mail.Save()
        mail.Close(0)
        print('Draft email created successfully in Outlook.')

    except Exception as e:
        print(f'An error occurred: {e}')

if __name__ == "__main__":
    config_path = os.path.join(CONFIG, 'Config.xlsx')
    entries = get_config(config_path)

    email_body_path = os.path.join(CONFIG, 'EmailBody.txt')
    email_body = read_text(email_body_path)

    for entry in entries:
        body = fill_text(email_body, entry)
        to = entry["TO"]
        cc = entry["CC"]
        subject = entry["SUBJECT"]
        attachments = [os.path.join(ATTACH, i) for i in entry["ATTACHMENTS"].split(" ; ")]
        create_outlook_draft(to, cc, subject, body, attachments)




