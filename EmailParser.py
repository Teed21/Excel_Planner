# Written and Developed by: Tyler Wright
# Date started: 04/25/2019
# Date when workable: 4/26/2019
# Last Updated: 04/26/2019

import win32com.client
import re


class EmailParser:

    def __init__(self):
        # Opening an outlook object to get data from outlook.
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Issuing "inbox" to default folder 6, outlook's default number for the inbox folder.
        self.inbox = self.outlook.GetDefaultFolder("6")

    def find_email(self, inbox_folder, email_filter):
        # This gets all emails in the specific inbox/folder.
        # emails = self.inbox.Items
        inbox_subfolder = self.inbox.Folders(inbox_folder)
        emails = inbox_subfolder.Items

        # Gets most recent email in inbox/subfolder.
        # message = emails.GetLast()
        # Gets entire body of the email message.
        # body_content = message.body
        # print(body_content)

        # List to hold all relevant emails when passed through filtering.
        s_requests = []

        # Filtering emails to find S Requests.
        for email in emails:
            if email_filter in email.Subject:
                # print(email.Subject)
                s_requests.append(email)

        # This helps set a name for the final excel file, if the user wants to name the excel file differently.
        if not s_requests:
            return False

        return s_requests

    def find_cps(self, emails):
        # Finds if more than one email exists. Takes either the first if one email, or the last if multiple.
        # If more than one is found, the list was populated from oldest to newest, so the newest email is the last one.
        if len(emails) > 1:
            # print("Multiple emails found.")
            # Getting last email in list.
            relevant_email = emails[-1]
        else:
            # print("One email found.")
            # Getting first email in list.
            relevant_email = emails[0]

        email_contents = relevant_email.body

        # This code is snatched from Stackoverflow:
        # (https://stackoverflow.com/questions/4666973/how-to-extract-the-substring-between-two-markers).
        # It uses a Regular Expression to find the text values between two words inside one string.
        cp_names = re.search('Location:(.+?)\nSubject', email_contents)
        if cp_names:
            suggested_cps = cp_names.group(1)
        else:
            suggested_cps = "Problem finding CPs in EmailParser/find_cps"

        return suggested_cps


#emailobj = EmailParser()

#emails = emailobj.find_email("Requests", "Request 2019S038")

#emailobj.find_cps(emails)
