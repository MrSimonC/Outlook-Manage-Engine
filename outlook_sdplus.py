import os
import pywintypes
import re
import sys
import win32com.client
from custom_modules.sdplus_api_rest import API
# from custom_modules.slack import API as SlackAPI
__version__ = '0.71'
# 0.3 - Updated the signature remover and inserted cssc@ line 27/Jan/16
# 0.4 - Added ActiveInspector.Close(0) to save changes else they're abandoned if you don't call .Display()
# 0.5 - Changed Outlook inbox parser to for no in range(inbox.items.count-1, -1, -1)
# 0.6 - Introduced searching for HD number and setting sdplus' supplier ref field
# 0.6 - Added functionality which looks for no assignee, if true, sends slack notification
# 0.7 - Moved API key to env variable
# 0.71 - Removed Slack hooks


class OutlookSDPlus:
    """
    Process outlook folder for manage engine helpdesk number, confirm number is valid, forward email into Manage Engine
    appending the local reference number with ##s
    """
    def __init__(self):
        print('Remedy email processor v' + __version__)
        self.outlook = win32com.client.Dispatch('Outlook.Application')
        self.inbox = None
        self.sdplus_api = None
        self.sdplus_api_key = os.environ['SDPLUS_ADMIN']
        self.sdplus_api_url = 'http://sdplus/sdpapi/request/'
        self.sdplus_clean = r'(?:##)(\d{6})(?:##)'  # clean sdplus number is group 1 as group(0)=entire match
        self.sdplus_csc = r'(?:NBNT|NBNTSD)(\d{6})'  # clean sdplus number is group 1
        self.hd_ref = r'(?:HD0*)(\d{7}\b)'  # clean 7 digit HD number is group 1 (match HD, 0 or infinite zeros, ref)
        self.service_desk_to = 'servicedeskplus@nbt.nhs.uk'
        self.destination_folder_name = 'Processed'
        # self.slack = SlackAPI()

    def process_emails(self):
        """
        Process outlook folder for badly-formed sdplus emails, edit subject and resend to help desk
        :return: None - Forwards emails appending ##131234## in subject and moves to a folder
        """
        mapi = self.outlook.GetNamespace('MAPI')
        # inbox = mapi.GetDefaultFolder(6)  # 6=olFolderInbox=my own inbox
        recipient = mapi.CreateRecipient('RM1048')  # 'Alias' of IT Third Party Response
        recipient.Resolve()
        if recipient.Resolved:
            # https://msdn.microsoft.com/en-us/library/office/ff869575.aspx
            self.inbox = mapi.GetSharedDefaultFolder(recipient, 6)
            messages = self.inbox.Items
            self.sdplus_api = API(self.sdplus_api_key, self.sdplus_api_url)
            print('Found ' + str(len(messages)) + ' messages to process:')
            # mailItem.Move changes inbox.Items.Count on a normal loop, but working from count to 0 works well
            for no in range(messages.Count-1, -1, -1):
                message = messages[no]
                hd = self.hd_ref_from_email(message)
                # sdplus clean, subject
                if re.search(self.sdplus_clean, message.Subject):
                    sdplus_found_number = re.search(self.sdplus_clean, message.Subject).group(1)
                    print(sdplus_found_number + ': sdplus clean, subject')
                    if self.sdplus_valid(sdplus_found_number):
                        self.update_sdplus(sdplus_found_number, 'Supplier Ref', hd)
                        # self.slack_warn_if_not_assigned(sdplus_found_number)
                        self.send_move(message)
                # sdplus, subject
                elif re.search(self.sdplus_csc, message.Subject):
                    sdplus_found_number = re.search(self.sdplus_csc, message.Subject).group(1)
                    print(sdplus_found_number + ': sdplus, subject')
                    if self.sdplus_valid(sdplus_found_number):
                        self.update_sdplus(sdplus_found_number, 'Supplier Ref', hd)
                        # self.slack_warn_if_not_assigned(sdplus_found_number)
                        self.send_move(message, ' ##' + sdplus_found_number + '##')
                # sdplus, body
                elif re.search(self.sdplus_csc, message.Body):
                    sdplus_found_number = re.search(self.sdplus_csc, message.Body).group(1)
                    print(sdplus_found_number + ': sdplus, body')
                    if self.sdplus_valid(sdplus_found_number):
                        self.update_sdplus(sdplus_found_number, 'Supplier Ref', hd)
                        # self.slack_warn_if_not_assigned(sdplus_found_number)
                        self.send_move(message, ' ##' + sdplus_found_number + '##')
                # sdplus, body, remove_no #s
                elif re.search(self.sdplus_csc, message.Body.replace('#', '')):
                    sdplus_found_number = re.search(self.sdplus_csc, message.Body.replace('#', '')).group(1)
                    print(sdplus_found_number + ': sdplus, body, remove_no #s')
                    if self.sdplus_valid(sdplus_found_number):
                        self.update_sdplus(sdplus_found_number, 'Supplier Ref', hd)
                        # self.slack_warn_if_not_assigned(sdplus_found_number)
                        self.send_move(message, ' ##' + sdplus_found_number + '##')
                else:
                    print("Can't work out sdplus number")

    def sdplus_valid(self, sdplus_ref):
        result = self.sdplus_api.send(sdplus_ref, 'GET_REQUEST')
        if result['response_status'] == 'Success':
            return True
        else:
            print('SDPlus number found (' + sdplus_ref + ') is not valid.')
            return False

    def hd_ref_from_email(self, message):
        if re.search(self.hd_ref, message.Subject):
            return re.search(self.hd_ref, message.Subject).group(1)
        if re.search(self.hd_ref, message.Body):
            return re.search(self.hd_ref, message.Body).group(1)

    def update_sdplus(self, sdplus_ref, field, value):
        call_current_values = self.sdplus_api.send(sdplus_ref, 'GET_REQUEST')
        if field in call_current_values or field.lower() in call_current_values:
            response = self.sdplus_api.send(sdplus_ref, 'EDIT_REQUEST', {field: value})
            if response['response_status'] == 'Success':
                return True
            else:
                return False

    # def _is_assigned(self, sdplus):
    #     # Check if sdplus call has an Assignee
    #     call_details = self.sdplus_api.send(sdplus, 'GET_REQUEST')
    #     call_details = dict((k.lower(), v) for k, v in call_details.items())
    #     if call_details['technician']:
    #         return True
    #     else:
    #         return False
    #
    # def slack_warn_if_not_assigned(self, sdplus_ref):
    #     if not self._is_assigned(sdplus_ref):
    #         sdplus_href = '<http://sdplus/WorkOrder.do?woMode=viewWO&woID={sdplus_ref}|{sdplus_ref}>'\
    #             .format(sdplus_ref=sdplus_ref)
    #         # Send message to backoffice group, with @mitch, @simon, @paul
    #         self.slack.send('G1FBB4L68', 'Hey <@U1FBYK4BZ>, <@U1F4X362D>, <@U1FA6DMFV> - CSC have responded '
    #                                      'to SDPlus {0}, but this is currently unassigned...'.format(sdplus_href))

    def send_move(self, mail_item, append_to_subject=''):
        new_mail = mail_item.Forward()
        new_mail.To = self.service_desk_to
        if append_to_subject:
            new_mail.Subject += append_to_subject
        self.remove_signature(new_mail)
        self.insert_line_at_top(new_mail, 'To respond to CSC, please use: ', 'cssc_nhs_it_helpdesk@csc.com')
        print('Sending.')
        new_mail.Send()
        print('Moving.')
        mail_item.Move(self.inbox.Folders(self.destination_folder_name))
        print('Processed mail successfully.')

    @staticmethod
    def remove_signature(mail_object):
        """
        Most not possible without this link:
        https://msdn.microsoft.com/en-us/library/dd492012(v=office.12).aspx
        :param mail_object: mailItem object
        :return: (removes signature)
        """
        # Remove my default signature
        # https://msdn.microsoft.com/en-us/library/dd492012(v=office.12).aspx
        print('Attempting to remove_no signature.')
        if mail_object.BodyFormat == 1:  # Plain Text:
            find_text = '-----Original Message-----'
            mail_object.Body = mail_object.Body[mail_object.Body.find(find_text) + len(find_text):]
        if mail_object.BodyFormat == 2 or mail_object.BodyFormat == 3:  # HTML or RTF
            try:
                active_inspector = mail_object.GetInspector
                word_doc = active_inspector.WordEditor
                word_bookmark = word_doc.Bookmarks('_MailAutoSig')
                if word_bookmark:
                    word_bookmark.Select()
                    word_doc.Windows(1).Selection.Delete()  # (HTMLBody property only updated after Display() is called)
                    # assumes you will later call active_inspector.Close(0) (necessary if not calling .Display())
                    print('Signature removal success.')
            except pywintypes.com_error:
                print('Failed to remove_no signature.')
        return

    @staticmethod
    def insert_line_at_top(mail_obj, line, email):
        """
        Most not possible without this link:
        https://msdn.microsoft.com/en-us/library/dd492012(v=office.12).aspx
        :param mail_obj: mailItem object
        :param line: Line to put in top of email
        :param email: Email to immediate right of line
        :return: (updates mailItem object)
        """
        if mail_obj.BodyFormat == 1:  # Plain text
            mail_obj.Body = line + email + '\n' + mail_obj.Body
        elif mail_obj.BodyFormat == 2 or mail_obj.BodyFormat == 3:  # HTML or RTF
            active_inspector = mail_obj.GetInspector
            word_doc = active_inspector.WordEditor
            word_selection = word_doc.Windows(1).Selection
            word_selection.Move(6, -1)  # 6=wdStory. Move to beginning
            word_doc.Characters(1).InsertBefore(line)
            word_selection.Move(5, 1)  # 5=wdLine. Move down 1 line
            word_selection.Move(1, -1)  # 1=wdCharacter. Move back 1
            word_doc.Hyperlinks.Add(word_selection.Range, 'mailto:' + email, '', '', email, '')
            active_inspector.Close(0)  # olSave = 0, olDiscard = 1, olPromptForSave = 2


if __name__ == '__main__':
    try:
        if not os.environ['SDPLUS_ADMIN']:
            print('Needed environment variable not found')
            sys.exit(1)  # exit as error
    except KeyError:
        print('Needed environment variable not found')
        sys.exit(1)  # exit as error

    o = OutlookSDPlus()
    o.process_emails()

# # Compilation:
# from custom_modules.compile_helper import CompileHelp
# c = CompileHelp(r'C:\simon_files_compilation_zone\outlook_sdplus')
# # c.create_env('pypiwin32 requests xmltodict slackclient')
# c.freeze(r'K:\Coding\Python\nbt work\outlook_sdplus.spec', [r'K:\Coding\Python\nbt work\outlook_sdplus.py'],
#          r'C:\Program Files\Outlook SDPlus')

"""
...aka:
rem ===Environment Setup===
cd C:\simon_files_compilation_zone\outlook_sdplus
rem ---Standard setup---
python -m venv env
env\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install pyinstaller
rem ---Specific modules---
python -m pip install pypiwin32 requests
pip freeze

rem ===Update/Compile===
cd C:\simon_files_compilation_zone\outlook_sdplus
cd env\Scripts\activate.bat
rem ---Copy source files in---
copy /y "K:\Coding\Python\nbt work\outlook_sdplus.py" outlook_sdplus.py
copy /y "K:\Coding\Python\nbt work\outlook_sdplus.spec" outlook_sdplus.spec
rem ---Compile---
rd /S /Q dist
rd /S /Q build
pyinstaller outlook_sdplus.spec
rem ---Copy exe to program files---
copy /y dist\outlook_sdplus.exe "C:\Program Files\Outlook SDPlus\outlook_sdplus.exe"
"""