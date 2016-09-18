import win32com.client
import re
from custom_modules.sdplus_api_rest import API
import pywintypes
__version__ = 0.4
# 0.3 - Updated the signature remover and inserted cssc@ line 27/Jan/16
# 0.4 - Added ActiveInspector.Close(0) to save changes else they're abandoned if you don't call .Display()


class OutlookSDPlus:
    def __init__(self):
        print('Remedy email processor v' + str(__version__))
        self.outlook = win32com.client.Dispatch('Outlook.Application')
        self.inbox = None
        self.sdplus_api = None
        self.sdplus_clean = r'(?:##)(1\d{5})(?:##)'  # clean sdplus number is group 1
        self.sdplus_csc = r'(?:NBNT|NBNTSD)(1\d{5})'  # clean sdplus number is group 1
        self.service_desk_to = 'servicedeskplus@nbt.nhs.uk'
        self.destination_folder_name = 'Processed'

    def process_emails(self):
        """
        Process outlook folder for badly-formed sdplus emails, edit subject and resend to helpdesk
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
            self.sdplus_api = API('16EE6838-8160-4EFC-AEC1-0B35A59AF42C', 'http://sdplus/sdpapi/request/')
            print('Found ' + str(len(messages)) + ' messages to process:')
            message = messages.GetFirst()
            while message:  # http://timgolden.me.uk/python/win32_how_do_i/read-my-outlook-inbox.html
                # sdplus clean, subject
                if re.search(self.sdplus_clean, message.Subject):
                    sdplus_found_number = re.search(self.sdplus_clean, message.Subject).group(1)
                    print(sdplus_found_number + ': sdplus clean, subject')
                    if self.sdplus_valid(sdplus_found_number):
                        self.send_move(message)
                # sdplus, subject
                elif re.search(self.sdplus_csc, message.Subject):
                    sdplus_found_number = re.search(self.sdplus_csc, message.Subject).group(1)
                    print(sdplus_found_number + ': sdplus, subject')
                    if self.sdplus_valid(sdplus_found_number):
                        self.send_move(message, ' ##' + sdplus_found_number + '##')
                # sdplus, body
                elif re.search(self.sdplus_csc, message.Body):
                    sdplus_found_number = re.search(self.sdplus_csc, message.Body).group(1)
                    print(sdplus_found_number + ': sdplus, body')
                    if self.sdplus_valid(sdplus_found_number):
                        self.send_move(message, ' ##' + sdplus_found_number + '##')
                # sdplus, body, remove_no #s
                elif re.search(self.sdplus_csc, message.Body.replace('#', '')):
                    sdplus_found_number = re.search(self.sdplus_csc, message.Body.replace('#', '')).group(1)
                    print(sdplus_found_number + ': sdplus, body, remove_no #s')
                    if self.sdplus_valid(sdplus_found_number):
                        self.send_move(message, ' ##' + sdplus_found_number + '##')
                else:
                    print("Can't work out sdplus number")
                message = messages.GetNext()
                # Moving messages results in the loop missing mails, but loop will stop. A dangerous alternative
                # (dangerous in case folder is never empty as a mail won't move) is
                # messages = self.inbox.Items
                # message = messages.GetFirst()

    def sdplus_valid(self, sdplus_ref):
        result = self.sdplus_api.send(sdplus_ref, 'GET_REQUEST')
        if result['response_status'] == 'Success':
            return True
        else:
            print('SDPlus number found (' + sdplus_ref + ') is not valid.')
            return False

    def send_move(self, mail_item, append_to_subject=''):
        new_mail = mail_item.Forward()
        new_mail.To = self.service_desk_to
        if append_to_subject:
            new_mail.Subject += append_to_subject
        self.remove_signature(new_mail)
        self.insert_line_at_top(new_mail, 'To respond to CSC, please use: ', 'cssc_nhs_it_helpdesk@csc.com')
        print('Sending.')
        new_mail.Send()
        # new_mail.Display()
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

o = OutlookSDPlus()
o.process_emails()
