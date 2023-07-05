# -*- coding: utf-8 -*-
from __future__ import annotations  # allow annotations when class not defined

import psutil
import win32com.client

from pathlib import Path
from typing import List, Optional, Dict, TypeVar, Iterator


class Folder:
    """
    Implementation of C# object Folder in Python. Allow to get all COMObject
    of .Folders method, like account's names, folder's names, subfolder's names,
    e.t.c. and store in self.Folders python dictionary.
    Working with dictionary in python is easy to understand.
    https://learn.microsoft.com/en-us/office/vba/api/outlook.folder?source=recommendations
    """
    __slots__ = ('Name', 'COMObject', 'Folders')

    def __init__(self, COMObject: win32com.client.CDispatch) -> None:
        self.Name: str = COMObject.Name
        self.COMObject: win32com.client.CDispatch = COMObject  # to get COMObject you must call appropriate method
        self.Folders: Dict[str, Folder] = {
            value.Name: Folder(COMObject=value) for value in self.COMObject.Folders
            }

    def loop_trought_last_messages(self, loops: int = None) -> Iterator[Mail]:
        folder_items = self.COMObject.Items
        loops = folder_items.Count if not loops else loops
        mail = Mail(folder_items.GetLast())
        for _ in range(loops):
            yield mail
            mail = Mail(folder_items.GetPrevious())

    def __repr__(self) -> str:
        return f'(Name={self.Name}, COMObject={self.COMObject}, Folders={self.Folders})'


class OutlookAPI:
    """
    Base Outlook class.
    """

    def __new__(cls, *args, **kwargs):
        if not hasattr(cls, "_outlookapi"):
            cls._outlookapi = super().__new__(cls)
        return cls._outlookapi

    def __init__(self, account: str) -> None:
        try:
            self.outlook: win32com.client.CDispatch = win32com.client.Dispatch('Outlook.Application')
            self.namespace: win32com.client.CDispatch = self.outlook.GetNameSpace("MAPI")
            self.accounts: Dict[str, Folder] = {
                value.name: Folder(value) for value in self.namespace.Folders
                }  # .keys() - corporate accounts names or available mail's
            self.account: win32com.client.CDispatch = self.accounts.get(account).COMObject
            self.mail: Optional[Mail] = None
        except Exception as ex:
            raise SystemExit(
                "Outlook open status = %s.\n%s: %s" %
                (self.check_outlook_is_open(), type(ex).__name__, ex)
                )
        # finally:
        #     self.outlook = None

    def create_mail(self):
        self.mail = Mail(self.outlook.CreateItem(0))

    def open_saved_mail(self, path: str):
        self.mail = Mail(self.namespace.OpenSharedItem(path))

    @classmethod
    def check_outlook_is_open(cls):
        for pid in psutil.pids():
            p = psutil.Process(pid)
            if p.name() == "OUTLOOK.EXE":
                return True
        return False

    def check_aoutoresponse(self, recipient: str):
        """
        Take email address as the recipient and return aouto response.
        """
        return self.namespace.CreateRecipient(recipient).AutoResponse

    def get_default_outlook_signature(self) -> str:
        """
        Create empty message, display it and return current html body with
        signature.
        """
        mail: win32com.client.CDispatch = self.outlook.CreateItem(0)
        mail.Display()
        signature: str = mail.HTMLBody
        mail.Close(1)
        return signature

    def get_drafts_folder(self, draft_folder_name: str = 'Черновики'):
        return self.account.Folders[draft_folder_name]

    def add_folder(self, new_folder_name):
        self.account.Folders.Add(new_folder_name)
        self.accounts = {
                value.name: Folder(value) for value in self.namespace.Folders
                }

    def delete_folder(self, folder_name):
        self.account.Folders[folder_name].Delete()
        self.accounts = {
                value.name: Folder(value) for value in self.namespace.Folders
                }


class Mail():  # TODO: description
    """
    """
    __slots__ = (
        'COMObject', 'Sender', 'To', 'Subject', 'HTMLBody',
        'HTMLformat', 'Body', 'Attachments', 'CC', 'BCC'
        )

    def __init__(self, COMObject=None):  # TODO: COMObject have got property Recipients which is not implemented.
        self.COMObject = COMObject
        self.Sender: Optional[str] = None  # Mail sender
        self.To: Optional[str] = 'SampleMail@mail.com'
        self.Subject: Optional[str] = 'Sample Email'
        self.HTMLBody: Optional[str] = None
        self.HTMLformat: Optional[str] = None
        self.HTMLBody: Optional[str] = "Sample test message created."
        self.Attachments: Optional[List[str]] = []
        self.CC: Optional[str] = "SampleSomebody@mail.com"
        self.BCC: Optional[str] = None

    def create(self):
        self.COMObject = win32com.client.Dispatch('Outlook.Application').CreateItem(0)

    def send(self):
        self.COMObject.Send()

    def save(self):
        self.COMObject.Save()

    def move_to(self, folder: win32com.client.CDispatch) -> None:
        """
        Move mail object into specified folder.
        """
        self.COMObject.Move(folder)

    def close(self):
        self.COMObject.Close(1)

    def display(self):
        self.COMObject.Display()

    def mark_unread(self):
        """
        Mark mail as unread in Outlook.
        Method should be callede before moving file into another folder.
        """
        self.COMObject.UnRead = True

    def add_recipients(self, recipients: List[str]) -> None:
        """
        Add list of email's recipients in email addres field 'To'.
        """
        for recipient in recipients:
            self.COMObject.Recipients.Add(recipient)

    def add_attachments(self, attachments: Optional[List[Path]] = None) -> None:
        attachments = attachments if attachments else self.Attachments
        for attachment in attachments:
            # TODO: check condition below, probably delete
            if isinstance(attachment, Path):  # attachment's path can't be like pathlib.WindowsPath otherwise com_error: (-2147024809, 'Параметр задан неверно.', None, None)
                attachment = str(attachment)
            self.COMObject.Attachments.Add(attachment)

    def remove_attachemnt(self):
        """
        Remove first attachment from email.
        """
        self.COMObject.Attachments.Remove(1)

    def html_body_format(self, characters: List[str]):
        self.COMObject.HTMLBody = self.HTMLBody.format(*characters)

    def _change_sender(self):
        self.COMObject._oleobj_.Invoke(
            *(64209, 0, 8, 0 ,
              self.outlook.Session.Accounts[self.Sender]
              )
            )

    def get_sender(self):
        print(
            self.COMObject.Sender.GetExchangeUser().PrimarySmtpAddress,
            self.COMObject.SenderEmailAddres
            )
        return self.COMObject.SenderEmailAddres


    def get_category(self, separator: str = '; ') -> List[str]:
        return self.COMObject.Categories.split(separator)

    def get_class(self):
        """
        IPM.Note - mail
        IPM.Document.Excel.Sheet.12 - excel file
        https://docs.microsoft.com/ru-ru/office/vba/outlook/concepts/forms/item-types-and-message-classes
        """
        return self.COMObject.MessageClass

    def get_senton_date(self):
        return self.COMObject.SentOn.strftime("%d.%m.%Y")

    def get_conversation_topic(self) -> str:
        return self.COMObject.ConversationTopic

    @staticmethod
    def check_valid_mail_address(address) -> bool:
        """
        Check if the mail address validated against address book.
        Return True if recipient.Address else False.
        """
        return address.Resolved

    def get_recipients_mail_address(self) -> Iterator[str]:
        """
        Get mail addres with domen name, example: ...@domen.com.
        print(recip.Address) - some domen names and other info like:
        """
        for recipient in self.COMObject.Recipients:
            if self.check_valid_mail_address(recipient):
                yield recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress # email
            else:
                print(recipient, f"addres={recipient.Address}") # email without mail addres
                pass

    def get_recipients_address(self) -> List[str]:
        return [
            recipient.AddressEntry.GetExchangeUser().Name
            for recipient in self.COMObject.Recipients
            if self.check_valid_mail_address(recipient)
            ]
