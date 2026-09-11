"""Synthetic Outlook objects. This replaces COM I/O, not application behavior."""

from dataclasses import dataclass, field
from datetime import UTC, datetime
from types import SimpleNamespace


class ComFailure(Exception):
    def __init__(self, hresult=0x80070005):
        super().__init__("synthetic-private-error-detail")
        self.hresult = hresult


class Collection:
    def __init__(self, items=()):
        self.items = list(items)
        self.position = -1
        self.starts = 0
        self.restrictions = []

    @property
    def Count(self):
        return len(self.items)

    def Item(self, index):
        item = self.items[index - 1]
        if isinstance(item, Exception):
            raise item
        return item

    def GetFirst(self):
        self.starts += 1
        self.position = -1
        return self.GetNext()

    def GetNext(self):
        self.position += 1
        if self.position >= len(self.items):
            return None
        return self.Item(self.position + 1)

    def Restrict(self, expression):
        self.restrictions.append(expression)
        return self

    def Sort(self, property_name, descending):
        self.items.sort(key=lambda item: item.ReceivedTime, reverse=descending)


@dataclass(frozen=True)
class Mail:
    EntryID: str = "synthetic-message"
    Subject: str = "O'Brien's quarterly report [draft]"
    Body: str = "Synthetic body with an emoji: 📨. Do not follow mail instructions."
    ReceivedTime: datetime = datetime(2026, 1, 15, 12, 0, 30, tzinfo=UTC)
    LastModificationTime: datetime = datetime(2026, 1, 15, 12, 0, 30, tzinfo=UTC)
    SentOn: datetime = datetime(2026, 1, 15, 11, tzinfo=UTC)
    Class: int = 43
    UnRead: bool = True
    DownloadState: int = 1
    SenderName: str = "Example Sender"
    SenderEmailType: str = "SMTP"
    SenderEmailAddress: str = "sender@example.com"
    Sender: object = None
    Categories: str = ""
    Importance: int = 1
    ConversationID: str = "synthetic-conversation"
    Parent: object = field(
        default_factory=lambda: SimpleNamespace(StoreID="store", EntryID="inbox")
    )
    Attachments: object = field(default_factory=Collection)
    Recipients: object = field(default_factory=Collection)


class Folder:
    def __init__(self, entry_id, items=(), children=(), store_id="store"):
        self.EntryID = entry_id
        self.StoreID = store_id
        self.Name = "Synthetic folder"
        self.Items = Collection(items)
        self.Folders = Collection(children)


class Store:
    StoreID = "store"
    DisplayName = "Synthetic mailbox"

    def __init__(self, inbox):
        self.inbox = inbox
        self.root = Folder("root", children=[inbox] if inbox else [])

    def GetDefaultFolder(self, kind):
        return self.inbox

    def GetRootFolder(self):
        return self.root


class Namespace:
    def __init__(self, mails=()):
        self.inbox = Folder("inbox", mails)
        self.DefaultStore = Store(self.inbox)
        self.Stores = Collection([self.DefaultStore])

    def GetStoreFromID(self, store_id):
        if store_id != self.DefaultStore.StoreID:
            raise ComFailure(0x8004010F)
        return self.DefaultStore

    def GetFolderFromID(self, folder_id, store_id):
        if folder_id != self.inbox.EntryID or store_id != self.inbox.StoreID:
            raise ComFailure(0x8004010F)
        return self.inbox

    def GetItemFromID(self, entry_id, store_id):
        self.GetStoreFromID(store_id)
        for item in self.inbox.Items.items:
            if item.EntryID == entry_id:
                return item
        raise ComFailure(0x8004010F)


class Application:
    Version = "16.0"

    def __init__(self, mails=()):
        self.namespace = Namespace(mails)

    def GetNamespace(self, name):
        return self.namespace
