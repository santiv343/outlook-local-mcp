"""Synthetic COM I/O for native draft, account and conversation behavior."""

from types import SimpleNamespace

from outlook_local_mcp.outlook_constants import DRAFTS_FOLDER

from .fakes import Collection, ComFailure, Folder, Mail


class Recipient:
    def __init__(self, address, kind=1):
        self.Name = address
        self.Type = kind
        self.Resolved = True
        self.AddressEntry = SimpleNamespace(Type="SMTP", Address=address)

    def Resolve(self):
        return self.Resolved


class Recipients(Collection):
    def Add(self, address):
        recipient = Recipient(address)
        self.items.append(recipient)
        return recipient

    def ResolveAll(self):
        return all(recipient.Resolved for recipient in self.items)


class Draft:
    def __init__(self, folder, account=None):
        self.__dict__.update(Mail().__dict__)
        self.EntryID = ""
        self.Parent = folder
        self.SendUsingAccount = account
        self.BodyFormat = 1
        self.Saved = False
        self.Sent = False
        self.SentOnBehalfOfName = ""
        self.Recipients = Recipients()
        self.PropertyAccessor = SimpleNamespace(GetProperty=lambda name: "")
        self.GetInspector = SimpleNamespace(Activate=self.activate)
        self.events = []
        self._oleobj_ = SimpleNamespace(GetIDsOfNames=self.member_id, Invoke=self.put_reference)

    def member_id(self, name):
        assert name == "SendUsingAccount"
        return 99

    def put_reference(self, member, locale, flags, result, account):
        import pythoncom

        assert member == 99 and flags == pythoncom.DISPATCH_PROPERTYPUTREF and not result
        self.SendUsingAccount = account

    def Save(self):
        self.events.append("save")
        self.EntryID = self.EntryID or f"draft-{len(self.Parent.Items.items)}"
        self.Saved = True

    def Send(self):
        self.events.append("send")
        self.Sent = True

    def Display(self, modal):
        self.events.append(("display", modal))

    def activate(self):
        self.events.append("activate")

    def Reply(self):
        draft = self.SendUsingAccount.DeliveryStore.drafts.Items.Add("IPM.Note")
        draft.SendUsingAccount = self.SendUsingAccount
        draft.Subject = "RE: " + self.Subject
        draft.Body = self.Body
        draft.Recipients.Add(self.SenderEmailAddress)
        return draft

    def ReplyAll(self):
        draft = self.Reply()
        draft.Recipients.Add("copy@example.com").Type = 2
        return draft


class DraftItems(Collection):
    def __init__(self, folder):
        super().__init__()
        self.folder = folder

    def Add(self, item_type):
        assert item_type == "IPM.Note"
        draft = Draft(self.folder)
        self.items.append(draft)
        return draft


class Store:
    IsConversationEnabled = True

    def __init__(self, store_id):
        self.StoreID = store_id
        self.DisplayName = "Synthetic account store"
        self.drafts = Folder("drafts-" + store_id, store_id=store_id)
        self.drafts.Items = DraftItems(self.drafts)
        self.inbox = Folder("inbox-" + store_id, store_id=store_id)

    def GetDefaultFolder(self, kind):
        return self.drafts if kind == DRAFTS_FOLDER else self.inbox


class Namespace:
    def __init__(self):
        self.DefaultStore = Store("default")
        self.secondary = Store("secondary")
        self.Stores = Collection([self.DefaultStore, self.secondary])
        self.Accounts = Collection(
            [
                SimpleNamespace(
                    SmtpAddress="first@example.com",
                    DisplayName="First",
                    DeliveryStore=self.DefaultStore,
                ),
                SimpleNamespace(
                    SmtpAddress="second@example.com",
                    DisplayName="Second",
                    DeliveryStore=self.secondary,
                ),
            ]
        )

    def GetStoreFromID(self, store_id):
        for store in self.Stores.items:
            if store.StoreID == store_id:
                return store
        raise ComFailure(0x8004010F)

    def GetItemFromID(self, entry_id, store_id):
        store = self.GetStoreFromID(store_id)
        for folder in (store.inbox, store.drafts):
            for item in folder.Items.items:
                if item.EntryID == entry_id:
                    return item
        raise ComFailure(0x8004010F)

    def CreateRecipient(self, name):
        aliases = {account.DisplayName: account.SmtpAddress for account in self.Accounts.items}
        return Recipient(aliases.get(name, name))


class Application:
    def __init__(self):
        self.namespace = Namespace()

    def GetNamespace(self, name):
        return self.namespace


class Table:
    def __init__(self, identities):
        self.identities = identities
        self.position = 0
        self.starts = 0
        self.Columns = SimpleNamespace(Add=lambda name: None)

    def MoveToStart(self):
        self.starts += 1
        self.position = 0

    @property
    def EndOfTable(self):
        return self.position >= len(self.identities)

    def GetNextRow(self):
        store_id, entry_id = self.identities[self.position]
        self.position += 1
        return SimpleNamespace(Item=lambda name: entry_id, BinaryToString=lambda name: store_id)
