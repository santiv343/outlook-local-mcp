"""Validate explicit SMTP input without network lookup or name guessing."""

from email.errors import HeaderParseError
from email.headerregistry import Address

from .com_types import IAddressEntry
from .outlook_constants import SMTP_ADDRESS_PROPERTY, SMTP_ADDRESS_TYPE


def explicit_smtp(value: str) -> str:
    value = value.strip()
    try:
        parsed = Address(addr_spec=value)
        if (
            not parsed.username
            or not parsed.domain
            or any(character in value for character in "\r\n")
        ):
            raise ValueError
    except (ValueError, HeaderParseError):
        raise ValueError("Use an explicit SMTP address without a display name.") from None
    return parsed.addr_spec


def smtp_address(value: object) -> str | None:
    if isinstance(value, str) and "@" in value and not value.startswith("/"):
        return value
    return None


def address_entry_email(entry: IAddressEntry | None) -> str | None:
    if entry is None:
        return None
    try:
        if entry.Type == SMTP_ADDRESS_TYPE:
            return smtp_address(entry.Address)
        for resolve_exchange in (entry.GetExchangeUser, entry.GetExchangeDistributionList):
            try:
                exchange = resolve_exchange()
                if exchange is not None:
                    address = smtp_address(exchange.PrimarySmtpAddress)
                    if address:
                        return address
            except Exception:
                continue
        return smtp_address(entry.PropertyAccessor.GetProperty(SMTP_ADDRESS_PROPERTY))
    except Exception:
        return None
