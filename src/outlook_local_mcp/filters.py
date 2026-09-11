"""Exact Python predicates and conservative, generated-only DASL filters."""

import re
import sys
from collections.abc import Callable
from datetime import UTC, datetime, timedelta, tzinfo

from .enums import EErrorCode
from .errors import OutlookError
from .models import EmailSummary, SearchArguments
from .outlook_constants import (
    RECEIVED_TIME_DASL_PROPERTY,
    WINDOWS_DATE_SHORTDATE,
    WINDOWS_TIME_NOSECONDS,
)


def local_timezone() -> tzinfo:
    if sys.platform == "win32":
        from win32timezone import TimeZoneInfo

        zone: tzinfo = TimeZoneInfo.local()
        return zone
    return datetime.now().astimezone().tzinfo or UTC


def parse_date(value: str | None, zone: tzinfo | None = None) -> datetime | None:
    if value is None:
        return None
    try:
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
            parsed = datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=zone or local_timezone())
        else:
            if not re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}", value):
                raise ValueError
            parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if parsed.tzinfo is None or parsed.utcoffset() is None:
            raise ValueError
        parsed.astimezone(UTC)  # Reject boundaries that overflow the UTC comparison domain.
        return parsed
    except (ValueError, OverflowError):
        raise OutlookError(
            EErrorCode.INVALID_ARGUMENT,
            "Dates must be YYYY-MM-DD or ISO 8601 date/time with a timezone.",
        ) from None


def date_range(arguments: SearchArguments) -> tuple[datetime | None, datetime | None]:
    after, before = parse_date(arguments.after), parse_date(arguments.before)
    if after is not None and before is not None and after >= before:
        raise OutlookError(EErrorCode.INVALID_ARGUMENT, "after must be earlier than before.")
    return after, before


def outlook_date(value: datetime) -> str:
    """Use Windows regional formatting, in UTC, without seconds."""
    import win32api

    naive = value.astimezone(UTC).replace(tzinfo=None)
    locale = win32api.GetUserDefaultLCID()
    day = win32api.GetDateFormat(locale, WINDOWS_DATE_SHORTDATE, naive)
    time = win32api.GetTimeFormat(locale, WINDOWS_TIME_NOSECONDS, naive)
    return f"{day} {time}"


def restrict_filter(
    after: datetime | None,
    before: datetime | None,
    formatter: Callable[[datetime], str] = outlook_date,
) -> str | None:
    clauses = []
    property_name = f'"{RECEIVED_TIME_DASL_PROPERTY}"'
    for boundary, operator in ((after, ">="), (before, "<=")):
        if boundary is None:
            continue
        rounded = boundary.astimezone(UTC).replace(second=0, microsecond=0)
        if operator == "<=":
            # Widen even an exact-minute upper bound; Python enforces exclusivity.
            try:
                rounded += timedelta(minutes=1)
            except OverflowError:
                continue
        literal = formatter(rounded).replace("'", "''")
        clauses.append(f"{property_name} {operator} '{literal}'")
    return "@SQL=" + " AND ".join(clauses) if clauses else None


def metadata_matches(
    email: EmailSummary,
    arguments: SearchArguments,
    after: datetime | None,
    before: datetime | None,
) -> bool:
    received = datetime.fromisoformat(email.received_at)
    if after is not None and received < after:
        return False
    if before is not None and received >= before:
        return False
    if arguments.unread is not None and email.unread != arguments.unread:
        return False
    if arguments.has_attachments is not None and email.has_attachments != arguments.has_attachments:
        return False
    if arguments.category is not None:
        if email.categories is None:
            raise OutlookError(EErrorCode.METADATA_UNAVAILABLE)
        if arguments.category.casefold() not in {
            category.casefold() for category in email.categories
        }:
            return False
    if arguments.importance is not None:
        if email.importance is None:
            raise OutlookError(EErrorCode.METADATA_UNAVAILABLE)
        if email.importance != arguments.importance:
            return False
    if arguments.sender:
        needle = arguments.sender.casefold()
        if not any(
            needle in value.casefold() for value in (email.sender.name, email.sender.email) if value
        ):
            if email.sender.name is None or email.sender.email is None:
                raise OutlookError(EErrorCode.METADATA_UNAVAILABLE)
            return False
    return True
