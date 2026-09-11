from datetime import UTC, datetime, timedelta, timezone

import pytest

from outlook_local_mcp.errors import OutlookError
from outlook_local_mcp.filters import date_range, metadata_matches, parse_date, restrict_filter
from outlook_local_mcp.mail import summary
from outlook_local_mcp.models import SearchArguments

from .fakes import Mail


@pytest.mark.parametrize(
    "value",
    ["2026-02-30", "2026-01-01T12:00", "2026-01-01 12:00Z", "today", "01/02/2026"],
)
def test_invalid_dates(value):
    with pytest.raises(OutlookError, check=lambda error: error.code == "INVALID_ARGUMENT"):
        parse_date(value)


def test_date_only_uses_supplied_local_midnight():
    local = timezone(timedelta(hours=-3))
    assert parse_date("2026-01-15", local) == datetime(2026, 1, 15, 3, tzinfo=UTC)
    assert parse_date("2026-01-15T12:34:56+02:00").hour == 12


def test_windows_zone_uses_date_specific_dst():
    from win32timezone import TimeZoneInfo

    zone = TimeZoneInfo("Eastern Standard Time")
    assert parse_date("2026-01-15", zone).utcoffset() == timedelta(hours=-5)
    assert parse_date("2026-07-15", zone).utcoffset() == timedelta(hours=-4)


@pytest.mark.parametrize("before", ["2026-01-15", "2026-01-14"])
def test_inverted_or_empty_range(before):
    with pytest.raises(OutlookError, check=lambda error: error.code == "INVALID_ARGUMENT"):
        date_range(SearchArguments(after="2026-01-15", before=before))


def test_exact_bounds_and_all_metadata_filters():
    email = summary(Mail(), "store", "inbox")
    instant = datetime.fromisoformat(email.received_at)
    filters = SearchArguments(unread=True, sender="EXAMPLE.COM", has_attachments=False)
    assert metadata_matches(email, filters, instant, instant + timedelta(seconds=1))
    assert not metadata_matches(email, filters, None, instant)
    assert not metadata_matches(email, filters, instant + timedelta(microseconds=1), None)
    assert not metadata_matches(email, SearchArguments(unread=False), None, None)


@pytest.mark.parametrize("regional_format", ["%m/%d/%Y %I:%M %p", "%d/%m/%Y %H:%M"])
def test_restrict_is_conservative_in_utc_and_regionally_formatted(regional_format):
    after = parse_date("2026-01-15T09:00:30-03:00")
    before = parse_date("2026-01-15T09:01:00-03:00")
    received = []

    def formatter(value):
        received.append(value)
        return value.strftime(regional_format)

    expression = restrict_filter(after, before, formatter)
    assert received == [
        datetime(2026, 1, 15, 12, tzinfo=UTC),
        datetime(2026, 1, 15, 12, 2, tzinfo=UTC),
    ]
    assert ">=" in expression and "<=" in expression


def test_generated_regional_literals_escape_apostrophes():
    expression = restrict_filter(datetime(2026, 1, 1, tzinfo=UTC), None, lambda _: "o'clock")
    assert "'o''clock'" in expression


@pytest.mark.parametrize("value", ["0001-01-01T00:00:00+14:00", "9999-12-31T23:59:00-14:00"])
def test_dates_outside_utc_representable_range_are_rejected(value):
    with pytest.raises(OutlookError, check=lambda error: error.code == "INVALID_ARGUMENT"):
        date_range(SearchArguments(after=value))
