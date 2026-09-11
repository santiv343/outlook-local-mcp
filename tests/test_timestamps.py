"""UTC storage values must survive Outlook's misleading local DATE wrappers."""

from datetime import UTC, datetime, timedelta
from types import SimpleNamespace

import pytest

from outlook_local_mcp.mail import detail, summary
from outlook_local_mcp.models import ReadArguments, SearchArguments
from outlook_local_mcp.outlook import Outlook
from outlook_local_mcp.outlook_constants import (
    MODIFIED_TIME_DASL_PROPERTY,
    RECEIVED_TIME_DASL_PROPERTY,
    SENT_TIME_DASL_PROPERTY,
)

from .fakes import Application, ComFailure, Mail


def native_mail(local_date, utc_date, sent_date=None):
    values = {
        RECEIVED_TIME_DASL_PROPERTY: utc_date,
        SENT_TIME_DASL_PROPERTY: sent_date,
        MODIFIED_TIME_DASL_PROPERTY: utc_date,
    }
    item = SimpleNamespace(**vars(Mail()))
    item.ReceivedTime = local_date
    item.SentOn = local_date
    item.Sent = sent_date is not None
    item.PropertyAccessor = SimpleNamespace(GetProperty=values.__getitem__)
    return item


def test_real_projection_and_search_use_utc_storage_across_local_midnight():
    utc_date = datetime(2026, 1, 1, 2, 59, 59, 123000, tzinfo=UTC)
    item = native_mail(utc_date - timedelta(hours=3), utc_date, utc_date)
    backend = Outlook(Application([item]))
    arguments = SearchArguments(after=utc_date.isoformat(), before="2026-01-01T03:00:00Z")
    page = backend.emails(arguments)
    assert len(page.items) == 1
    assert datetime.fromisoformat(page.items[0].received_at) == utc_date
    assert not backend.emails(SearchArguments(before=utc_date.isoformat())).items
    result = detail(item, ReadArguments(entry_id=item.EntryID, store_id="store"))
    assert datetime.fromisoformat(result.sent_at) == utc_date


def test_repeated_dst_wall_time_keeps_two_distinct_instants():
    wall_time = datetime(2026, 11, 1, 1, 30, tzinfo=UTC)
    first = datetime(2026, 11, 1, 5, 30, tzinfo=UTC)
    second = first + timedelta(hours=1)
    dates = [
        datetime.fromisoformat(
            summary(native_mail(wall_time, instant), "store", "inbox").received_at
        )
        for instant in (first, second)
    ]
    assert dates == [first, second]


def test_unsent_message_has_no_invented_sent_date():
    instant = datetime(2026, 1, 15, 12, tzinfo=UTC)
    result = detail(native_mail(instant, instant), ReadArguments(entry_id="item", store_id="store"))
    assert result.sent_at is None
    assert "SENT_DATE_UNAVAILABLE" not in {warning.code for warning in result.warnings}


def test_sent_message_with_no_sent_date_reports_unavailable():
    instant = datetime(2026, 1, 15, 12, tzinfo=UTC)
    item = native_mail(instant, instant)
    item.Sent = True
    result = detail(item, ReadArguments(entry_id="item", store_id="store"))
    assert result.sent_at is None
    assert "SENT_DATE_UNAVAILABLE" in {warning.code for warning in result.warnings}


@pytest.mark.parametrize("value", [None, "not-a-date"])
def test_missing_required_utc_date_is_an_omission_not_empty_complete(value):
    instant = datetime(2026, 1, 15, 12, tzinfo=UTC)
    backend = Outlook(Application([native_mail(instant, value)]))
    result = backend.emails(SearchArguments())
    assert not result.items
    assert result.omitted == 1
    assert not result.coverage.evaluation_complete


def test_denied_required_utc_date_is_not_replaced_with_local_wall_time():
    instant = datetime(2026, 1, 15, 12, tzinfo=UTC)
    item = native_mail(instant, instant)

    def denied(name):
        raise ComFailure()

    item.PropertyAccessor.GetProperty = denied
    result = Outlook(Application([item])).emails(SearchArguments())
    assert not result.items and result.omitted == 1
    assert not result.coverage.evaluation_complete
