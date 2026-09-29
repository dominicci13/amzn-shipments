"""Tests for the bounded, checked workbook refresh.

Until 2026-09-28 ``main`` called ``refresh_workbook(wait=0)`` once with no time bound:
a modal on the hidden Excel (the old macro's ``MsgBox``, or a Save As) hung the run
silently, and a transient COM error crashed it outright.

Public repo: every path and value here is fabricated.
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest
import pywintypes

from seller_automation_utils import WorkbookRefreshError

WB = "C:/fake/Shipments.xlsm"


def _com_error() -> pywintypes.com_error:
    return pywintypes.com_error(-2147352567, "Exception occurred.", None, None)


class _FakeRefresh:
    """Stands in for ``refresh_workbook``, raising the queued outcomes one call at a time."""

    def __init__(self, *outcomes: BaseException | None):
        self.outcomes = list(outcomes)
        self.calls: list[dict] = []

    def __call__(self, workbook_path, *args, **kwargs):
        self.calls.append({"path": workbook_path, "args": args, **kwargs})
        outcome = self.outcomes.pop(0) if self.outcomes else None
        if outcome is not None:
            raise outcome


@pytest.fixture
def sleeps(shipments, monkeypatch) -> list[float]:
    slept: list[float] = []
    monkeypatch.setattr(shipments.time, "sleep", slept.append)
    return slept


@pytest.mark.parametrize(
    ("outcomes", "expected_calls", "expected_error"),
    [
        pytest.param([None], 1, None, id="success-first-try"),
        pytest.param([_com_error(), None], 2, None, id="com-error-retried-then-succeeds"),
        pytest.param([_com_error(), _com_error(), None], 3, None, id="succeeds-on-last-attempt"),
        pytest.param([_com_error()] * 10, 3, pywintypes.com_error, id="com-error-x3-reraises"),
        pytest.param(
            [WorkbookRefreshError("Shipments.xlsm: Power Query refresh failed")], 1,
            WorkbookRefreshError, id="macro-reason-not-retried",
        ),
        pytest.param(
            [WorkbookRefreshError("Shipments.xlsm: refresh exceeded 300s, Excel was killed")], 1,
            WorkbookRefreshError, id="timeout-not-retried",
        ),
        pytest.param(
            [_com_error(), WorkbookRefreshError("reason")], 2, WorkbookRefreshError,
            id="verdict-after-com-error-stops-retrying",
        ),
        pytest.param([FileNotFoundError(WB)], 1, FileNotFoundError, id="other-error-not-retried"),
    ],
)
def test_refresh_retries_only_com_errors_and_only_three_times(
    shipments, monkeypatch, sleeps, kills, outcomes, expected_calls, expected_error
):
    fake = _FakeRefresh(*outcomes)
    monkeypatch.setattr(shipments, "refresh_workbook", fake)

    if expected_error is None:
        shipments._refresh_with_retry(WB)
    else:
        with pytest.raises(expected_error):
            shipments._refresh_with_retry(WB)

    assert len(fake.calls) == expected_calls
    retried = sum(isinstance(o, pywintypes.com_error) for o in outcomes[: expected_calls - 1])
    assert sleeps == [shipments.REFRESH_RETRY_DELAY_SEC] * retried
    assert kills == [], "a name-based kill ends every Excel on the machine"


def test_retry_constants_match_the_spec(shipments):
    assert shipments.MAX_REFRESH_ATTEMPTS == 3
    assert shipments.REFRESH_RETRY_DELAY_SEC == 5


def test_refresh_is_time_bounded(shipments, monkeypatch, sleeps):
    """``timeout=`` is what makes the library kill only its own Excel, by pid, on a hang."""
    fake = _FakeRefresh(None)
    monkeypatch.setattr(shipments, "refresh_workbook", fake)

    shipments._refresh_with_retry(WB)

    call = fake.calls[0]
    assert call["path"] == WB
    assert call["wait"] == 0
    assert call["timeout"] == shipments.REFRESH_TIMEOUT_SEC
    assert isinstance(call["timeout"], int) and call["timeout"] >= 300
    assert "macro_name" not in call and call["args"] == (), "must run the default modUtilities.refresh"


def test_module_has_no_machine_wide_kill(shipments):
    """Refusal guard: no kill by image name may come into this module."""
    source = Path(shipments.__file__).read_text(encoding="utf-8")

    assert "kill_app" not in source
    assert not re.search(r"taskkill|/im\b|pkill|killall", source, re.IGNORECASE)


@pytest.fixture
def pipeline(shipments, monkeypatch, sleeps) -> dict:
    """Stub every I/O step of ``main`` around the refresh; record emails and crashes.

    No accounts are iterated, so no download folder is touched.
    """
    record: dict = {"sent": [], "crashes": []}

    class _Driver:
        def quit(self):
            pass

    monkeypatch.setattr(shipments.chrome, "start_browser", lambda *_a, **_k: _Driver())
    monkeypatch.setattr(shipments.accounts, "iter_amazon_accounts", lambda: iter(()))
    monkeypatch.setattr(shipments.outlook, "send_email", lambda **kw: record["sent"].append(kw))
    monkeypatch.setattr(
        shipments.alert_utils, "handle_crash", lambda _d, tb, name: record["crashes"].append(tb)
    )
    monkeypatch.setattr(shipments, "shipments_wb_path", WB)
    return record


@pytest.mark.parametrize(
    ("failure", "error_name"),
    [
        pytest.param([_com_error()] * 3, "com_error", id="com-error-x3"),
        pytest.param([WorkbookRefreshError("reason")], "WorkbookRefreshError", id="macro-reason"),
        pytest.param([WorkbookRefreshError("refresh exceeded 300s")], "WorkbookRefreshError", id="timeout"),
    ],
)
def test_failed_refresh_crashes_loudly_and_sends_nothing(
    shipments, monkeypatch, pipeline, kills, failure, error_name
):
    """Refusal: a workbook the refresh did not vouch for is never emailed; the crash handler fires."""
    monkeypatch.setattr(shipments, "refresh_workbook", _FakeRefresh(*failure))

    with pytest.raises(SystemExit) as exit_info:
        shipments.main()

    assert exit_info.value.code == 1
    assert pipeline["sent"] == []
    assert kills == []
    assert len(pipeline["crashes"]) == 1 and error_name in pipeline["crashes"][0]


def test_successful_refresh_emails_the_workbook(shipments, monkeypatch, pipeline):
    fake = _FakeRefresh(None)
    monkeypatch.setattr(shipments, "refresh_workbook", fake)

    shipments.main()

    assert len(fake.calls) == 1 and fake.calls[0]["timeout"] == shipments.REFRESH_TIMEOUT_SEC
    assert pipeline["crashes"] == []
    assert len(pipeline["sent"]) == 1 and pipeline["sent"][0]["attachments"] == [WB]


def test_bas_is_the_hardened_function():
    """The version-controlled macro returns a reason and never raises a modal."""
    bas = (Path(__file__).resolve().parents[1] / "vba" / "modUtilities.bas").read_text(encoding="utf-8")
    code = "\n".join(line for line in bas.splitlines() if not line.lstrip().startswith("'"))

    assert "Function refresh() As String" in code
    assert "Sub refresh" not in code
    assert "MsgBox" not in code
