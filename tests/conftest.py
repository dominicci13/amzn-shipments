"""Shared fixtures for the amzn-shipments test suite."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


@pytest.fixture(scope="session")
def shipments():
    """Import the automation module.

    Safe to import only because the ``ask_user``/``run_on_schedule`` block at the
    bottom of the module is ``__main__``-guarded; module level otherwise just reads
    ``.env`` and ``config/paths.json``. No browser, Excel or Outlook opens.

    Returns:
        The imported ``run_amzn_shipments`` module.
    """
    import run_amzn_shipments

    return run_amzn_shipments


class _NullLog:
    """Keeps simulated failures out of the production ``logs/amzn_shipments.log``."""

    def __getattr__(self, _level: str):
        return lambda *_a, **_k: None


@pytest.fixture(autouse=True)
def quiet_log(shipments, monkeypatch):
    monkeypatch.setattr(shipments, "log", _NullLog())


@pytest.fixture(autouse=True)
def kills(shipments, monkeypatch) -> list[str]:
    """Never let a test reach the real ``kill_app``: it is ``taskkill /f /im``, machine-wide."""
    killed: list[str] = []
    monkeypatch.setattr(shipments.custom_functions, "kill_app", killed.append)
    return killed
