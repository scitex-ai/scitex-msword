"""Isolation for the e2e layer: point HOME at a tmp dir.

Same rationale as the smoke layer (see tests/smoke/conftest.py):
user-wide config must resolve to defaults, never the operator's state.
Explicit save/restore (no monkeypatch).
"""

from __future__ import annotations

import os

import pytest


@pytest.fixture(autouse=True)
def _isolated_home(tmp_path):
    previous = os.environ.get("HOME")
    os.environ["HOME"] = str(tmp_path)
    try:
        yield
    finally:
        if previous is None:
            os.environ.pop("HOME", None)
        else:
            os.environ["HOME"] = previous
