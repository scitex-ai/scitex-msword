#!/usr/bin/env python3
"""E2E layer (PS-212): real-subsystem workflows, loopback only, no network.

Covers the DOCX round trip end to end: build a Word file with python-docx
-> load into the JSON-like intermediate -> re-render with a journal
profile -> reload and confirm content survived. Gated by `RUN_E2E=1`
(skipped by default); the `e2e` marker is registered in `pyproject.toml`.
"""

from __future__ import annotations

import os

import pytest

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_E2E") != "1",
    reason="e2e layer runs only with RUN_E2E=1",
)

docx = pytest.importorskip("docx")

import scitex_msword as sxm


@pytest.mark.e2e
def test_docx_round_trip_preserves_text(tmp_path):
    """load -> save -> reload must preserve paragraph text."""
    # Arrange
    src = tmp_path / "draft.docx"
    out = tmp_path / "styled.docx"
    seed = docx.Document()
    seed.add_paragraph("Hello scitex-msword")
    seed.save(str(src))
    # Act
    sxm.save_docx(sxm.load_docx(str(src), profile="generic"), str(out))
    # Assert
    assert "Hello scitex-msword" in "\n".join(
        p.text for p in docx.Document(str(out)).paragraphs
    )
