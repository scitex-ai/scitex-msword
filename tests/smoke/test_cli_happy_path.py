#!/usr/bin/env python3
"""Smoke layer (PS-211): fast subprocess CLI happy-path tests (<60s).

Runs the installed `scitex-msword` console script in a subprocess with an
isolated HOME (see conftest.py) and asserts the happy path stays green.
"""

from __future__ import annotations

import shutil
import subprocess
import sys

import pytest

_CONSOLE = shutil.which("scitex-msword")
_BASE_CMD = [_CONSOLE] if _CONSOLE else [sys.executable, "-m", "scitex_msword"]


@pytest.mark.smoke
def test_cli_version_reports_package():
    """`--version` must exit 0 and name the distribution."""
    # Arrange
    cmd = [*_BASE_CMD, "--version"]
    # Act
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    # Assert
    assert result.returncode == 0 and "scitex-msword" in result.stdout


@pytest.mark.smoke
def test_list_python_apis_shape():
    """`list-python-apis` must exit 0 and expose the core verbs."""
    # Arrange
    cmd = [*_BASE_CMD, "list-python-apis"]
    # Act
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    # Assert
    assert result.returncode == 0 and "load_docx" in result.stdout
