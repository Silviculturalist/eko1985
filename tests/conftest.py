"""Test configuration for path management."""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

for path in (SRC, ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))


@pytest.fixture(scope="session")
def artifacts_dir() -> Path:
    """Return the directory used for pytest artefacts.

    The location defaults to ``tests/artifacts`` within the repository, but it
    can be overridden by setting the ``EKO1985_PARITY_ARTIFACTS`` environment
    variable to an absolute or relative path. The directory is created if it
    does not already exist so that tests can safely write snapshot outputs.
    """

    env_value = os.environ.get("EKO1985_PARITY_ARTIFACTS")
    base_path = Path(env_value) if env_value else ROOT / "tests" / "artifacts"
    base_path.mkdir(parents=True, exist_ok=True)
    return base_path
