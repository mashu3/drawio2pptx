"""Shared fixtures: project root and sample paths."""

from pathlib import Path

import pytest

# Repository root (drawio2pptx/)
ROOT_DIR = Path(__file__).resolve().parent.parent


@pytest.fixture
def sample_dir() -> Path:
    """Path to the sample/ directory."""
    return ROOT_DIR / "sample"


@pytest.fixture
def sample_drawio_path(sample_dir: Path) -> Path:
    """Path to sample/sample.drawio."""
    path = sample_dir / "sample.drawio"
    if not path.exists():
        pytest.skip(f"Sample file not found: {path}")
    return path


@pytest.fixture
def sample_pptx_path(sample_dir: Path) -> Path:
    """Path to sample/sample.pptx, or test_output.pptx / output.pptx."""
    for name in ("sample.pptx", "test_output.pptx", "output.pptx"):
        path = sample_dir / name if name == "sample.pptx" else ROOT_DIR / name
        if path.exists():
            return path
    pytest.skip("No sample PPTX file found")
