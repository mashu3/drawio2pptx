"""
Tests for CLI entry point (drawio2pptx.main).
"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from drawio2pptx.main import main


def test_main_success_creates_pptx(sample_drawio_path: Path, tmp_path: Path) -> None:
    """Running main with valid input creates the output pptx file."""
    out_pptx = tmp_path / "out.pptx"
    argv = ["drawio2pptx", str(sample_drawio_path), str(out_pptx)]

    with patch("sys.argv", argv):
        main()  # success path does not call sys.exit()

    assert out_pptx.exists()
    assert out_pptx.stat().st_size > 0


def test_main_missing_input_exits_with_error(tmp_path: Path) -> None:
    """Main exits with code 1 when input file does not exist."""
    out_pptx = tmp_path / "out.pptx"
    argv = ["drawio2pptx", str(tmp_path / "nonexistent.drawio"), str(out_pptx)]

    with patch("sys.argv", argv):
        with pytest.raises(SystemExit) as exc_info:
            main()

    assert exc_info.value.code == 1
    assert not out_pptx.exists()


def test_main_with_analyze_flag(sample_drawio_path: Path, tmp_path: Path) -> None:
    """Main runs and creates pptx when --analyze is passed; compare_conversion is invoked."""
    out_pptx = tmp_path / "out.pptx"
    argv = ["drawio2pptx", str(sample_drawio_path), str(out_pptx), "--analyze"]

    with patch("sys.argv", argv):
        with patch("drawio2pptx.main.compare_conversion", MagicMock()) as mock_compare:
            main()

    assert out_pptx.exists()
    mock_compare.assert_called_once_with(sample_drawio_path, out_pptx)


def test_main_short_analyze_flag(sample_drawio_path: Path, tmp_path: Path) -> None:
    """Main runs with -a (short analyze flag)."""
    out_pptx = tmp_path / "out.pptx"
    argv = ["drawio2pptx", str(sample_drawio_path), str(out_pptx), "-a"]

    with patch("sys.argv", argv):
        with patch("drawio2pptx.main.compare_conversion", MagicMock()):
            main()

    assert out_pptx.exists()


def test_main_no_diagrams_exits_with_error(sample_drawio_path: Path, tmp_path: Path) -> None:
    """Main exits with code 1 when the file contains no diagrams."""
    out_pptx = tmp_path / "out.pptx"
    argv = ["drawio2pptx", str(sample_drawio_path), str(out_pptx)]

    with patch("sys.argv", argv):
        with patch("drawio2pptx.main.DrawIOLoader") as mock_loader_cls:
            mock_loader = MagicMock()
            mock_loader.load_file.return_value = []  # no diagrams
            mock_loader_cls.return_value = mock_loader

            with pytest.raises(SystemExit) as exc_info:
                main()

    assert exc_info.value.code == 1
    assert not out_pptx.exists()


def test_main_prints_warnings_when_present(sample_drawio_path: Path, tmp_path: Path, capsys) -> None:
    """Main prints warnings when the logger reports any."""
    from drawio2pptx.logger import ConversionLogger

    out_pptx = tmp_path / "out.pptx"
    argv = ["drawio2pptx", str(sample_drawio_path), str(out_pptx)]

    real_logger = ConversionLogger()
    real_logger.warn_unsupported_effect("elem1", "test_effect", {})

    with patch("sys.argv", argv):
        with patch("drawio2pptx.main.ConversionLogger", return_value=real_logger):
            main()

    assert out_pptx.exists()
    out, _ = capsys.readouterr()
    assert "Warnings (1):" in out
    assert "test_effect" in out


def test_main_exception_exits_with_error(sample_drawio_path: Path, tmp_path: Path) -> None:
    """Main exits with code 1 and prints error when conversion raises."""
    out_pptx = tmp_path / "out.pptx"
    argv = ["drawio2pptx", str(sample_drawio_path), str(out_pptx)]

    with patch("sys.argv", argv):
        with patch("drawio2pptx.main.DrawIOLoader") as mock_loader_cls:
            mock_loader = MagicMock()
            mock_loader.load_file.side_effect = RuntimeError("load failed")
            mock_loader_cls.return_value = mock_loader

            with pytest.raises(SystemExit) as exc_info:
                main()

    assert exc_info.value.code == 1
    assert not out_pptx.exists()
