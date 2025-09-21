from __future__ import annotations

from pathlib import Path

from ppt_translator.utils import clean_path, iter_presentation_files


def test_clean_path_strips_quotes_and_escapes():
    raw = "'~/My\\ Documents/presentation.pptx'"
    cleaned = clean_path(raw)
    assert cleaned == "~/My Documents/presentation.pptx"


def test_iter_presentation_files_returns_expected(tmp_path):
    pptx_file = tmp_path / "deck.pptx"
    ppt_file = tmp_path / "deck.ppt"
    txt_file = tmp_path / "notes.txt"
    nested_dir = tmp_path / "nested"
    nested_dir.mkdir()
    nested_pptx = nested_dir / "nested.pptx"

    pptx_file.touch()
    ppt_file.touch()
    txt_file.touch()
    nested_pptx.touch()

    results = sorted(iter_presentation_files(tmp_path))
    assert results == sorted([pptx_file, ppt_file, nested_pptx])


def test_iter_presentation_files_handles_file_input(tmp_path):
    pptx_file = tmp_path / "deck.pptx"
    pptx_file.touch()
    results = list(iter_presentation_files(pptx_file))
    assert results == [pptx_file]


def test_iter_presentation_files_missing_path(tmp_path):
    missing = tmp_path / "missing"
    results = list(iter_presentation_files(missing))
    assert results == []
