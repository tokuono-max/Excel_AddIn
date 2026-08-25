# -*- coding: utf-8 -*-
"""svc_hd_in: 出荷履歴項目 JSON の読込。"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from svc.svc_hd_in import HdInConfigError, load_hd_in_labels


def _write_json(path: Path, obj: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def test_default_labels_from_config_json(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    _write_json(
        default,
        {"LABELS": ["出荷情報１", "製品コード"], "OVERRIDE_FILE": "出荷履歴項目.json"},
    )
    labels = load_hd_in_labels(default_path=default, override_dir=tmp_path)
    assert labels == ["出荷情報１", "製品コード"]


def test_override_at_install_root_wins(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    _write_json(
        default,
        {"LABELS": ["既定１"], "OVERRIDE_FILE": "出荷履歴項目.json"},
    )
    _write_json(tmp_path / "出荷履歴項目.json", {"LABELS": ["現場１", "現場２"]})
    labels = load_hd_in_labels(default_path=default, override_dir=tmp_path)
    assert labels == ["現場１", "現場２"]


def test_missing_both_raises(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    with pytest.raises(HdInConfigError, match="見つかりません"):
        load_hd_in_labels(default_path=default, override_dir=tmp_path)


def test_broken_default_raises_when_no_override(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    default.parent.mkdir(parents=True)
    default.write_text("{ not json", encoding="utf-8")
    with pytest.raises(HdInConfigError, match="形式が正しくありません"):
        load_hd_in_labels(default_path=default, override_dir=tmp_path)


def test_broken_override_falls_back_to_default(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    _write_json(default, {"LABELS": ["既定Ａ"], "OVERRIDE_FILE": "出荷履歴項目.json"})
    (tmp_path / "出荷履歴項目.json").write_text("{ broken", encoding="utf-8")
    labels = load_hd_in_labels(default_path=default, override_dir=tmp_path)
    assert labels == ["既定Ａ"]


def test_empty_labels_in_both_raises(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    _write_json(default, {"LABELS": ["", "  "], "OVERRIDE_FILE": "出荷履歴項目.json"})
    _write_json(tmp_path / "出荷履歴項目.json", {"LABELS": []})
    with pytest.raises(HdInConfigError, match="LABELS が空"):
        load_hd_in_labels(default_path=default, override_dir=tmp_path)


def test_override_file_path_is_rejected(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    _write_json(
        default,
        {"LABELS": ["既定"], "OVERRIDE_FILE": "../secret.json"},
    )
    _write_json(tmp_path / "出荷履歴項目.json", {"LABELS": ["現場"]})
    labels = load_hd_in_labels(default_path=default, override_dir=tmp_path)
    assert labels == ["現場"]


def test_custom_override_basename(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    _write_json(default, {"LABELS": ["既定"], "OVERRIDE_FILE": "現場ヘッダ.json"})
    _write_json(tmp_path / "現場ヘッダ.json", {"LABELS": ["別名"]})
    labels = load_hd_in_labels(default_path=default, override_dir=tmp_path)
    assert labels == ["別名"]


def test_override_only_without_default_file(tmp_path: Path) -> None:
    default = tmp_path / "config" / "hd_in.json"
    _write_json(tmp_path / "出荷履歴項目.json", {"LABELS": ["現場のみ"]})
    labels = load_hd_in_labels(default_path=default, override_dir=tmp_path)
    assert labels == ["現場のみ"]


def test_shipped_hd_in_json_has_labels(tmp_path: Path) -> None:
    repo = Path(__file__).resolve().parent.parent
    labels = load_hd_in_labels(
        default_path=repo / "config" / "hd_in.json",
        override_dir=tmp_path,
    )
    assert labels[0] == "出荷情報１"
    assert labels[-1] == "拡張項目２４"
    assert "製品コード" in labels
    assert "群番/副番" in labels
    assert len(labels) == 41
