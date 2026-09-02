# -*- coding: utf-8 -*-
"""進捗 1 行目用 I/O 参照マーク ([UNC]/[LOC]/[CCH])。"""
from __future__ import annotations

import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Sequence

from svc.data_agg_path_network import path_is_network

PROGRESS_MARK_UNC = "[UNC] "
PROGRESS_MARK_LOC = "[LOC] "
PROGRESS_MARK_CCH = "[CCH] "

_MARK_PREFIXES = (
    PROGRESS_MARK_UNC,
    PROGRESS_MARK_LOC,
    PROGRESS_MARK_CCH,
    "[UNC]",
    "[LOC]",
    "[CCH]",
    # 旧表記（進行中テキストの strip 互換）
    "[N] ",
    "[F] ",
    "[C] ",
    "[N]",
    "[F]",
    "[C]",
)


@dataclass
class FileProgressMarkStore:
    """並列抽出ワーカーとメインスレッド間でファイル単位のマークを共有する。"""

    _lock: threading.Lock = field(default_factory=threading.Lock, repr=False)
    _marks: dict[int, str] = field(default_factory=dict, repr=False)

    def set(self, file_index: int, mark: str) -> None:
        m = str(mark or "").strip()
        if not m:
            return
        if not m.endswith(" "):
            m = m + " "
        with self._lock:
            self._marks[int(file_index)] = m

    def get(self, file_index: int, io_path: str | Path) -> str:
        with self._lock:
            hit = self._marks.get(int(file_index))
        if hit:
            return hit
        return progress_io_ref_mark(io_path)


def progress_scan_mark(start_path: str | Path) -> str:
    """走査起点パスから [UNC] または [LOC]。"""
    return PROGRESS_MARK_UNC if path_is_network(start_path) else PROGRESS_MARK_LOC


def progress_io_ref_mark(io_path: str | Path) -> str:
    """
    実 I/O パスから動的マーク。
    キャッシュ命中 → [CCH]、ネットワーク読込 → [UNC]、それ以外のローカル読込 → [LOC]。
    """
    try:
        from svc.svc_data_agg_extract import xlsx_workbook_path_cached  # noqa: WPS433

        if xlsx_workbook_path_cached(io_path):
            return PROGRESS_MARK_CCH
    except Exception:
        pass
    if path_is_network(io_path):
        return PROGRESS_MARK_UNC
    return PROGRESS_MARK_LOC


def extract_progress_io_mark_prefix(text: str) -> str:
    """文字列先頭の [UNC]/[LOC]/[CCH] マーク（スペース付き）を返す。"""
    s = str(text or "").lstrip()
    for prefix in _MARK_PREFIXES:
        if s.startswith(prefix):
            return prefix if prefix.endswith(" ") else (prefix + " ")
    return ""


def strip_progress_io_mark(text: str) -> str:
    """詳細行から先頭マークを除去（1 行目へ移したあとの二重表示防止）。"""
    s = str(text or "").lstrip()
    for prefix in _MARK_PREFIXES:
        if s.startswith(prefix):
            return s[len(prefix) :].lstrip()
    return s


def progress_phase_with_mark(
    phase_head: str,
    *,
    mark: str,
    cur_file: str = "",
) -> str:
    """1 行目: マーク + フェーズ + 任意でファイル名。"""
    head = str(phase_head or "").strip() or "準備中"
    m = str(mark or "")
    if m:
        head = "%s%s" % (m, head)
    cf = str(cur_file or "").strip()
    if cf and cf not in head:
        head = "%s — %s" % (head, cf)
    return head


@dataclass
class ProgressIoMarkState:
    """同一ファイル内の項目進捗などでマークを維持する。"""

    last_mark: str = ""
    last_fi: int = 0

    def resolve(
        self,
        *,
        suffix: str = "",
        io_path: str | Path | None = None,
        file_index: int | None = None,
    ) -> str:
        from_prefix = extract_progress_io_mark_prefix(suffix)
        fi = int(file_index or 0)
        if from_prefix:
            self.last_mark = from_prefix
            self.last_fi = fi
            return from_prefix
        if io_path is not None and str(io_path).strip():
            mark = progress_io_ref_mark(io_path)
            self.last_mark = mark
            self.last_fi = fi
            return mark
        if fi and fi == self.last_fi:
            return self.last_mark
        return ""


def apply_batch_hook_io_mark(
    phase_txt: str,
    detail_txt: str,
    *,
    suffix: str,
    io_paths: Sequence[str | Path],
    file_index: int | None,
    mark_state: ProgressIoMarkState,
) -> tuple[str, str]:
    """本番一括 hook: 1 行目 phase にマーク、2 行目 detail からマーク除去。"""
    io_path: str | Path | None = None
    if file_index is not None:
        try:
            ix = int(file_index) - 1
            if 0 <= ix < len(io_paths):
                io_path = io_paths[ix]
        except (TypeError, ValueError):
            pass
    mark = mark_state.resolve(
        suffix=suffix,
        io_path=io_path,
        file_index=file_index,
    )
    detail_out = strip_progress_io_mark(detail_txt or suffix)
    phase_out = progress_phase_with_mark(phase_txt, mark=mark)
    return phase_out[:120], detail_out[:120] if detail_out else ""
