# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_data_agg_debug.py
Purpose: データ集約デバッグウィンドウ（要求定義 §3.1.3）。文言・列見出し・ツールチップは config/ui_data_agg.json の SCREENS.DEBUG（TIP_*）。
History (latest 3):
  - 2026-08-27 シナリオ結果の #n[項目] 展開: 形式外1行で全体失敗しない（POWが連携キー1列のままになる対策）。
  - 2026-08-26 シナリオデバッグの前置「・」は live_items の carry_empty を参照（scenario_for_dry_run 未渡し対策）。
  - 2026-08-26 デバッグ結果一覧: 前置保持(carry_empty)対象の項目名文頭に「・」（本番一括は対象外）。
"""
from __future__ import annotations

import copy
import logging
import queue
import re
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Sequence

from PySide6.QtCore import QEventLoop, QObject, Qt, QThread, QTimer, Signal
from PySide6.QtGui import (
    QBrush,
    QCloseEvent,
    QColor,
    QFont,
    QFontMetrics,
    QPainter,
    QPen,
    QTextCursor,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QSplitter,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

_logger = logging.getLogger(__name__)
try:
    from core.core_log import get_logger as _get_core_logger

    _diag_logger = _get_core_logger(__name__)
except Exception:
    _diag_logger = _logger

from core.core_log import get_data_agg_diag_logger
from ui_qt.ui_common import _deep_merge, _normalize_tooltip_text, _normalize_message_newlines, set_widget_tooltip
from ui_qt.ui_common import create_progress_dialog
_data_agg_probe_log = get_data_agg_diag_logger()

from svc.data_agg_master_preview import (
    FROZEN_SNAPSHOT_VERSION,
    master_preview_one_shot_eligible,
    preview_compute_file_paths,
    run_preview_compute,
    scenario_for_stepped_preview,
)
from svc.data_agg_name_extract_summary import name_extract_debug_slot_editor_lines
from svc.data_agg_source_list_display import scenario_source_tooltip_plain
from svc.data_agg_source_ui import source_ui_block
from ui_qt import ipc_file

_NE_DETAIL_NAME_CACHE: dict[str, Any] | None = None
_NE_DETAIL_CELL_CACHE: dict[str, Any] | None = None


def _none_tips(n: int) -> list[str | None]:
    out: list[str | None] = [None] * n
    return out


def _ne_detail_name_cfg() -> dict[str, Any]:
    """SCENARIO_EDIT.DETAIL_NAME（名前から取得フォームのラベル・列挙）。循環 import 回避のため JSON を直接読む。"""
    global _NE_DETAIL_NAME_CACHE
    if _NE_DETAIL_NAME_CACHE is not None:
        return _NE_DETAIL_NAME_CACHE
    try:
        import json
        from pathlib import Path

        from core.core_cst import resolve_config_file_path

        p = resolve_config_file_path("ui_data_agg.json")
        raw = json.loads(p.read_text(encoding="utf-8"))
        dn = ((raw.get("SCREENS") or {}).get("SCENARIO_EDIT") or {}).get("DETAIL_NAME")
        _NE_DETAIL_NAME_CACHE = dict(dn) if isinstance(dn, dict) else {}
    except Exception:
        _NE_DETAIL_NAME_CACHE = {}
    return _NE_DETAIL_NAME_CACHE


def _ne_detail_cell_cfg() -> dict[str, Any]:
    """SCENARIO_EDIT.DETAIL_CELL（セル座標ツールチップ用）。JSON を直接読む。"""
    global _NE_DETAIL_CELL_CACHE
    if _NE_DETAIL_CELL_CACHE is not None:
        return _NE_DETAIL_CELL_CACHE
    try:
        import json
        from pathlib import Path

        from core.core_cst import resolve_config_file_path

        p = resolve_config_file_path("ui_data_agg.json")
        raw = json.loads(p.read_text(encoding="utf-8"))
        dc = ((raw.get("SCREENS") or {}).get("SCENARIO_EDIT") or {}).get("DETAIL_CELL")
        _NE_DETAIL_CELL_CACHE = dict(dc) if isinstance(dc, dict) else {}
    except Exception:
        _NE_DETAIL_CELL_CACHE = {}
    return _NE_DETAIL_CELL_CACHE


MAX_PHASE_SLOTS_DEFAULT = 16
MAX_VALUE_ROWS_DEFAULT = 50

# マスタ実行ログ: タイムスタンプ直後のインデント（1段=項目、2段=シナリオ）
_LOG_INDENT_COLS_PER_LEVEL = 4
# ログは先頭が最新。行数がこれを超えたら末尾（古い方）を削除
DEBUG_LOG_MAX_LINES = 2500

# デバッグ結果エリア: 見出し・フェーズ列の薄いグレー、マスタ一覧のシナリオ登録済み項目の色
_DEBUG_RESULTS_HEADER_BG = "#e8e8e8"
# マスタプレビュー: 全項目スナップショット取得済みで閲覧可能なときの左上コーナー／スナップショット表示中の項目名帯
_DEBUG_RESULTS_SNAPSHOT_TINT_BG = "#ddeef9"
_DEBUG_RESULTS_SNAPSHOT_TINT_QCOLOR = QColor(221, 238, 249)
_DEBUG_SUMMARY_PHASE_COL_BG = QColor(232, 232, 232)
_DEBUG_MASTER_REGISTERED_NAME_COLOR = QColor(0, 51, 153)  # 濃い青
_DEBUG_MASTER_REGISTERED_ROW_BG = QColor(245, 240, 232)  # 薄ベージュ（登録行）
_DEBUG_MASTER_ACTIVE_ROW_BG = QColor(228, 212, 188)  # 濃いベージュ（実行中・選択中）
_VALUE_GRID_PHASE_LINE_COLOR = QColor(0, 51, 153)
_VALUE_GRID_PHASE_LINE_WIDTH = 1


def phase_start_columns_from_spans(
    spans: list[tuple[int, int]],
    ncols: int,
    *,
    scenario_mode: bool,
) -> frozenset[int]:
    """シナリオ結果一覧でフェーズ境界になる列（先頭列は除く）。"""
    if not scenario_mode or ncols <= 1:
        return frozenset()
    starts: set[int] = set()
    for start, _end in spans:
        if start > 0:
            starts.add(int(start))
    if spans:
        last_end = max(int(end) for _start, end in spans)
        if last_end + 1 < ncols:
            starts.add(last_end + 1)
    return frozenset(c for c in starts if 0 < c < ncols)


def _paint_value_grid_phase_divider(
    painter: QPainter, rect, logical_col: int, starts: frozenset[int]
) -> None:
    if logical_col not in starts:
        return
    painter.save()
    try:
        pen = QPen(_VALUE_GRID_PHASE_LINE_COLOR)
        pen.setWidth(_VALUE_GRID_PHASE_LINE_WIDTH)
        pen.setCapStyle(Qt.PenCapStyle.FlatCap)
        painter.setPen(pen)
        x = rect.left() + _VALUE_GRID_PHASE_LINE_WIDTH // 2
        painter.drawLine(x, rect.top(), x, rect.bottom())
    finally:
        painter.restore()

COND_KEYS_DEFAULT = [
    "ファイル検索",
    "シート名検索",
    "主キー",
    "連携キー",
    "結合キー",
]

COND_KEYS_NAME_EXTRACT_DEFAULT = [
    "ファイル検索",
    "抜取り文字",
    "関連付け",
]

SUMMARY_HEADERS_NAME_EXTRACT_SHORT_DEFAULT = [
    "対象ファイル",
    "検索条件",
    "抜取り文字条件",
    "関連付け",
]

SUMMARY_HEADERS_NAME_EXTRACT_LOG_SHORT_DEFAULT = [
    "ファイルフィルタ後の対象ファイル数",
    "検索条件に一致したファイル数",
    "ユニーク値数（抜取り文字条件）",
    "関連付けでパスが一致した反復数",
]

SUMMARY_HEADERS_LOG_DEFAULT = [
    "ファイルフィルタ検索件数",
    "シート名件数",
    "主キー総数",
    "連携取得数/定義数",
    "結合取得数/定義数",
]

SUMMARY_HEADERS_2L_DEFAULT = [
    "ファイル\nフィルタ",
    "シート名\n件数",
    "主キー\n総数",
    "連携\n(取得/定義)",
    "結合キー\n(取得/定義)",
]


def _summary_metric_cell_display(val: str) -> str:
    """結果サマリ表の指標セル: 未設定・該当なしの '-' は空表示（内部リストは従来どおり）。"""
    t = str(val).strip()
    return "" if t == "-" else str(val)


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _cap_list_capped(lst: list[str], cap: int) -> list[str]:
    if len(lst) <= cap:
        return list(lst)
    out = list(lst[:cap])
    out.append("…（以降省略・上限%d件）" % cap)
    return out


def _dash_row(n: int) -> list[str]:
    return ["-"] * n


def _slot_third_column(slot: dict[str, Any]) -> str:
    """シナリオ編集に寄せた表示（要約に頼らず詳細を列挙）。"""
    if slot.get("editor_lines"):
        return "\n".join(slot["editor_lines"])
    d = slot.get("details") or []
    if not d:
        return slot.get("short", "")
    return " / ".join("%s=%s" % (k, v) for k, v in d)


def _format_condition_step_tooltip(heading: str, slot: dict[str, Any]) -> str:
    """条件ステップのツールチップ: 大項目見出し + 当該スロットの設定行。"""
    h = str(heading or "").strip() or "—"
    el = slot.get("editor_lines")
    lines: list[str] = []
    if isinstance(el, list) and el:
        for x in el:
            for s in str(x).split("\n"):
                t = s.strip()
                if t:
                    lines.append(t)
    if not lines:
        txt = _slot_third_column(slot)
        lines = [s.strip() for s in str(txt).split("\n") if s.strip()]
    if not lines:
        return h
    return "%s\n%s" % (h, "\n".join(lines))


def _parse_hash_bracket_column_values(colvals: list[str]) -> dict[str, list[str]]:
    """値列内の #n[項目名] 値 形式を項目名ごとにまとめる（結合フェーズ列用）。"""
    out: dict[str, list[str]] = {}
    for v in colvals:
        m = re.match(r"^#(\d+)\[([^\]]*)\]\s*(.*)$", str(v))
        if not m:
            continue
        tgt = (m.group(2) or "").strip() or "未指定"
        out.setdefault(tgt, []).append(str(m.group(3)))
    return out


def expand_hash_bracket_value_groups(
    colvals: list[str],
) -> list[tuple[str, list[str]]] | None:
    """
    シナリオ結果の連携／結合列: #n[項目名] 値 を項目列へ展開する。

    1 行でも形式外があると従来は展開全体を諦めて「連携キー」1 列のままになっていた。
    形式に合う行だけを採用し、1 件でも取れれば展開する（形式外は無視）。
    展開できないときは None。
    """
    groups: dict[str, tuple[str, list[str]]] = {}
    skipped = 0
    for v in colvals:
        m = re.match(r"^#(\d+)\[([^\]]*)\]\s*(.*)$", str(v))
        if not m:
            skipped += 1
            continue
        key = m.group(1)
        tgt = (m.group(2) or "").strip() or "未指定"
        val = m.group(3)
        if key not in groups:
            groups[key] = (tgt, [])
        groups[key][1].append(val)
    if not groups:
        return None
    if skipped:
        try:
            _diag_logger.info(
                "[DATA_AGG_DEBUG] hash_bracket_expand skip_non_matching=%s matched_keys=%s",
                skipped,
                len(groups),
            )
        except Exception:
            pass
    return [
        (groups[k][0], list(groups[k][1]))
        for k in sorted(groups.keys(), key=lambda x: int(x))
    ]


_CARRY_EMPTY_HEADER_MARK = "・"


def carry_empty_target_names_from_items(items: list[Any] | None) -> set[str]:
    """
    連携キー carry_empty が ON の登録先項目名集合。
    デバッグ結果一覧の見出し装飾専用（本番 Excel 見出しには使わない）。
    """
    from svc.data_agg_source_ui import source_ui_block
    from svc.svc_data_agg_extract import link_def_wants_carry_empty

    names: set[str] = set()
    for it in items or []:
        if not isinstance(it, dict):
            continue
        for src in it.get("sources") or []:
            if not isinstance(src, dict):
                continue
            ui = source_ui_block(src)
            if not isinstance(ui, dict):
                continue
            for ld in ui.get("link_defs") or []:
                if not isinstance(ld, dict):
                    continue
                if not link_def_wants_carry_empty(ld):
                    continue
                tgt = str(ld.get("item") or "").strip()
                if tgt:
                    names.add(tgt)
    return names


def decorate_debug_carry_empty_headers(
    headers: list[str] | None,
    carry_names: set[str] | None,
) -> list[str]:
    """前置保持対象の項目名文頭に「・」を付ける（既に付いていれば二重にしない）。"""
    mark = _CARRY_EMPTY_HEADER_MARK
    carry = carry_names or set()
    out: list[str] = []
    for h in headers or []:
        raw = str(h or "")
        bare = raw[len(mark) :] if raw.startswith(mark) else raw
        if bare in carry:
            out.append(mark + bare)
        else:
            out.append(bare if raw.startswith(mark) else raw)
    return out


def _link_detail_lines(link_defs: list[Any]) -> list[str]:
    """連携キー定義をデバッグ表示用に1行ずつ列挙する。"""
    lines = []
    for i, ld in enumerate(link_defs):
        if not isinstance(ld, dict):
            continue
        vss = str(ld.get("value_shape_script") or "").strip()
        extra = ""
        if vss:
            extra = " | 整形=%s" % (vss[:40] + ("…" if len(vss) > 40 else ""))
        lines.append(
            "#%d セル=%s 行=%s 列=%s 項目=%s%s"
            % (
                i + 1,
                ld.get("cell", ""),
                ld.get("row", ""),
                ld.get("col", ""),
                ld.get("item", ""),
                extra,
            )
        )
    return lines or ["（連携キー定義なし）"]


def _join_detail_lines(join_defs: list[Any]) -> list[str]:
    lines = []
    for i, jd in enumerate(join_defs):
        if not isinstance(jd, dict):
            continue
        extra = ""
        chk = jd.get("checks")
        if isinstance(chk, list):
            labels = [str(x).strip() for x in chk if str(x).strip()]
            if labels:
                extra += " 加工=%s" % "、".join(labels)
        vss = str(jd.get("value_shape_script") or "").strip()
        if vss:
            extra += " DSL=%s" % vss
        lines.append(
            "#%d セル=%s 行=%s 列=%s 項目=%s%s"
            % (
                i + 1,
                jd.get("cell", ""),
                jd.get("row", ""),
                jd.get("col", ""),
                jd.get("item", ""),
                extra,
            )
        )
    return lines or ["（結合キー定義なし）"]


def _append_extract_primaries_to_col(
    col_vals: list[str],
    primary_values: Any,
    *,
    max_rows: int,
) -> None:
    """主値を列に追加。空の primary_values は行数に数えない。"""
    if not isinstance(primary_values, list) or not primary_values:
        return
    for v in primary_values:
        if len(col_vals) >= max_rows:
            break
        col_vals.append("" if v is None else str(v))


def build_debug_scenarios_from_items(
    items: list[dict[str, Any]],
    scan_paths: list[str] | None = None,
) -> list[dict[str, Any]]:
    """
    メイン／シナリオ編集の項目・ソースからデバッグ左一覧用シナリオ行を生成する。
    フェーズ実行の数値は UI プレビュー用（完全な抽出は svc との段階的接続）。
    """
    paths = [str(x).strip() for x in (scan_paths or []) if str(x).strip()]
    n_scan = len(paths)
    nfs = str(n_scan)
    col_paths: list[str]
    if paths:
        col_paths = _cap_list_capped(paths, MAX_VALUE_ROWS_DEFAULT)
    else:
        col_paths = ["（基準フォルダ条件でファイルが0件。メインでフォルダ・拡張子を確認してください）"]
    out: list[dict[str, Any]] = []
    for item in items:
        base_title = str(item.get("name") or item.get("id") or "項目").strip() or "項目"
        srcs = item.get("sources") or []
        if not srcs:
            out.append(
                {
                    "title": base_title,
                    "summary": "（取得シナリオなし）",
                    "slots": [None] * 5,
                    "source_kind": "cell",
                    "source": None,
                }
            )
            continue

        for si, s0 in enumerate(srcs):
            if not isinstance(s0, dict):
                continue
            sn0 = str(s0.get("scenario_name") or "").strip()
            title = sn0 if sn0 else ("%s_シナリオ%d" % (base_title, si + 1))
            p = source_ui_block(s0) or {}
            typ = str(s0.get("type") or "cell").strip().lower()
            raw_join = p.get("join_defs")
            join_defs: list[Any] = raw_join if isinstance(raw_join, list) else []
            raw_link = p.get("link_defs")
            link_defs: list[Any] = raw_link if isinstance(raw_link, list) else []
            nj = len(join_defs)
            nl = len(link_defs)
            fp = str(p.get("file_pattern") or "")
            fn_rule = str(p.get("file_name_rule") or "含む")

            slots: list[dict[str, Any] | None] = [None] * 5
            if typ == "cell":
                slots[0] = {
                    "short": fp or "ファイル検索",
                    "editor_lines": [
                        "ファイル名判定: %s" % fn_rule,
                        "パターン: %s" % (fp or "—"),
                        "タブ2検出: %s 件" % nfs,
                    ],
                    "details": [("ファイル名判定", fn_rule), ("パターン", fp or "—"), ("検出件数", nfs)],
                    "summary_vals": [nfs, "-", "-", "-", "-"],
                    "values_column": list(col_paths),
                }
                rule = str(p.get("sheet_rule") or "")
                sn = str(s0.get("sheet_name") or "")
                slots[1] = {
                    "short": rule or "シート名検索",
                    "editor_lines": ["判定: %s" % (rule or "—"), "シート名入力: %s" % (sn or "—")],
                    "details": [("判定", rule or "—"), ("シート名", sn or "—")],
                    "summary_vals": [nfs, nfs, "-", "-", "-"],
                    "values_column": [sn] if sn else ["（左端シート等）"],
                }
                cref_raw = str(s0.get("cell_ref") or "").strip()
                cref = cref_raw if cref_raw else "（空＝既定）"
                slots[2] = {
                    "short": cref,
                    "editor_lines": [
                        "セル: %s" % cref,
                        "オフセット 行=%s 列=%s"
                        % (s0.get("row_offset", 0), s0.get("col_offset", 0)),
                    ],
                    "details": [("セル", cref)],
                    "summary_vals": [nfs, nfs, "1", "-", "-"],
                    "values_column": ["（抽出プレビュー）"],
                }
                link_v = "%d/%d" % (min(nl, 99), max(nl, 1)) if nl else "0/0"
                slots[3] = {
                    "short": "連携キー",
                    "editor_lines": _link_detail_lines(link_defs),
                    "details": [("登録", str(nl))],
                    "summary_vals": [nfs, nfs, "1", link_v, "-"],
                    "values_column": ["連携プレビュー"],
                    "defined": nl > 0,
                }
                join_metric = "-" if nj <= 0 else "%d/%d" % (min(nj, 9), min(nj, 9))
                slots[4] = {
                    "short": "結合キー",
                    "editor_lines": _join_detail_lines(join_defs),
                    "details": [(str(i + 1), str(jd.get("cell", ""))) for i, jd in enumerate(join_defs[:5])]
                    or [("—", "—")],
                    "summary_vals": [nfs, nfs, "1", link_v, join_metric],
                    "values_column": ["key:preview"],
                    "defined": nj > 0,
                }
            else:
                dn = _ne_detail_name_cfg()
                ne_hit_names = (
                    _cap_list_capped([Path(x).name for x in paths], MAX_VALUE_ROWS_DEFAULT)
                    if paths
                    else col_paths
                )
                slots[0] = {
                    "short": "ファイル検索",
                    "editor_lines": name_extract_debug_slot_editor_lines(s0, p, dn, 0),
                    "details": [],
                    "summary_vals": [nfs, "-", "-", "-", "-"],
                    "values_column": ne_hit_names,
                }
                slots[1] = {
                    "short": "抜取り文字",
                    "editor_lines": name_extract_debug_slot_editor_lines(s0, p, dn, 1),
                    "details": [],
                    "summary_vals": [nfs, "-", "-", "-", "-"],
                    "values_column": ["（名前抽出プレビュー）"],
                }
                slots[2] = {
                    "short": "関連付け",
                    "editor_lines": name_extract_debug_slot_editor_lines(s0, p, dn, 2),
                    "details": [],
                    "summary_vals": [nfs, "-", "-", "-", "-"],
                    "values_column": ["（結合パス・プレビュー）"],
                }
                slots[3] = None
                slots[4] = None
            if typ == "cell":
                summ = " | ".join(
                    x
                    for x in (fp, str(p.get("sheet_rule") or ""), str(s0.get("cell_ref") or ""))
                    if x
                )
                skind = "cell"
            else:
                path_one = str(p.get("path_item") or "").strip()
                stx = str(s0.get("search_text") or "").strip()
                summ = " | ".join(x for x in (fp, stx, path_one) if x)
                skind = "name_extract"
            out.append(
                {
                    "title": title,
                    "summary": (summ or title)[:120],
                    "slots": slots,
                    "source_kind": skind,
                    "source": copy.deepcopy(s0),
                }
            )
    return out


def _master_debug_csv_precache_progress_hook(
    batch_hook: Callable[..., None] | None,
    *,
    cancel_check: Callable[..., None] | None = None,
) -> Callable[[str], None] | None:
    """compute_batch 用 progress_hook を CSV precache 文言 (str) 向けにラップ。"""
    if batch_hook is None and cancel_check is None:
        return None

    def _hook(msg: str) -> None:
        if cancel_check is not None:
            cancel_check(force=True)
        if batch_hook is not None:
            try:
                batch_hook(4, str(msg))
            except Exception:
                pass

    return _hook


def _precache_csv_for_master_debug_extract(
    file_path: str,
    *,
    progress_hook: Callable[[str], None] | None = None,
) -> None:
    """xlsx_workbook_scope 内: 本番一括と同様に CSV を先読み（lazy cache と結果同等）。"""
    if not str(file_path).lower().endswith(".csv"):
        return
    from svc.svc_data_agg_extract import precache_csv_matrix_for_file  # noqa: WPS433

    precache_csv_matrix_for_file(file_path, progress_hook=progress_hook)


def build_master_items_live(
    items: list[dict[str, Any]],
    scan_paths: list[str] | None,
    max_rows: int,
    *,
    preload_values: bool = True,
) -> list[dict[str, Any]]:
    """
    メインから起動したマスタ項目デバッグ用。項目×ソースごとにシナリオ行を生成し、
    本番と同一の extract_item_bundle で主値列を埋める（上限 max_rows）。
    """
    paths = [str(p).strip() for p in (scan_paths or []) if str(p).strip()]
    try:
        from svc.svc_data_agg import filter_file_paths_for_master_preview
        from svc.svc_data_agg_extract import (
            extract_item_bundle,
            file_paths_for_source_extract,
            xlsx_workbook_scope,
        )
    except Exception:
        extract_item_bundle = None  # type: ignore[misc, assignment]
        filter_file_paths_for_master_preview = None  # type: ignore[misc, assignment]
        file_paths_for_source_extract = None  # type: ignore[misc, assignment]
        xlsx_workbook_scope = None  # type: ignore[misc, assignment]
    if filter_file_paths_for_master_preview is not None and paths:
        paths = list(filter_file_paths_for_master_preview(paths, items))

    def _empty_slot(msg: str) -> dict[str, Any]:
        return {
            "short": "—",
            "editor_lines": [msg],
            "details": [],
            "summary_vals": ["-", "-", "-", "-", "-"],
            "values_column": ["（なし）"],
            "values_prod": ["（なし）"],
        }

    out: list[dict[str, Any]] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        title = str(item.get("name") or item.get("id") or "項目").strip() or "項目"
        sources = item.get("sources") or []
        scenarios: list[dict[str, Any]] = []
        if not sources:
            scenarios.append(
                {
                    "title": "（シナリオなし）",
                    "cond_sum": "—",
                    "slot": {**_empty_slot("取得シナリオが登録されていません"), "defined": False},
                }
            )
        else:
            for si, src in enumerate(sources):
                if not isinstance(src, dict):
                    continue
                one = {**item, "sources": [copy.deepcopy(src)]}
                rows = build_debug_scenarios_from_items([one], paths)
                raw_slots = rows[0].get("slots") if rows else []
                slots: list[Any] = raw_slots if isinstance(raw_slots, list) else []
                summary_vals = ["-", "-", "-", "-", "-"]
                for j in range(4, -1, -1):
                    if j < len(slots) and slots[j] is not None:
                        sv = slots[j].get("summary_vals")
                        if isinstance(sv, list) and len(sv) >= 5:
                            summary_vals = [str(x) for x in sv[:5]]
                        break
                values_column = ["（抽出プレビュー）"]
                if len(slots) > 2 and slots[2] is not None:
                    vc = slots[2].get("values_column")
                    if isinstance(vc, list) and vc:
                        values_column = [str(x) for x in vc]

                col_vals: list[str] = []
                if preload_values:
                    if extract_item_bundle is None:
                        col_vals = ["（svc_data_agg_extract を読み込めませんでした）"]
                    elif not paths:
                        col_vals = ["（検出ファイルがありません。メインで基準フォルダ・拡張子を確認するかデバッグを開き直してください。）"]
                    else:
                        item_id = str(item.get("id") or title)
                        if file_paths_for_source_extract is not None:
                            paths_list = file_paths_for_source_extract(paths, src)
                        else:
                            paths_list = list(paths)
                        for fp in paths_list:
                            if len(col_vals) >= max_rows:
                                break
                            with xlsx_workbook_scope():  # type: ignore[misc]
                                try:
                                    _precache_csv_for_master_debug_extract(fp)
                                    jp_hdr = str(one.get("name") or one.get("id") or "").strip()
                                    b = extract_item_bundle(
                                        fp,
                                        one,
                                        item_id=item_id,
                                        cell_positions={},
                                        join_path_header=jp_hdr or None,
                                    )
                                except Exception:
                                    b = {"primary_values": []}
                            _append_extract_primaries_to_col(
                                col_vals,
                                b.get("primary_values"),
                                max_rows=max_rows,
                            )
                        if not col_vals:
                            col_vals = ["（該当する主値がありません）"]
                else:
                    col_vals = list(values_column)

                capped = _cap_list_capped(col_vals, max_rows)
                editor_lines: list[str] = []
                for s in slots:
                    if s and isinstance(s, dict) and s.get("editor_lines"):
                        editor_lines.extend(s["editor_lines"])
                if not editor_lines:
                    editor_lines = ["（要約）"]
                short_txt = title
                if rows:
                    short_txt = str(rows[0].get("summary") or title)[:120]
                sn_src = str(src.get("scenario_name") or "").strip()
                stitle = sn_src if sn_src else ("シナリオ %d" % (si + 1))
                cond_sum = str(src.get("type") or "cell")
                scenarios.append(
                    {
                        "title": stitle,
                        "cond_sum": cond_sum,
                        "source": copy.deepcopy(src),
                        "slot": {
                            "short": short_txt,
                            "editor_lines": editor_lines[:24],
                            "details": [("種別", cond_sum)],
                            "summary_vals": summary_vals,
                            "values_column": capped,
                            "values_prod": list(capped),
                        },
                    }
                )
            if not scenarios:
                scenarios.append(
                    {
                        "title": "（シナリオ不正）",
                        "cond_sum": "—",
                        "slot": {**_empty_slot("有効な取得シナリオがありません"), "defined": False},
                    }
                )

        summ = "%d シナリオ" % len(sources) if sources else "シナリオなし"
        out.append({"title": title, "summary": summ, "scenarios": scenarios})

    if not out:
        return [
            {
                "title": "（項目なし）",
                "summary": "—",
                "scenarios": [
                    {
                        "title": "—",
                        "cond_sum": "—",
                        "slot": _empty_slot("メインで項目を定義してください"),
                    }
                ],
            }
        ]
    return out


def _empty_debug_scenarios_data() -> list[dict[str, Any]]:
    """live_items なし時の空プレースホルダ（本番にデモシナリオは同梱しない）。"""
    return [
        {
            "title": "（デバッグ対象なし）",
            "summary": "メインで項目とシナリオを定義するか、シナリオ編集から開いてください。",
            "slots": [None] * 5,
            "source_kind": "cell",
            "source": None,
        }
    ]


def _empty_debug_master_items() -> list[dict[str, Any]]:
    """マスタモードで live_items が無いとき。"""
    return [
        {
            "title": "（項目なし）",
            "summary": "メインで項目を定義してからデバッグを開いてください。",
            "scenarios": [
                {
                    "title": "—",
                    "cond_sum": "",
                    "slot": None,
                }
            ],
        }
    ]


# マスタ項目ステップ実行の進捗（10 段）。done/total の total は本配列の長さ。
_MASTER_DEBUG_PROGRESS_PHASES: tuple[str, ...] = (
    "準備",
    "サマリー更新",
    "値取得",
    "読込開始",
    "ファイル読込",
    "行まとめ",
    "結合準備",
    "結合索引",
    "結合照合",
    "一覧組立",
)
_MASTER_DEBUG_PROGRESS_PHASE_DONE = "完了"
# compute_batch の progress_hook sub 4〜7 が担う UI 段（読込開始〜一覧組立）
_MASTER_DEBUG_BATCH_UI_PHASE_COUNT = 7

# シナリオデバッグ: 連携(3)／結合(4)フェーズで、検出ファイルがこの件数以上のときだけファイル単位進捗を表示
# （SCREENS.DEBUG.SCENARIO_PROGRESS_MIN_FILES で上書き可。0＝閾値なしで常に進捗フック）
SCENARIO_PROGRESS_MIN_FILES = 15
_SCENARIO_PROGRESS_PHASE_MSGPRIMARY = "主キーを取得中"
_SCENARIO_PROGRESS_PHASE_MSGLINK = "連携キーを取得中"
_SCENARIO_PROGRESS_PHASE_MSGJOIN = "結合キーを取得中"
_SCENARIO_PROGRESS_PHASE_MSGSTAGE = "ネットワークからマウント中"


def _scenario_progress_min_files_from_cfg(cfg: dict[str, Any]) -> int:
    raw = (cfg or {}).get("SCENARIO_PROGRESS_MIN_FILES")
    if raw is None:
        return SCENARIO_PROGRESS_MIN_FILES
    try:
        return max(0, int(raw))
    except (TypeError, ValueError):
        return SCENARIO_PROGRESS_MIN_FILES


class _ValueGridNoElideDelegate(QStyledItemDelegate):
    """結果一覧で長い主キー等が … 省略されないよう style option を固定する。"""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.phase_start_cols: frozenset[int] = frozenset()

    def initStyleOption(self, option: QStyleOptionViewItem, index) -> None:  # type: ignore[override]
        super().initStyleOption(option, index)
        option.textElideMode = Qt.TextElideMode.ElideNone

    def paint(self, painter, option, index) -> None:  # type: ignore[override]
        super().paint(painter, option, index)
        _paint_value_grid_phase_divider(
            painter, option.rect, index.column(), self.phase_start_cols
        )


class _ValueGridPhaseHeader(QHeaderView):
    """シナリオ結果一覧の列見出しにフェーズ境界の太線を描く。"""

    def __init__(self, orientation: Qt.Orientation, parent: QWidget | None = None) -> None:
        super().__init__(orientation, parent)
        self.phase_start_cols: frozenset[int] = frozenset()
        self.setSectionsClickable(True)
        self.setHighlightSections(True)

    def paintSection(self, painter, rect, logicalIndex) -> None:  # type: ignore[override]
        super().paintSection(painter, rect, logicalIndex)
        _paint_value_grid_phase_divider(
            painter, rect, int(logicalIndex), self.phase_start_cols
        )


class _ScenarioLinkPrefetchBridge(QObject):
    finished = Signal(int, int)


class _DebugCondTreeWidget(QTreeWidget):
    """行高が大きいとき開閉インジケータを上寄せで描画する。"""

    def drawBranches(self, painter: QPainter, rect, index) -> None:  # type: ignore[override]
        fm = self.fontMetrics()
        slack = max(0, rect.height() - fm.height())
        dy = slack // 2
        if dy > 0:
            painter.translate(0, -dy)
        try:
            super().drawBranches(painter, rect, index)
        finally:
            if dy > 0:
                painter.translate(0, dy)


class DataAggDebugDialog(QDialog):
    """デバッグ実行 UI（プレビュー専用）。設定キーは SCREENS.DEBUG。"""

    def __init__(
        self,
        parent: QWidget | None = None,
        debug_cfg: dict[str, Any] | None = None,
        live_items: list[dict[str, Any]] | None = None,
        scan_paths: list[str] | None = None,
        fixed_mode: int | None = None,
        scenario_for_dry_run: dict[str, Any] | None = None,
        scan_root: str | None = None,
    ) -> None:
        super().__init__(parent)
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_AlwaysShowToolTips, True)
        except Exception:
            pass
        self._cfg = dict(debug_cfg or {})
        self._scenario_progress_min_files = _scenario_progress_min_files_from_cfg(
            self._cfg
        )
        self._fixed_mode = fixed_mode
        self._mode = int(fixed_mode) if fixed_mode is not None else 0
        self._master_items_override: list[dict[str, Any]] | None = None
        self._live_items: list[dict[str, Any]] = list(live_items or [])
        self._debug_scan_paths: list[str] = list(scan_paths or [])
        self._scan_root: str | None = (str(scan_root).strip() or None) if scan_root else None
        self._scenario_bundle_caches: dict[int, dict[str, dict[str, Any]]] = {}
        from svc.svc_data_agg_debug_run import ScenarioDebugStageSession  # noqa: WPS433

        self._scenario_stage_session = ScenarioDebugStageSession()
        self._scenario_link_prefetch_gen: int = 0
        self._scenario_link_prefetch_thread: threading.Thread | None = None
        self._scenario_link_prefetch_cancel = threading.Event()
        self._scenario_link_prefetch_cancel.set()
        self._scenario_link_prefetch_bridge = _ScenarioLinkPrefetchBridge(self)
        self._scenario_link_prefetch_bridge.finished.connect(
            self._on_scenario_link_prefetch_finished
        )
        self._scenario_for_dry_run: dict[str, Any] | None = (
            copy.deepcopy(scenario_for_dry_run) if scenario_for_dry_run else None
        )
        if live_items:
            self._scenarios_data = build_debug_scenarios_from_items(
                live_items, scan_paths
            )
        else:
            self._scenarios_data = _empty_debug_scenarios_data()
        if self._mode == 1:
            if live_items:
                self._master_items_override = build_master_items_live(
                    live_items, scan_paths, self._max_value_rows(), preload_values=False
                )
            else:
                self._master_items_override = _empty_debug_master_items()
        wc = self._cfg.get("WINDOW") or {}
        w = int(wc.get("DEFAULT_WIDTH") or 860)
        h = int(wc.get("DEFAULT_HEIGHT") or 800)
        if w > 0 and h > 0:
            self.resize(w, h)
        self._sc_idx = 0
        self._mi_idx = 0
        self._phase_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._last_master_active_count = 0
        self._master_exec_armed = False
        self._master_global_row_idx: int = 0
        self._master_run_progress_active: bool = False
        self._master_progress_window_title: str = ""
        self._master_progress_pct_floor: int = 0
        self._master_batch_hook_last_fi: int = 1
        self._master_batch_hook_last_nf: int = 1
        # 項目ループ中もファイル開始時の [CCH]/[LOC] を 1 行目に残す
        self._master_batch_hook_last_cache_mark: str = ""
        self._master_batch_hook_mark_fi: int = 0
        # マスタプレビュー（シナリオ単位）で最後だけ確定表示を走らせるための一時フラグ
        self._master_force_finalize_preview: bool = False
        # mpv: 現在列のマージ値（gcell）を表示するか。進行中は False、最終反映時のみ True。
        self._mpv_show_merged_current: bool = False
        # Run ボタン内の複数ステップ実行中（途中のマージ反映抑止）
        self._master_step_loop_busy: bool = False

        self._summary_rows: list[list[str]] = []
        self._value_cols: list[list[str]] = []
        self._value_col_tooltips: list[list[str | None]] = []
        self._value_col_spans: list[tuple[int, int]] = []
        self._mpv_join_table_active: bool = False
        self._mpv_join_table_ncols: int = 0
        self._master_sparse_notice_shown: bool = False
        # mpv: 描画直前の table_rows 同期バッファ（extract マージ結果は入れない）。
        self._mpv_grid: list[list[Any]] | None = None
        # mpv: 直近の非空 table_rows（compute 待ち中に前回表示を保持）
        self._mpv_last_valid_table_rows: list[list[Any]] = []
        # mpv: 見出し3指標（ファイル数・読込行数・走査上限到達）
        self._mpv_last_stats_files_read: int = 0
        self._mpv_last_stats_read_rows: int = 0
        self._mpv_last_stats_scan_cap_hit: bool = False
        # mpv: 結合 compute 中カウンタ（res_hint「計算中」）
        self._mpv_join_compute_busy: int = 0
        # mpv: extract_item_bundle 主値列のキャッシュ（(mi_idx, scenario_si)）
        self._mpv_extract_cache: dict[tuple[int, int], list[str]] = {}
        # mpv: step 中に確定した現在列値のキャッシュ（軽量 progress 合成用）
        self._mpv_colvals_cache: dict[tuple[int, int], list[str]] = {}
        # mpv 描画: 進捗マスク用の部分 compute_batch 行キャッシュ (key, rows)
        self._mpv_progress_rows_cache: tuple[Any, list[list[Any]]] | None = None
        # mpv 描画: n_pick（現在項目に取り込むソース本数）単位の progress 行キャッシュ
        self._mpv_progress_rows_step_cache: dict[tuple[Any, ...], list[list[Any]]] = {}
        # バックグラウンド先読みを打ち切る世代（キャッシュクリア・項目変更時に進める）
        self._mpv_prefetch_cancel_gen: int = 0
        # メインと先読みで compute_batch を同時に走らせない（GIL 奪い合い防止）
        self._mpv_prog_compute_lock = threading.Lock()
        # 先読みジョブは最大1件（古い依頼は捨てて最新だけ）
        self._mpv_prefetch_q: queue.Queue[Any] = queue.Queue(maxsize=1)
        self._mpv_prefetch_worker_started = False
        self._mpv_prefetch_worker_lock = threading.Lock()
        self._mpv_prefetch_debounce_ms = 35
        self._mpv_prefetch_debounce_timer = QTimer(self)
        self._mpv_prefetch_debounce_timer.setSingleShot(True)
        self._mpv_prefetch_debounce_timer.timeout.connect(
            self._mpv_prefetch_debounced_fire
        )
        # single_slot warmup 投入済みで step キャッシュ未反映（先読み進行中）
        self._mpv_single_slot_prefetch_pending_sk: tuple[Any, ...] | None = None
        # mpv: 項目単位の workbook 共有キャッシュ（スレッド横断で frame dict を bind）
        self._mpv_item_wb_frame: dict[str, Any] | None = None
        self._mpv_item_wb_mi: int | None = None
        self._mpv_item_wb_pending_close: list[dict[str, Any]] = []
        self._mpv_wb_worker_thread: threading.Thread | None = None
        # 進捗 hook の file_index 解決用（compute/extract 実パス。全 scan_paths とは限らない）
        self._mpv_progress_hook_paths: list[str] | None = None
        # mpv 描画: 項目ごとの progress 行キャッシュ（step_idx, rows）
        self._mpv_progress_rows_by_mi: dict[int, tuple[int, list[list[Any]]]] = {}
        # 直近まで compute 済みの項目（実行可能シナリオなし項目へ移ったときの prog フォールバック用）
        self._last_master_completed_mi_idx: int | None = None
        # mpv: 完了項目列の凍結（行キー __norm_path + __iter_index）。次項目 compute の再走査を抑える。
        self._mpv_frozen_snapshots: dict[int, dict[str, Any]] = {}
        # 項目ごとの progress 行数ピーク（段階キャッシュの中途半端な行数を弾く）
        self._mpv_progress_row_peak_by_mi: dict[int, int] = {}
        # 連続する結合項目間で join_search プールを再利用（次項目の再走査を抑える）
        self._mpv_join_search_pool_seed: list[dict[str, Any]] | None = None
        self._mpv_join_search_pool_seed_paths_count: int = -1
        self._mpv_join_pool_by_mi: dict[int, list[dict[str, Any]]] = {}
        # 積み上げ join seed 用: 項目ごとの table_rows 行に対応する参照元ファイルパス
        self._mpv_row_file_paths_by_mi: dict[int, list[str]] = {}
        # 全項目完了後の結果一覧（file_path + iter_index 順の本番同等行）
        self._mpv_final_table_rows: list[list[Any]] | None = None
        # 描画時に「現在列」として扱う項目 index（フォールバック表示整合用）
        self._mpv_display_mi_idx: int | None = None
        # シナリオなし項目の直後に「実行あり」項目へ入ったとき、入場直後の value グリッド再構築を
        # その項目の全ステップ完了時（離脱直前）まで遅延する。対象 mi（到着先の index）。
        self._mpv_deferred_value_grid_mi: int | None = None
        # 連続実行中は列幅 content_fit を最後に 1 回だけ行う
        self._mpv_column_fit_pending: bool = False
        self._mpv_final_grid_applied: bool = False
        self._value_grid_header_programmatic: bool = False
        self._value_grid_programmatic_gen: int = 0
        self._value_grid_user_resized: bool = False
        self._value_grid_saved_widths: list[int] | None = None
        self._value_grid_structure_key: tuple[str, ...] | None = None
        self._active_slot_indices: list[int] = []
        self._summary_phase_labels: list[str] = []

        self._scenario_snapshots: dict[int, dict[str, Any]] = {}

        self._continuous_busy: bool = False
        self._continuous_steps_left: int = 0
        self._continuous_initial_steps: int = 0
        self._continuous_was_full_master: bool = False
        self._master_full_continuous_allowed: bool = True
        self._mpv_pending_trunc_warns: list[dict[str, Any]] = []
        self._master_step_pass_complete: bool = False
        self._master_snapshot_browse_after_cancel: bool = False
        self._master_cancel_mi: int = 0
        self._master_cancel_step: int = 0
        self._run_progress_dlg: Any | None = None
        self._run_progress_path: Path | None = None
        self._run_cancel_path: Path | None = None
        self._master_cancel_check: Callable[..., None] | None = None
        self._master_cancel_scope_cm: Any | None = None
        self._master_step_cancelled: bool = False
        self._master_continuous_cancel_requested: bool = False
        self._master_cancel_event = threading.Event()
        self._master_cancel_pump_timer: QTimer | None = None
        self._master_step_exec_depth: int = 0
        self._master_cooperative_abort_retries: int = 0
        self._master_abort_in_progress: bool = False
        self._run_progress_seq: int = 0
        self._debug_progress_locked: bool = False
        self._master_item_snapshots: dict[int, dict[str, Any]] = {}
        self._master_item_snapshot_done: set[int] = set()
        self._master_step_snapshots: dict[tuple[int, int], dict[str, Any]] = {}
        # マスタ実行時間（秒）: (mi_idx, step_idx) / mi_idx 合計 / 連続実行全体
        self._master_step_elapsed_sec: dict[tuple[int, int], float] = {}
        self._master_item_elapsed_sec: dict[int, float] = {}
        self._master_step_timing_t0: float | None = None
        self._master_continuous_run_t0: float | None = None

        self._build_ui()
        try:
            from ui_qt.ui_common import apply_window_config

            ph = 0
            p: QWidget | None = self.parentWidget()
            while p is not None:
                if hasattr(p, "_parent_hwnd"):
                    ph = int(getattr(p, "_parent_hwnd", 0) or 0)
                    break
                p = p.parentWidget()
            win_cfg = dict(self._cfg.get("WINDOW") or {})
            win_cfg["CENTER_ON_EXCEL"] = False
            apply_window_config(
                self, {"WINDOW": win_cfg}, ph, "DEBUG"
            )
        except Exception:
            pass
        self._apply_mode()
        self._refresh_all()

    def showEvent(self, event: Any) -> None:
        """親（メイン／シナリオ編集）の中央付近に重ねて表示する。"""
        super().showEvent(event)
        self._center_on_parent_widget()
        ph = 0
        pwalk: QWidget | None = self.parentWidget()
        while pwalk is not None:
            if hasattr(pwalk, "_parent_hwnd"):
                ph = int(getattr(pwalk, "_parent_hwnd", 0) or 0)
                break
            pwalk = pwalk.parentWidget()
        if ph:
            try:
                from ui_qt.ui_common import ensure_front

                eh = ph

                def _front0() -> None:
                    try:
                        ensure_front(self, eh)
                    except Exception:
                        pass

                QTimer.singleShot(0, _front0)
            except Exception:
                pass
        QTimer.singleShot(0, self._center_on_parent_widget)
        QTimer.singleShot(160, self._center_on_parent_widget)

    def _center_on_parent_widget(self) -> None:
        pw = self.parentWidget()
        if pw is None:
            return
        pr = pw.frameGeometry()
        gr = self.frameGeometry()
        x = pr.x() + (pr.width() - gr.width()) // 2
        y = pr.y() + (pr.height() - gr.height()) // 2
        self.move(x, y)

    def _d(self, key: str, default: str) -> str:
        s = _normalize_message_newlines(str(self._cfg.get(key) or default).strip())
        if key.endswith("_HTML"):
            return s.replace("\n", "<br/>")
        return s

    def _tip(self, key: str, default: str = "") -> str:
        """DEBUG 用ツールチップ文言（プレーン）。JSON の TIP_* を想定。"""
        raw = str(self._cfg.get(key) or default).strip()
        return _normalize_tooltip_text(raw) if raw else ""

    def _set_tip(self, w: QWidget | None, key: str, default: str = "") -> None:
        if w is None:
            return
        t = self._tip(key, default)
        if not t and default:
            t = _normalize_tooltip_text(str(default).strip())
        if t:
            set_widget_tooltip(w, t)

    def _debug_window_title_for_mode(self) -> str:
        if self._mode == 0:
            key = "TITLE_SCENARIO"
        else:
            key = "TITLE_MASTER"
        t = str(self._d(key, "") or "").strip()
        if not t:
            t = str(self._d("TITLE", "データ集約 デバッグ") or "").strip()
        return t or "データ集約 デバッグ"

    def _update_debug_window_title(self) -> None:
        self.setWindowTitle(self._debug_window_title_for_mode())

    def _btn_run_all_label(self) -> str:
        if self._mode == 0:
            key = "BTN_RUN_ALL_SCENARIO"
        else:
            key = "BTN_RUN_ALL_MASTER"
        t = str(self._cfg.get(key) or "").strip()
        if not t:
            t = str(self._d("BTN_RUN_ALL", "項目実行") or "").strip()
        return t or "項目実行"

    def _apply_static_debug_tooltips(self) -> None:
        """ボタン・表・タブなどモードに依存しないツールチップ。"""
        self._set_tip(
            self._lbl_mode,
            "TIP_LABEL_MODE",
            "プレビュー対象の切り替えです。シナリオ単位とマスタ全項目（ステップ）のどちらで動かすかを表します。",
        )
        self._set_tip(
            self.mode_combo,
            "TIP_MODE_COMBO",
            "モードを選ぶと左の一覧・実行ボタンの意味が切り替わります。シナリオ編集／メインから固定起動されたときは表示のみです。",
        )
        mf = getattr(self, "_mode_fixed_label", None)
        if mf is not None:
            self._set_tip(
                mf,
                "TIP_MODE_COMBO",
                "モードを選ぶと左の一覧・実行ボタンの意味が切り替わります。シナリオ編集／メインから固定起動されたときは表示のみです。",
            )
        self._set_tip(
            self.hint,
            "TIP_MAIN_HINT",
            "現在のモードでの操作説明です。内容はシナリオ／マスタで切り替わります。",
        )
        self._set_tip(
            self.left_title,
            "TIP_LEFT_TITLE",
            "左ペインの見出しです。登録シナリオ一覧かマスタ項目一覧かを示します。",
        )
        self._set_tip(
            self.left_table,
            "TIP_LEFT_TABLE",
            "シナリオまたはマスタ項目の一覧です。行を選ぶと条件ステップ・右ペインが連動します。",
        )
        self._set_tip(
            self._lbl_steps,
            "TIP_LABEL_PHASE_STEPS",
            "選択中の行に紐づく条件ステップ（フェーズ）です。番号は結果サマリ列・ログと対応します。",
        )
        self._set_tip(
            self.left_steps,
            "TIP_LEFT_STEPS",
            "フェーズ番号と識別名です。行を選ぶとログや結果の見る位置の目安になります。",
        )
        self._set_tip(
            self.btn_prev,
            "TIP_BTN_PREV",
            "一覧で前のシナリオまたはマスタ項目へ移動します。",
        )
        self._set_tip(
            self.btn_next,
            "TIP_BTN_NEXT",
            "一覧で次のシナリオまたはマスタ項目へ移動します。",
        )
        self._set_tip(
            self.tabs,
            "TIP_TABS_WIDGET",
            "条件・結果・ログの表示を切り替えます。",
        )
        try:
            self.tabs.setTabToolTip(
                0,
                self._tip(
                    "TIP_TAB_CONDITIONS",
                    "ファイル検索・主キー・連携・結合などの条件をツリー表示します。",
                ),
            )
            self.tabs.setTabToolTip(
                1,
                self._tip(
                    "TIP_TAB_RESULTS",
                    "サマリ指標と抽出結果のプレビュー表を表示します（マスタへは書き込みません）。",
                ),
            )
            self.tabs.setTabToolTip(
                2,
                self._tip(
                    "TIP_TAB_LOG",
                    "実行の要約・EVENT 行などのログを表示します（先頭が最新）。",
                ),
            )
        except Exception:
            pass
        self._set_tip(
            self.btn_run_all_master,
            "TIP_RUN_ALL_MASTER_ITEMS",
            "全項目を一巡するまで自動実行します（シナリオ未登録の項目は 1 ステップでスキップ）。",
        )
        self._set_tip(
            self.btn_clear_res,
            "TIP_BTN_CLEAR_RESULTS",
            "結果サマリと結果一覧グリッドの表示をクリアします。",
        )
        self._set_tip(
            self.btn_clear_log,
            "TIP_BTN_CLEAR_LOG",
            "ログ本文を空にします。",
        )
        self._set_tip(
            self.btn_cancel,
            "TIP_BTN_CANCEL",
            "デバッグウィンドウを閉じます。",
        )
        self._set_tip(
            self.cond_hint,
            "TIP_COND_HINT",
            "条件ツリーの見方（シナリオ／マスタ）の説明です。",
        )
        self._set_tip(
            self.cond_tree,
            "TIP_COND_TREE",
            "条件の階層です。親子行を展開して要約を確認します。",
        )
        self._set_tip(
            self.master_cond_tree,
            "TIP_COND_TREE",
            "条件の階層です。親子行を展開して要約を確認します。",
        )
        self._set_tip(
            self.res_hint,
            "TIP_RES_HINT",
            "結果タブのサマリ・一覧の見方の説明です。",
        )
        self._refresh_fold_button_tooltip()
        self._set_tip(
            self.summary_table,
            "TIP_SUMMARY_TABLE",
            "フェーズごとの指標（ファイル件数・主キー件数など）のサマリです。",
        )
        self._set_tip(
            self.values_title,
            "TIP_VALUES_TITLE",
            "結果一覧グリッドの見出しです。表示中の項目名などが含まれます。",
        )
        self._set_tip(
            self.value_grid,
            "TIP_VALUE_GRID",
            "抽出・結合後の値プレビューです。セルにマウスを乗せると全文のツールチップが出ます。",
        )
        self._set_tip(
            self.log_intro,
            "TIP_LOG_INTRO",
            "ログ欄に出る行の意味（EVENT など）の説明です。",
        )
        self._set_tip(
            self.log,
            "TIP_LOG_TEXT",
            "読み取り専用の実行ログです。クリアは「ログクリア」ボタンで行います。",
        )
        self._set_tip(
            self._debug_main_splitter,
            "TIP_SPLITTER_MAIN",
            "左ペイン（一覧・フェーズステップ）と右ペイン（タブ）の幅をドラッグで調整します。",
        )
        self._set_tip(
            self._left_col_split,
            "TIP_SPLITTER_LEFT_COLUMN",
            "上段の一覧テーブルと下段のフェーズステップ表の高さを調整します。",
        )
        self._set_tip(
            self._debug_left_panel,
            "TIP_LEFT_PANE_CONTAINER",
            "一覧・フェーズステップ・前後移動をまとめた左側の作業領域です。",
        )
        self._set_tip(
            self._debug_right_panel,
            "TIP_RIGHT_PANE_CONTAINER",
            "タブと実行操作をまとめた右側の作業領域です。",
        )
        self._set_tip(
            self._debug_steps_host,
            "TIP_STEPS_HOST",
            "フェーズステップ見出しと一覧テーブルを含む領域です。",
        )
        self._set_tip(
            self.cond_stack,
            "TIP_COND_STACK",
            "シナリオ用／マスタ用の条件ツリーを切り替えて表示します。",
        )
        self._set_tip(
            self.tab_cond,
            "TIP_TAB_PAGE_CONDITIONS",
            "条件タブのページ本体です（ツリーが表示されます）。",
        )
        self._set_tip(
            self.tab_res,
            "TIP_TAB_PAGE_RESULTS",
            "結果タブのページ本体です（サマリ・一覧グリッドが表示されます）。",
        )
        self._set_tip(
            self.tab_log,
            "TIP_TAB_PAGE_LOG",
            "ログタブのページ本体です（説明ラベルとログテキストが表示されます）。",
        )
        self._set_tip(
            self.left_table.horizontalHeader(),
            "TIP_TABLE_HEADER_LEFT_LIST",
            "シナリオ名またはマスタ項目名の列見出しです。",
        )
        self._set_tip(
            self.left_steps.horizontalHeader(),
            "TIP_TABLE_HEADER_LEFT_STEPS",
            "フェーズ番号と識別ラベルの列見出しです。",
        )
        self._set_tip(
            self.summary_table.horizontalHeader(),
            "TIP_TABLE_HEADER_SUMMARY",
            "結果サマリ表の列見出しです（指標名はモードにより変化します）。",
        )
        self._set_tip(
            self.value_grid.horizontalHeader(),
            "TIP_TABLE_HEADER_VALUE_GRID",
            "結果一覧の列見出しです（フェーズ列は実行に応じて増えます）。",
        )
        self._set_tip(
            self.cond_tree.header(),
            "TIP_COND_TREE_HEADER",
            "番号・項目・要約の各列見出しです。",
        )
        self._set_tip(
            self.master_cond_tree.header(),
            "TIP_COND_TREE_HEADER",
            "番号・項目・要約の各列見出しです。",
        )
        self._set_tip(
            self.summary_table.verticalHeader(),
            "TIP_TABLE_VERTICAL_SUMMARY",
            "サマリ表の行見出し（フェーズや指標の行）です。",
        )
        self._set_tip(
            self.value_grid.verticalHeader(),
            "TIP_TABLE_VERTICAL_VALUE_GRID",
            "結果一覧の行番号です。",
        )
        sw = getattr(self, "_summary_table_wrap", None)
        self._set_tip(
            sw,
            "TIP_SUMMARY_WRAP",
            "折りたたみ可能な結果サマリ表のコンテナです。上のボタンで表示／非表示を切り替えます。",
        )

    def _apply_mode_dependent_tooltips(self) -> None:
        """実行ボタン等、モードで意味が変わるツールチップ。"""
        run_all_fb = self._tip(
            "TIP_RUN_ALL",
            "選択されている項目に登録されている、シナリオを一括実行します。",
        )
        if self._mode == 0:
            self._set_tip(
                self.btn_run,
                "TIP_BTN_RUN_SCENARIO",
                "選択中シナリオのフェーズを順に実行し、プレビューを更新します。",
            )
            self._set_tip(self.btn_run_all, "TIP_RUN_ALL_SCENARIO", run_all_fb)
        else:
            self._set_tip(
                self.btn_run,
                "TIP_BTN_RUN_MASTER",
                "選択中マスタ項目のステップを実行し、プレビューを更新します。",
            )
            self._set_tip(self.btn_run_all, "TIP_RUN_ALL_MASTER", run_all_fb)

    def _refresh_fold_button_tooltip(self) -> None:
        btn = getattr(self, "btn_summary_fold", None)
        if btn is None:
            return
        if btn.isChecked():
            self._set_tip(
                btn,
                "TIP_BTN_SUMMARY_EXPANDED",
                "クリックで結果サマリ表を折りたたみます。",
            )
        else:
            self._set_tip(
                btn,
                "TIP_BTN_SUMMARY_COLLAPSED",
                "クリックで結果サマリ表を展開して表示します。",
            )

    def _window_int(self, key: str, default: int) -> int:
        wc = self._cfg.get("WINDOW") or {}
        try:
            return int(wc.get(key, default))
        except (TypeError, ValueError):
            return default

    def _wrap_desc_label(
        self,
        w: QLabel,
        *,
        h_policy: QSizePolicy.Policy = QSizePolicy.Policy.Expanding,
    ) -> None:
        """説明文を親幅に合わせて折り返し、ウィンドウ最小幅を抑える。"""
        w.setWordWrap(True)
        w.setMinimumWidth(0)
        w.setSizePolicy(h_policy, QSizePolicy.Policy.Preferred)

    def _table_hdr_min_section(self) -> int:
        v = self._window_int("TABLE_MIN_HEADER_SECTION", 16)
        return max(12, min(80, v))

    def _window_int_list(self, key: str) -> list[int]:
        wc = self._cfg.get("WINDOW") or {}
        raw = wc.get(key)
        if not isinstance(raw, list):
            return []
        out: list[int] = []
        for x in raw:
            try:
                out.append(int(x))
            except (TypeError, ValueError):
                pass
        return out

    def _column_lo_hi_simple(
        self, col: int, mins: list[int], maxs: list[int], default_lo: int
    ) -> tuple[int, int]:
        lo = default_lo
        if col < len(mins) and int(mins[col]) > 0:
            lo = int(mins[col])
        hi = 3200
        if col < len(maxs) and int(maxs[col]) > 0:
            hi = max(lo, int(maxs[col]))
        return lo, hi

    def _header_global_floor(self, mins: list[int], global_default: int) -> int:
        pos: list[int] = []
        for m in mins:
            try:
                v = int(m)
                if v > 0:
                    pos.append(v)
            except (TypeError, ValueError):
                pass
        return min(pos) if pos else global_default

    def _summary_col_default_lo(self, col: int) -> int:
        if col == 0:
            return 56
        return self._table_hdr_min_section()

    def _value_grid_col_default_lo(self, col: int) -> int:
        _ = col
        return max(48, self._table_hdr_min_section())

    def _fit_qtable_columns(
        self,
        table: QTableWidget,
        *,
        mins_key: str,
        maxs_key: str,
        skip_last_stretch: bool,
        global_default: int,
        default_lo_for_col: Callable[[int], int],
    ) -> None:
        """表示内容に合わせて列幅を調整し、JSON の min/max でクリップする。"""
        n = table.columnCount()
        if n <= 0:
            return
        mins = self._window_int_list(mins_key)
        maxs = self._window_int_list(maxs_key)
        hdr = table.horizontalHeader()
        hdr.setMinimumSectionSize(self._header_global_floor(mins, global_default))
        last = n - 1 if skip_last_stretch and n > 1 else n
        for c in range(last):
            dlo = default_lo_for_col(c)
            lo, hi = self._column_lo_hi_simple(c, mins, maxs, dlo)
            table.resizeColumnToContents(c)
            w = table.columnWidth(c)
            table.setColumnWidth(c, max(lo, min(w, hi)))

    def _bump_qtable_columns_for_header_labels(
        self,
        table: QTableWidget,
        *,
        mins_key: str,
        maxs_key: str,
        default_lo_for_col: Callable[[int], int],
        hpad: int = 28,
    ) -> None:
        """列見出しが … 省略されないよう、見出し文字列幅まで列幅を広げる（JSON min/max 内）。"""
        n = table.columnCount()
        if n <= 0:
            return
        hdr = table.horizontalHeader()
        fm = QFontMetrics(hdr.font())
        mins = self._window_int_list(mins_key)
        maxs = self._window_int_list(maxs_key)
        for c in range(n):
            hi = table.horizontalHeaderItem(c)
            text = hi.text() if hi is not None else ""
            lines = [
                ln for ln in str(text).replace("\r", "").split("\n") if str(ln).strip()
            ]
            need = 0
            for line in lines:
                need = max(need, fm.horizontalAdvance(line))
            need += hpad
            if need <= 0:
                continue
            dlo = default_lo_for_col(c)
            lo, hiw = self._column_lo_hi_simple(c, mins, maxs, dlo)
            w = table.columnWidth(c)
            table.setColumnWidth(c, max(w, min(max(need, lo), hiw)))

    def _apply_cond_tree_initial_column_widths(self, tw: QTreeWidget) -> None:
        mins = self._window_int_list("COND_TREE_COL_MIN_WIDTHS")
        maxs = self._window_int_list("COND_TREE_COL_MAX_WIDTHS")
        defaults = [56, 120, 96]
        for c in range(tw.columnCount()):
            dlo = defaults[c] if c < len(defaults) else self._table_hdr_min_section()
            lo, _ = self._column_lo_hi_simple(c, mins, maxs, dlo)
            tw.setColumnWidth(c, lo)

    def _fit_cond_tree_columns(self, tw: QTreeWidget) -> None:
        n = tw.columnCount()
        if n <= 0:
            return
        mins = self._window_int_list("COND_TREE_COL_MIN_WIDTHS")
        maxs = self._window_int_list("COND_TREE_COL_MAX_WIDTHS")
        defaults = [56, 120, 96]
        hdr = tw.header()
        hdr.setMinimumSectionSize(
            self._header_global_floor(mins, self._table_hdr_min_section())
        )
        for c in range(n):
            dlo = defaults[c] if c < len(defaults) else self._table_hdr_min_section()
            lo, hi = self._column_lo_hi_simple(c, mins, maxs, dlo)
            tw.resizeColumnToContents(c)
            w = tw.columnWidth(c)
            tw.setColumnWidth(c, max(lo, min(w, hi)))

    def _fit_summary_table_columns(self) -> None:
        self._fit_qtable_columns(
            self.summary_table,
            mins_key="SUMMARY_TABLE_COL_MIN_WIDTHS",
            maxs_key="SUMMARY_TABLE_COL_MAX_WIDTHS",
            skip_last_stretch=False,
            global_default=self._table_hdr_min_section(),
            default_lo_for_col=self._summary_col_default_lo,
        )
        self._bump_qtable_columns_for_header_labels(
            self.summary_table,
            mins_key="SUMMARY_TABLE_COL_MIN_WIDTHS",
            maxs_key="SUMMARY_TABLE_COL_MAX_WIDTHS",
            default_lo_for_col=self._summary_col_default_lo,
        )

    def _value_grid_note_structure(self, headers: list[str]) -> None:
        """列見出しセットが変わったら手動列幅状態を捨てる。"""
        key = tuple(str(h) for h in headers)
        if not key:
            return
        if self._value_grid_structure_key != key:
            self._value_grid_structure_key = key
            self._value_grid_user_resized = False
            self._value_grid_saved_widths = None

    def _value_grid_current_headers_key(self) -> tuple[str, ...]:
        n = self.value_grid.columnCount()
        out: list[str] = []
        for c in range(n):
            hi = self.value_grid.horizontalHeaderItem(c)
            out.append(str(hi.text()) if hi is not None else "")
        return tuple(out)

    def _on_value_grid_section_resized(self, _logical: int, _old: int, _new: int) -> None:
        if self._value_grid_header_programmatic:
            return
        n = self.value_grid.columnCount()
        if n <= 0:
            return
        self._value_grid_user_resized = True
        self._value_grid_saved_widths = [self.value_grid.columnWidth(c) for c in range(n)]

    def _value_grid_schedule_programmatic_end(self, gen: int) -> None:
        """setColumnWidth / resizeColumnToContents 後に遅延する sectionResized を無視するため、
        programmatic 解除を次イベントループへずらす。連続フィットは _value_grid_programmatic_gen で無効化。"""
        def _clear() -> None:
            try:
                if int(getattr(self, "_value_grid_programmatic_gen", 0)) != int(gen):
                    return
                self._value_grid_header_programmatic = False
            except RuntimeError:
                pass

        QTimer.singleShot(0, _clear)

    def _diag_log_value_grid_columns(
        self,
        branch: str,
        n: int,
        *,
        sample_lines: list[str] | None = None,
    ) -> None:
        """結果一覧の列幅まわりを DATA_AGG_DIAG に出す（マスタ列重なり等の調査用）。"""
        try:
            if n <= 0:
                return
            widths = [int(self.value_grid.columnWidth(c)) for c in range(n)]
            vw = -1
            try:
                vp = self.value_grid.viewport()
                if vp is not None:
                    vw = int(vp.width())
            except Exception:
                pass
            key_eq = self._value_grid_current_headers_key() == self._value_grid_structure_key
            saved = self._value_grid_saved_widths or []
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] value_grid_columns branch=%s mode=%s ncols=%s "
                "user_resized=%s saved_len=%s key_match=%s sum_widths=%s viewport_w=%s "
                "widths_first12=%s",
                branch,
                int(self._mode),
                n,
                bool(self._value_grid_user_resized),
                len(saved),
                key_eq,
                sum(widths),
                vw,
                widths[:12],
            )
            if n > 12:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] value_grid_columns branch=%s widths_tail=%s",
                    branch,
                    widths[-4:],
                )
            if sample_lines:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] value_grid_columns branch=%s col_samples %s",
                    branch,
                    " | ".join(sample_lines),
                )
        except Exception:
            pass

    def _fit_value_grid_columns(self) -> None:
        """全列を内容幅で調整。同一構成でユーザーが列幅を変えていれば復元する（最終列 Stretch なし）。"""
        n = self.value_grid.columnCount()
        if n <= 0:
            return
        cur_key = self._value_grid_current_headers_key()
        if self._value_grid_structure_key is None and cur_key:
            self._value_grid_structure_key = cur_key
        saved_w = self._value_grid_saved_widths
        restore = (
            self._value_grid_user_resized
            and saved_w is not None
            and len(saved_w) == n
            and cur_key == self._value_grid_structure_key
        )
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] value_grid_fit_enter mode=%s ncols=%s restore_path=%s "
                "user_resized=%s saved_len=%s key_eq=%s",
                int(self._mode),
                n,
                restore,
                bool(self._value_grid_user_resized),
                len(saved_w or []),
                cur_key == self._value_grid_structure_key,
            )
        except Exception:
            pass
        if restore and saved_w is not None:
            hdr = self.value_grid.horizontalHeader()
            floor = hdr.minimumSectionSize()
            self._value_grid_programmatic_gen += 1
            gen = self._value_grid_programmatic_gen
            self._value_grid_header_programmatic = True
            try:
                hdr.blockSignals(True)
                for c, w in enumerate(saved_w):
                    if c < n:
                        self.value_grid.setColumnWidth(c, max(int(floor), int(w)))
            finally:
                hdr.blockSignals(False)
            self._bump_qtable_columns_for_header_labels(
                self.value_grid,
                mins_key="VALUE_GRID_COL_MIN_WIDTHS",
                maxs_key="VALUE_GRID_COL_MAX_WIDTHS",
                default_lo_for_col=self._value_grid_col_default_lo,
            )
            self._diag_log_value_grid_columns("restore_saved", n)
            self._value_grid_schedule_programmatic_end(gen)
            return
        mins = self._window_int_list("VALUE_GRID_COL_MIN_WIDTHS")
        maxs = self._window_int_list("VALUE_GRID_COL_MAX_WIDTHS")
        hdr = self.value_grid.horizontalHeader()
        hdr.setMinimumSectionSize(self._header_global_floor(mins, self._table_hdr_min_section()))
        sample_lines: list[str] = []
        self._value_grid_programmatic_gen += 1
        gen = self._value_grid_programmatic_gen
        self._value_grid_header_programmatic = True
        try:
            hdr.blockSignals(True)
            for c in range(n):
                dlo = self._value_grid_col_default_lo(c)
                lo, hi = self._column_lo_hi_simple(c, mins, maxs, dlo)
                self.value_grid.resizeColumnToContents(c)
                w_raw = int(self.value_grid.columnWidth(c))
                w_fin = max(lo, min(w_raw, hi))
                self.value_grid.setColumnWidth(c, w_fin)
                if c < 3 or (n > 11 and 10 <= c <= 11):
                    sample_lines.append(
                        "c%d lo=%s hi=%s raw=%s fin=%s" % (c, lo, hi, w_raw, w_fin)
                    )
        finally:
            hdr.blockSignals(False)
        self._bump_qtable_columns_for_header_labels(
            self.value_grid,
            mins_key="VALUE_GRID_COL_MIN_WIDTHS",
            maxs_key="VALUE_GRID_COL_MAX_WIDTHS",
            default_lo_for_col=self._value_grid_col_default_lo,
        )
        self._diag_log_value_grid_columns(
            "content_fit", n, sample_lines=sample_lines or None
        )
        self._value_grid_schedule_programmatic_end(gen)

    def _cfg_debug_int(self, key: str, default: int, *legacy_keys: str) -> int:
        for k in (key, *legacy_keys):
            raw = self._cfg.get(k)
            if raw is None:
                continue
            try:
                return int(raw)
            except (TypeError, ValueError):
                continue
        return int(default)

    def _max_value_rows(self) -> int:
        return self._cfg_debug_int(
            "SCENARIO_DEBUG_VALUE_ROWS",
            MAX_VALUE_ROWS_DEFAULT,
            "MAX_VALUE_ROWS",
        )

    def _master_preview_display_rows(self) -> int:
        return max(
            1,
            self._cfg_debug_int(
                "MASTER_DEBUG_DISPLAY_ROWS",
                100,
                "MASTER_PREVIEW_DISPLAY_ROWS",
            ),
        )

    def _master_debug_join_max_files(self) -> int:
        """結合項目のファイル読込上限。0 で無制限。"""
        return self._cfg_debug_int("MASTER_DEBUG_JOIN_MAX_FILES", 20)

    def _master_debug_max_files(self) -> int:
        """非結合項目のファイル読込上限。0 で無制限。"""
        return self._cfg_debug_int("MASTER_DEBUG_MAX_FILES", 20)

    def _mpv_begin_join_compute(self) -> None:
        self._mpv_join_compute_busy = int(getattr(self, "_mpv_join_compute_busy", 0)) + 1
        self._update_values_title_master()

    def _mpv_end_join_compute(self) -> None:
        self._mpv_join_compute_busy = max(
            0, int(getattr(self, "_mpv_join_compute_busy", 0)) - 1
        )
        self._update_values_title_master()

    def _master_values_title_rows_suffix(self) -> str:
        """結果一覧見出し文末: N/M 行数 or 計算中。"""
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_debug_format_row_count,
            master_debug_values_title_rows_busy_text,
            master_debug_values_title_rows_stats_fmt,
            master_debug_values_title_scan_cap_suffix,
        )

        if int(getattr(self, "_mpv_join_compute_busy", 0) or 0) > 0:
            return self._d(
                "VALUES_TITLE_MASTER_ROWS_BUSY",
                master_debug_values_title_rows_busy_text(),
            )
        disp_n = len(self._mpv_last_valid_table_rows or [])
        disp_m = self._master_preview_display_rows()
        files_n = int(getattr(self, "_mpv_last_stats_files_read", 0) or 0)
        read_n = int(getattr(self, "_mpv_last_stats_read_rows", 0) or 0)
        fmt = self._d(
            "VALUES_TITLE_MASTER_ROWS_FMT",
            master_debug_values_title_rows_stats_fmt(),
        )
        suffix = fmt % (
            master_debug_format_row_count(disp_n),
            master_debug_format_row_count(disp_m),
            master_debug_format_row_count(files_n),
            master_debug_format_row_count(read_n),
        )
        if bool(getattr(self, "_mpv_last_stats_scan_cap_hit", False)):
            suffix += self._d(
                "VALUES_TITLE_MASTER_SCAN_CAP_SUFFIX",
                master_debug_values_title_scan_cap_suffix(),
            )
        return suffix

    def _update_master_res_hint(self) -> None:
        """マスタモード結果タブ res_hint（結果サマリ説明のみ）。"""
        if self._mode != 1:
            return
        self.res_hint.setText(
            self._d("RES_HINT_MASTER_HTML", "")
            or (
                "<b>結果サマリ</b>：全ステップを積み重ね。"
                " <b>結果一覧</b>：結合後テーブルの最大表示行でプレビューします。"
            )
        )

    def _mpv_item_stats_for_snapshot(self) -> dict[str, Any]:
        return {
            "display_n": len(self._mpv_last_valid_table_rows or []),
            "display_m": self._master_preview_display_rows(),
            "files_read": int(getattr(self, "_mpv_last_stats_files_read", 0) or 0),
            "read_rows": int(getattr(self, "_mpv_last_stats_read_rows", 0) or 0),
            "scan_cap_hit": bool(getattr(self, "_mpv_last_stats_scan_cap_hit", False)),
        }

    def _mpv_apply_item_stats_snapshot(self, stats: Any) -> None:
        if not isinstance(stats, dict):
            return
        self._mpv_last_stats_files_read = int(stats.get("files_read") or 0)
        self._mpv_last_stats_read_rows = int(stats.get("read_rows") or 0)
        self._mpv_last_stats_scan_cap_hit = bool(stats.get("scan_cap_hit"))

    def _mpv_note_item_stats(self, scen: dict[str, Any]) -> None:
        """直近 compute の3指標を記録（結果一覧見出し用）。"""
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_read_rows_for_display,
        )

        dd = scen.get("__debug_diag")
        join_item = self._mpv_current_item_has_join_defs(int(self._mi_idx))
        if isinstance(dd, dict):
            self._mpv_last_stats_files_read = int(
                dd.get("master_preview_stats_files_read") or 0
            )
            scan = int(dd.get("master_preview_stats_scan_rows") or 0)
            join_ref = int(dd.get("master_preview_stats_join_ref_rows") or 0)
            self._mpv_last_stats_read_rows = master_preview_join_read_rows_for_display(
                scan_rows=scan,
                join_ref_rows=join_ref,
                join_item=bool(join_item),
            )
            self._mpv_last_stats_scan_cap_hit = bool(
                dd.get("master_preview_stats_scan_cap_hit")
            )
            return
        pool = self._mpv_join_pool_by_mi.get(int(self._mi_idx))
        if pool and join_item:
            self._mpv_last_stats_read_rows = len(pool)

