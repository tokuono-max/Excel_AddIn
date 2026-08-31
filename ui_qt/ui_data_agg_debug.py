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
_SCENARIO_PROGRESS_PHASE_MSGLINK = "連携キーを取得中"
_SCENARIO_PROGRESS_PHASE_MSGJOIN = "結合キーを取得中"


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
        # 項目ループ中もファイル開始時の [C]/[F] を 1 行目に残す
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

    def _max_phase_slots(self) -> int:
        try:
            return int(self._cfg.get("MAX_PHASE_SLOTS") or MAX_PHASE_SLOTS_DEFAULT)
        except (TypeError, ValueError):
            return MAX_PHASE_SLOTS_DEFAULT

    def _scenario_source_kind(self) -> str:
        if self._mode != 0 or not self._scenarios_data:
            return "cell"
        return str(self._current_scenario().get("source_kind") or "cell").strip().lower()

    def _cond_keys(self) -> list[str]:
        if self._scenario_source_kind() == "name_extract":
            ck = self._cfg.get("COND_KEYS_NAME_EXTRACT")
            if isinstance(ck, list) and ck:
                return [str(x) for x in ck]
            return list(COND_KEYS_NAME_EXTRACT_DEFAULT)
        ck = self._cfg.get("COND_KEYS")
        if isinstance(ck, list) and ck:
            return [str(x) for x in ck]
        return list(COND_KEYS_DEFAULT)

    def _summary_metric_labels(self) -> list[str]:
        def _lbl(x: Any) -> str:
            return _normalize_message_newlines(str(x).strip())

        if self._scenario_source_kind() == "name_extract":
            h = self._cfg.get("SUMMARY_HEADERS_NAME_EXTRACT_SHORT")
            if isinstance(h, list) and h:
                return [_lbl(x) for x in h]
            return [_lbl(x) for x in SUMMARY_HEADERS_NAME_EXTRACT_SHORT_DEFAULT]
        h = self._cfg.get("SUMMARY_HEADERS_SHORT")
        if isinstance(h, list) and h:
            return [_lbl(x) for x in h]
        h = self._cfg.get("SUMMARY_HEADERS")
        if isinstance(h, list) and h:
            return [_lbl(x) for x in h]
        return [_lbl(x) for x in SUMMARY_HEADERS_2L_DEFAULT]

    def _summary_log_headers(self) -> list[str]:
        def _lbl(x: Any) -> str:
            return _normalize_message_newlines(str(x).strip())

        n = len(self._summary_metric_labels())
        if self._scenario_source_kind() == "name_extract":
            h = self._cfg.get("SUMMARY_HEADERS_NAME_EXTRACT_LOG_SHORT")
            if isinstance(h, list) and len(h) == n:
                return [_lbl(x) for x in h]
            h = self._cfg.get("SUMMARY_HEADERS_NAME_EXTRACT_LOG")
            if isinstance(h, list) and len(h) == n:
                return [_lbl(x) for x in h]
            return [_lbl(x) for x in SUMMARY_HEADERS_NAME_EXTRACT_LOG_SHORT_DEFAULT[:n]]
        h = self._cfg.get("SUMMARY_HEADERS_LOG_SHORT")
        if isinstance(h, list) and len(h) == n:
            return [_lbl(x) for x in h]
        h = self._cfg.get("SUMMARY_HEADERS_LOG")
        if isinstance(h, list) and len(h) == n:
            return [_lbl(x) for x in h]
        return [_lbl(x) for x in self._summary_metric_labels()]

    def _summary_cols_display(self) -> list[str]:
        labels = self._summary_metric_labels()
        return [self._d("SUMMARY_PHASE_COLUMN", "シナリオ名")] + [
            labels[i] for i in self._visible_metric_indices()
        ]

    def _n_metrics(self) -> int:
        return len(self._summary_metric_labels())

    def _visible_metric_indices(self) -> list[int]:
        n = self._n_metrics()
        if self._mode != 0:
            return list(range(n))
        if self._scenario_source_kind() == "name_extract":
            return list(range(min(4, n)))
        out: list[int] = []
        for gi in self._active_slot_indices:
            if 0 <= gi < n and gi not in out:
                out.append(gi)
        return out

    def _summary_vals_for_display(self, vals: list[str]) -> list[str]:
        out: list[str] = []
        for i in self._visible_metric_indices():
            out.append(str(vals[i]) if i < len(vals) else "-")
        return out

    def _ensure_summary_table_columns(self) -> None:
        cols = self._summary_cols_display()
        self.summary_table.setColumnCount(len(cols))
        self.summary_table.setHorizontalHeaderLabels(cols)
        self._apply_summary_table_resize_modes()
        self._style_results_table_header_rows()

    def _summary_table_min_height(self) -> int:
        try:
            h = int(self._cfg.get("SUMMARY_TABLE_MIN_HEIGHT") or 100)
        except (TypeError, ValueError):
            h = 100
        return max(72, h)

    def _value_grid_min_height(self) -> int:
        try:
            h = int(self._cfg.get("VALUE_GRID_MIN_HEIGHT") or 240)
        except (TypeError, ValueError):
            h = 240
        return max(120, h)

    def _debug_parent_hwnd(self) -> int:
        ph = 0
        p: QWidget | None = self.parentWidget()
        while p is not None:
            if hasattr(p, "_parent_hwnd"):
                ph = int(getattr(p, "_parent_hwnd", 0) or 0)
                break
            p = p.parentWidget()
        return ph

    def _process_events_light(self) -> None:
        """進捗表示中はユーザ入力（キャンセルボタン等）も処理する。"""
        if bool(
            getattr(self, "_master_run_progress_active", False)
            or getattr(self, "_debug_progress_locked", False)
            or getattr(self, "_continuous_busy", False)
        ):
            self._process_events_for_master_cancel()
            return
        try:
            QApplication.processEvents(
                QEventLoop.ProcessEventsFlag.ExcludeUserInputEvents
            )
        except Exception:
            try:
                QApplication.processEvents()
            except Exception:
                pass

    def _process_events_for_master_cancel(self) -> None:
        """進捗ダイアログのキャンセル等を届けるため、ユーザ入力込みでイベントを処理する。"""
        app = QApplication.instance()
        if app is None:
            return
        try:
            flags = QEventLoop.ProcessEventsFlag.AllEvents
        except AttributeError:
            flags = QEventLoop.AllEvents  # type: ignore[attr-defined]
        try:
            app.processEvents(flags, 50)
            app.processEvents(flags, 50)
        except TypeError:
            try:
                app.processEvents()
            except Exception:
                pass
        except Exception:
            try:
                app.processEvents()
            except Exception:
                pass

    def _show_master_run_cancel_notice(self, *, continuous: bool = False) -> None:
        """進捗キャンセル後の短い通知（ログ欄とは別）。"""
        try:
            from ui_qt.ui_common import show_info_notice

            title = self._d(
                "DIALOG_MASTER_RUN_CANCEL_TITLE",
                self._d("DIALOG_RUN_ALL_DONE_TITLE", "データ集約 デバッグ"),
            )
            if continuous:
                body = self._d(
                    "MSG_MASTER_RUN_CANCEL_CONTINUOUS_NOTICE",
                    "連続実行を中止しました。",
                )
            else:
                body = self._d(
                    "MSG_MASTER_RUN_CANCEL_NOTICE",
                    "実行を中止しました。",
                )
            show_info_notice(self, title, _normalize_message_newlines(body))
        except Exception:
            pass

    def _activate_master_run_progress_for_cancel(self, dlg: Any | None = None) -> None:
        """進捗表示直後: デバッグではなく進捗を前面化し、キャンセルボタンへフォーカスする。"""
        if dlg is None:
            dlg = getattr(self, "_run_progress_dlg", None)
        if dlg is None:
            return
        try:
            dlg.raise_()
            dlg.activateWindow()
            btn = getattr(dlg, "_btn_cancel", None)
            if btn is not None and btn.isEnabled() and btn.isVisible():
                btn.setFocus(Qt.FocusReason.OtherFocusReason)
        except Exception:
            pass

    def _ensure_master_cancel_pump_timer(self) -> None:
        t = getattr(self, "_master_cancel_pump_timer", None)
        if t is None:
            t = QTimer(self)
            t.setInterval(30)
            t.timeout.connect(self._master_cancel_pump_tick)  # type: ignore[attr-defined]
            self._master_cancel_pump_timer = t
        if not t.isActive():
            t.start()

    def _stop_master_cancel_pump_timer(self) -> None:
        t = getattr(self, "_master_cancel_pump_timer", None)
        if t is not None:
            try:
                t.stop()
            except Exception:
                pass

    def _master_cancel_pump_tick(self) -> None:
        if not (
            getattr(self, "_master_run_progress_active", False)
            or getattr(self, "_debug_progress_locked", False)
            or getattr(self, "_continuous_busy", False)
        ):
            self._stop_master_cancel_pump_timer()
            return
        self._process_events_for_master_cancel()
        from svc.data_agg_cancel import DataAggCancelled, cancel_requested  # noqa: WPS433

        p = self._run_cancel_path
        if p is not None and (
            self._master_cancel_event.is_set() or cancel_requested(p)
        ):
            self._master_note_cancel_requested()
            self._master_sync_run_progress_after_cancel_signal()
            self._master_schedule_cooperative_abort()
            return
        try:
            self._master_poll_cancel(force=True)
        except DataAggCancelled:
            self._master_note_cancel_requested()
            self._master_schedule_cooperative_abort()
            return
        self._master_sync_run_progress_after_cancel_signal()

    def _trigger_master_run_cancel(self, *, source: str = "ui") -> None:
        """協調キャンセル pickle / Event を立て、連続実行も止める。"""
        self._ensure_master_run_cancel()
        p = self._run_cancel_path
        if p is not None:
            try:
                ipc_file.write_pickle(p, {"cancel": True, "v": 1})
            except Exception:
                pass
        pp = self._run_progress_path
        if pp is not None:
            try:
                from svc.data_agg_cancel import write_progress_cancel_status  # noqa: WPS433

                write_progress_cancel_status(pp)
            except Exception:
                pass
        try:
            self._master_cancel_event.set()
        except Exception:
            pass
        self._master_note_cancel_requested()
        try:
            from core.core_log import get_logger  # noqa: WPS433

            get_logger(__name__).info(
                "[DATA_AGG] master_debug cancel triggered source=%s path=%s",
                str(source or "ui"),
                str(p or ""),
            )
        except Exception:
            pass
        try:
            QTimer.singleShot(0, self._master_sync_run_progress_after_cancel_signal)
        except Exception:
            pass
        self._master_schedule_cooperative_abort()

    def _master_schedule_cooperative_abort(self) -> None:
        """協調キャンセル着信後、ステップ中断をイベントループ境界で試みる。"""
        if not self._master_cancel_pending():
            return
        try:
            QTimer.singleShot(0, self._master_cooperative_abort_tick)
        except Exception:
            pass

    def _master_cooperative_abort_tick(self) -> None:
        """Qt タイマー／ボタン経由のキャンセルで abort まで到達させる（再入安全）。"""
        if not self._master_cancel_pending():
            self._master_cooperative_abort_retries = 0
            return
        if getattr(self, "_master_abort_in_progress", False):
            return
        depth = int(getattr(self, "_master_step_exec_depth", 0) or 0)
        if depth > 0:
            retries = int(getattr(self, "_master_cooperative_abort_retries", 0) or 0)
            if retries < 60:
                self._master_cooperative_abort_retries = retries + 1
                try:
                    QTimer.singleShot(50, self._master_cooperative_abort_tick)
                except Exception:
                    pass
                return
        self._master_cooperative_abort_retries = 0
        if depth <= 0 and not (
            getattr(self, "_continuous_busy", False)
            or getattr(self, "_debug_progress_locked", False)
            or getattr(self, "_master_run_progress_active", False)
        ):
            return
        self._master_abort_step_after_cancel()

    def _master_sync_run_progress_after_cancel_signal(self) -> None:
        """進捗 pickle の CANCEL でダイアログだけ閉じたとき、親側の進捗フラグを同期する。"""
        dlg = getattr(self, "_run_progress_dlg", None)
        try:
            closed = dlg is None
            if dlg is not None:
                try:
                    closed = not dlg.isVisible()
                except Exception:
                    closed = True
            if closed and getattr(self, "_master_run_progress_active", False):
                self._master_run_progress_active = False
                self._run_progress_dlg = None
        except Exception:
            pass

    def _set_debug_progress_locked(self, locked: bool) -> None:
        if self._mode != 1:
            return
        self._debug_progress_locked = locked
        self._refresh_master_nav_lock_state()

    def _reset_master_cancel_state(self) -> None:
        """キャンセル協調状態を初期化（再実行前・キャンセル完了後）。"""
        self._master_step_cancelled = False
        self._master_continuous_cancel_requested = False
        self._master_cooperative_abort_retries = 0
        try:
            self._master_cancel_event.clear()
        except Exception:
            pass
        self._stop_master_cancel_pump_timer()
        self._clear_master_run_cancel()

    def _ensure_master_run_cancel(self) -> None:
        """マスタ実行中の協調キャンセル pickle（進捗ダイアログのキャンセルボタン用）。"""
        from svc.data_agg_cancel import (  # noqa: WPS433
            cancel_request_path_data_agg_master_debug,
            make_cancel_check,
            reset_cancel_path,
        )

        if self._run_cancel_path is not None and self._master_cancel_check is not None:
            try:
                reset_cancel_path(self._run_cancel_path)
            except Exception:
                pass
            try:
                self._master_cancel_event.clear()
            except Exception:
                pass
            return
        tok = datetime.now().strftime("%Y%m%d%H%M%S%f")
        p = cancel_request_path_data_agg_master_debug(ipc_file.get_ipc_root(), token=tok)
        reset_cancel_path(p)
        self._run_cancel_path = p
        try:
            self._master_cancel_event.clear()
        except Exception:
            pass
        base_check = make_cancel_check(p, min_interval_sec=0.0)
        evt = self._master_cancel_event
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        def _combined_cancel_check(*, force: bool = False) -> None:
            if evt.is_set():
                raise DataAggCancelled()
            if base_check is not None:
                base_check(force=force)

        self._master_cancel_check = _combined_cancel_check
        if getattr(self, "_master_cancel_scope_cm", None) is None:
            from svc.data_agg_cancel import batch_cancel_scope  # noqa: WPS433

            cm = batch_cancel_scope(self._master_cancel_check)
            cm.__enter__()
            self._master_cancel_scope_cm = cm

    def _clear_master_run_cancel(self) -> None:
        from svc.data_agg_cancel import reset_cancel_path  # noqa: WPS433

        cm = getattr(self, "_master_cancel_scope_cm", None)
        if cm is not None:
            try:
                cm.__exit__(None, None, None)
            except Exception:
                pass
            self._master_cancel_scope_cm = None
        p = self._run_cancel_path
        self._run_cancel_path = None
        self._master_cancel_check = None
        try:
            self._master_cancel_event.clear()
        except Exception:
            pass
        if p is not None:
            try:
                reset_cancel_path(p)
            except Exception:
                pass

    def _master_run_cancel_check(self) -> Callable[..., None] | None:
        return getattr(self, "_master_cancel_check", None)

    def _master_poll_cancel(self, *, force: bool = False) -> None:
        if getattr(self, "_master_run_progress_active", False):
            self._process_events_for_master_cancel()
        chk = self._master_run_cancel_check()
        if chk is not None:
            chk(force=force)

    def _on_master_progress_cancel_clicked(self) -> None:
        """進捗ダイアログのキャンセル押下（pickle は ProgressDialog 側で書込済み）。"""
        self._trigger_master_run_cancel(source="progress_dialog")

    def _master_on_ui_thread(self) -> bool:
        app = QApplication.instance()
        if app is None:
            return threading.current_thread() is threading.main_thread()
        try:
            return QThread.currentThread() is app.thread()
        except Exception:
            return threading.current_thread() is threading.main_thread()

    def _master_should_offthread_compute(self, probe_caller: str) -> bool:
        if str(probe_caller or "") == "mpv_progress_prefetch":
            return False
        return self._master_on_ui_thread()

    def _master_bridge_progress_hook(
        self,
        real_hook: Callable[..., None] | None,
        hook_q: queue.SimpleQueue,
    ) -> Callable[..., None] | None:
        if real_hook is None:
            return None

        def _bridged(*args: Any, **kwargs: Any) -> None:
            hook_q.put((args, kwargs))

        return _bridged

    def _master_drain_progress_hook_queue(
        self,
        hook_q: queue.SimpleQueue,
        real_hook: Callable[..., None] | None,
    ) -> None:
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        while True:
            try:
                args, kwargs = hook_q.get_nowait()
            except queue.Empty:
                break
            if real_hook is None:
                continue
            try:
                real_hook(*args, **kwargs)
            except DataAggCancelled:
                self._master_note_cancel_requested()
                raise

    def _master_run_blocking_with_ui_pump(
        self,
        fn: Callable[[], Any],
        *,
        progress_hook: Callable[..., None] | None = None,
        hook_q: queue.SimpleQueue | None = None,
    ) -> Any:
        """compute 中もメインスレッドで Qt イベントを処理し、キャンセルボタンを有効にする。"""
        import contextvars

        drain_q = hook_q if hook_q is not None else queue.SimpleQueue()
        result_box: list[Any] = []
        exc_box: list[BaseException] = []

        def worker() -> None:
            try:
                result_box.append(fn())
            except BaseException as exc:
                exc_box.append(exc)

        ctx = contextvars.copy_context()
        t = threading.Thread(
            target=lambda: ctx.run(worker),
            daemon=True,
            name="master_dbg_compute",
        )
        self._mpv_wb_worker_thread = t
        t.start()
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        cancelled_early = False
        try:
            while t.is_alive():
                try:
                    self._master_drain_progress_hook_queue(drain_q, progress_hook)
                except DataAggCancelled:
                    self._master_note_cancel_requested()
                    cancelled_early = True
                    break
                self._process_events_for_master_cancel()
                if self._master_cancel_pending():
                    cancelled_early = True
                    break
                try:
                    self._master_poll_cancel(force=True)
                except DataAggCancelled:
                    self._master_note_cancel_requested()
                    cancelled_early = True
                    break
                time.sleep(0.005)
            if cancelled_early or self._master_cancel_pending():
                # 項目 workbook キャッシュ破棄前にワーカーが参照を手放すのを待つ
                try:
                    t.join(timeout=8.0)
                except Exception:
                    pass
                raise DataAggCancelled()
            try:
                self._master_drain_progress_hook_queue(drain_q, progress_hook)
            except DataAggCancelled:
                self._master_note_cancel_requested()
                try:
                    t.join(timeout=8.0)
                except Exception:
                    pass
                raise
            if self._master_cancel_pending():
                try:
                    t.join(timeout=8.0)
                except Exception:
                    pass
                raise DataAggCancelled()
            if exc_box:
                exc = exc_box[0]
                if isinstance(exc, DataAggCancelled):
                    self._master_note_cancel_requested()
                raise exc
            return result_box[0]
        finally:
            self._mpv_clear_wb_worker_if_done()

    def _master_note_cancel_requested(self) -> None:
        """協調キャンセル検知時: 連続実行も止める。"""
        self._master_step_cancelled = True
        self._master_continuous_cancel_requested = True
        self._bump_mpv_prefetch_cancel()

    def _master_cancel_pending(self) -> bool:
        return bool(
            getattr(self, "_master_step_cancelled", False)
            or getattr(self, "_master_continuous_cancel_requested", False)
        )

    def _master_raise_if_cancelled(self, *, force_poll: bool = False) -> None:
        """キャンセル保留または協調 pickle を検知したら即 DataAggCancelled。"""
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        if self._master_cancel_pending():
            raise DataAggCancelled()
        if force_poll or getattr(self, "_master_run_progress_active", False) or getattr(
            self, "_continuous_busy", False
        ):
            self._master_poll_cancel(force=True)

    def _master_abort_step_after_cancel(self) -> tuple[bool, bool]:
        """ステップ実行をキャンセルで打ち切り（途中 upsert はロールバックしない）。"""
        if getattr(self, "_master_abort_in_progress", False):
            return False, False
        self._master_abort_in_progress = True
        try:
            self._master_step_cancelled = False
            self._master_cooperative_abort_retries = 0
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] master_step_abort reason=cancel mi_idx=%s step_idx=%s",
                    self._mi_idx,
                    self._master_step_idx,
                )
            except Exception:
                pass
            try:
                _diag_logger.info(
                    "[DATA_AGG] master_step_abort reason=cancel mi_idx=%s step_idx=%s",
                    self._mi_idx,
                    self._master_step_idx,
                )
            except Exception:
                pass
            self._log_append(
                self._d("MSG_MASTER_RUN_CANCEL", "（実行を中止しました）")
            )
            self._show_master_run_cancel_notice(
                continuous=bool(getattr(self, "_continuous_busy", False))
            )
            try:
                from core.core_cursor import progress_dialog_wait_cursor_off  # noqa: WPS433

                progress_dialog_wait_cursor_off(cancel_reason="master_debug_cancel")
            except Exception:
                pass
            self._close_run_progress(cancelled=True)
            if getattr(self, "_continuous_busy", False):
                self._finish_continuous_run()
            else:
                self._stop_master_cancel_pump_timer()
                self._clear_master_run_cancel()
            self._reset_master_cancel_state()
            self._enter_master_snapshot_browse_after_cancel()
            return False, False
        finally:
            self._master_abort_in_progress = False
            self._master_step_exec_depth_leave()
            self._master_step_timing_t0 = None

    def _master_return_if_step_cancelled(self) -> tuple[bool, bool] | None:
        if self._master_cancel_pending():
            return self._master_abort_step_after_cancel()
        return None

    def _mpv_store_row_file_paths_for_mi(
        self,
        mi_idx: int,
        rows: list[list[Any]],
        *,
        iteration_contexts: list[dict[str, Any]] | None = None,
    ) -> None:
        """積み上げ join seed 用に、table_rows 行ごとの参照元ファイルパスを保持する。"""
        from svc.data_agg_master_preview import (  # noqa: WPS433
            table_row_file_paths_for_stacked_seed,
        )

        n = len(rows or [])
        if n <= 0:
            return
        stored: list[str] | None = None
        if iteration_contexts:
            fps = [
                str(ctx.get("file_path") or "").strip()
                for ctx in iteration_contexts
                if isinstance(ctx, dict)
            ]
            if len(fps) >= n and any(fps[:n]):
                stored = fps[:n]
        paths = table_row_file_paths_for_stacked_seed(
            self._mpv_preview_headers(),
            list(rows),
            scan_paths=list(self._debug_scan_paths or []),
            stored_row_paths=stored,
        )
        if not (paths and any(paths)):
            return
        from svc.data_agg_master_preview import row_file_paths_real_count  # noqa: WPS433

        prev = (getattr(self, "_mpv_row_file_paths_by_mi", {}) or {}).get(int(mi_idx))
        if isinstance(prev, list) and len(prev) == len(paths):
            if row_file_paths_real_count(prev) > row_file_paths_real_count(paths):
                return
        self._mpv_row_file_paths_by_mi[int(mi_idx)] = list(paths)

    def _mpv_find_row_file_paths_for_stacked_seed(
        self,
        n_rows: int,
        *,
        start_mi: int | None = None,
    ) -> list[str] | None:
        """積み上げ seed: 前段 mi から synthetic でない行パスを遡って取得。"""
        from svc.data_agg_master_preview import row_file_paths_real_count  # noqa: WPS433

        n = max(0, int(n_rows))
        if n <= 0:
            return None
        begin = int(start_mi) if start_mi is not None else len(
            (getattr(self, "_mpv_row_file_paths_by_mi", {}) or {})
        )
        for mi in range(begin, -1, -1):
            prev = (getattr(self, "_mpv_row_file_paths_by_mi", {}) or {}).get(mi)
            if not isinstance(prev, list) or len(prev) < n:
                continue
            chunk = [str(p) for p in prev[:n]]
            if row_file_paths_real_count(chunk) > 0:
                return chunk
        return None

    def _mpv_row_file_paths_for_stacked_seed(
        self,
        n_rows: int,
        prows: list[list[Any]] | None = None,
        headers: list[str] | None = None,
        *,
        source_mi: int | None = None,
    ) -> list[str]:
        """積み上げ join seed: 前段 table_rows 行ごとの __file_path。"""
        from svc.data_agg_master_preview import (  # noqa: WPS433
            table_row_file_paths_for_stacked_seed,
        )

        n = max(0, int(n_rows))
        if n <= 0:
            return []
        hdrs = list(headers or [])
        rows = list(prows or [])[:n]
        stored = self._mpv_find_row_file_paths_for_stacked_seed(
            n,
            start_mi=int(source_mi) if source_mi is not None else None,
        )
        paths = table_row_file_paths_for_stacked_seed(
            hdrs,
            rows,
            scan_paths=list(self._debug_scan_paths or []),
            stored_row_paths=stored,
        )
        if paths:
            return paths[:n]
        scan = list(self._debug_scan_paths or [])
        if not scan:
            return []
        return [str(p) for p in scan[:n]]

    def _show_run_progress(
        self,
        phase: str,
        done: int,
        total: int,
        *,
        window_title: str = "",
        pct_override: int | None = None,
        detail: str = "",
        current_file: str = "",
    ) -> None:
        """ステップ実行時の共通進捗（NonModal・親 HWND・細かい phase_total を pickle で更新）。"""
        try:
            tot = max(1, int(total))
            dn = max(0, min(tot, int(done)))
            if pct_override is not None:
                raw_pct = max(0, min(100, int(round(float(pct_override)))))
            else:
                raw_pct = max(0, min(100, int(dn * 100 / tot)))
            if dn <= 1:
                self._master_progress_pct_floor = 0
                self._master_batch_hook_last_fi = 1
                self._master_batch_hook_last_nf = 1
            fl = int(getattr(self, "_master_progress_pct_floor", 0) or 0)
            pct = max(fl, raw_pct)
            if dn < tot:
                pct = min(pct, 95)
            self._master_progress_pct_floor = pct
            ph = self._debug_parent_hwnd()
            if self._run_progress_dlg is None or self._run_progress_path is None:
                self._ensure_master_run_cancel()
                pdir = ipc_file.get_ipc_root() / "progress"
                pdir.mkdir(parents=True, exist_ok=True)
                self._run_progress_path = pdir / (
                    "progress_data_agg_debug_%s.pkl"
                    % datetime.now().strftime("%Y%m%d%H%M%S%f")
                )
                req = {
                    "action": "progress",
                    "progress_path": str(self._run_progress_path),
                    "phase_total": tot,
                    "excel_lock": False,
                    "no_native_window": True,
                    "close_parent_when_done": False,
                    "non_modal_progress": False,
                    "done_delay_ms": 220,
                    "progress_poll_ms": 120,
                    "center_on_parent_widget": True,
                    "bring_excel_first": False,
                }
                cr = str(getattr(self, "_run_cancel_path", "") or "").strip()
                if cr:
                    req["cancel_request_path"] = cr
                    req["master_debug_cancel_cb"] = (
                        self._on_master_progress_cancel_clicked
                    )
                cfg = {}
                try:
                    import json

                    from core.core_cst import resolve_config_file_path

                    p = resolve_config_file_path("ui_data_agg.json")
                    raw = json.loads(p.read_text(encoding="utf-8"))
                    main_cfg = dict(raw.get("MAIN") or {})
                    pr_cfg = dict(
                        ((raw.get("SCREENS") or {}).get("PROGRESS") or {})
                    )
                    cfg = _deep_merge(main_cfg, pr_cfg)
                    cfg["TITLE"] = ""
                except Exception:
                    cfg = {"TITLE": ""}
                self._run_progress_dlg = create_progress_dialog(
                    req, ph, self, progress_cfg=cfg
                )
                self._run_progress_dlg.show()
                try:
                    QTimer.singleShot(
                        0,
                        self._activate_master_run_progress_for_cancel,
                    )
                except Exception:
                    pass
                self._run_progress_seq = 0
                self._set_debug_progress_locked(True)
                self._master_run_progress_active = True
                self._ensure_master_cancel_pump_timer()
            self._run_progress_seq += 1
            wt = str(window_title or "").strip()
            det = str(detail or "").strip()
            curf = str(current_file or "").strip()
            payload: dict[str, Any] = {
                "status": "RUN",
                "seq": int(self._run_progress_seq),
                "phase_total": tot,
                "phase_i": max(1, dn) if dn > 0 else 1,
                "phase": str(phase or "実行中"),
                "done": dn,
                "total": tot,
                "pct": pct,
                "current_file": curf,
                "window_title": wt,
                "msg": det,
                "detail": det,
            }
            ipc_file.write_pickle(
                self._run_progress_path,
                payload,
            )
            self._process_events_light()
        except Exception:
            pass

    def _mpv_effective_progress_hook_paths(self) -> list[str]:
        hook = getattr(self, "_mpv_progress_hook_paths", None)
        if hook:
            return list(hook)
        return list(self._debug_scan_paths or [])

    def _mpv_clear_wb_worker_if_done(self) -> None:
        t = getattr(self, "_mpv_wb_worker_thread", None)
        if t is not None and not t.is_alive():
            self._mpv_wb_worker_thread = None

    def _mpv_wb_worker_alive(self) -> bool:
        self._mpv_clear_wb_worker_if_done()
        t = getattr(self, "_mpv_wb_worker_thread", None)
        return t is not None and t.is_alive()

    def _mpv_queue_wb_frame_pending(self, frame: dict[str, Any]) -> None:
        pending = getattr(self, "_mpv_item_wb_pending_close", None)
        if not isinstance(pending, list):
            pending = []
            self._mpv_item_wb_pending_close = pending
        pending.append(frame)

    def _mpv_close_item_wb_frame(self) -> None:
        """項目 workbook 共有フレームを閉じる。compute 中は遅延破棄する。"""
        frame = getattr(self, "_mpv_item_wb_frame", None)
        self._mpv_item_wb_frame = None
        self._mpv_item_wb_mi = None
        if frame is None:
            return
        self._mpv_dispose_wb_frame(frame)

    def _mpv_dispose_wb_frame(self, frame: dict[str, Any]) -> None:
        """フレームを閉じる。ワーカー生存中または compute ロック中は pending。"""
        # join タイムアウト後など、ロックは空でもワーカーが frame を参照し得る
        if self._mpv_wb_worker_alive():
            self._mpv_queue_wb_frame_pending(frame)
            return
        lock = getattr(self, "_mpv_prog_compute_lock", None)
        if lock is not None:
            got = False
            try:
                got = bool(lock.acquire(blocking=False))
            except Exception:
                got = False
            if not got:
                self._mpv_queue_wb_frame_pending(frame)
                return
            try:
                self._mpv_close_wb_frame_obj(frame)
                self._mpv_flush_pending_wb_frames_unlocked()
            finally:
                try:
                    lock.release()
                except Exception:
                    pass
            return
        self._mpv_close_wb_frame_obj(frame)

    def _mpv_close_wb_frame_obj(self, frame: dict[str, Any] | None) -> None:
        if not frame:
            return
        try:
            from svc.svc_data_agg_extract import close_workbook_cache_frame

            close_workbook_cache_frame(frame)
        except Exception:
            pass

    def _mpv_flush_pending_wb_frames_unlocked(self) -> None:
        if self._mpv_wb_worker_alive():
            return
        pending = getattr(self, "_mpv_item_wb_pending_close", None)
        if not pending:
            return
        self._mpv_item_wb_pending_close = []
        for fr in list(pending):
            self._mpv_close_wb_frame_obj(fr)

    def _mpv_ensure_item_wb_frame(self, mi_idx: int) -> dict[str, Any]:
        mi = int(mi_idx)
        cur = getattr(self, "_mpv_item_wb_frame", None)
        cur_mi = getattr(self, "_mpv_item_wb_mi", None)
        if cur is not None and cur_mi is not None and int(cur_mi) == mi:
            return cur
        self._mpv_close_item_wb_frame()
        from svc.svc_data_agg_extract import new_workbook_cache_frame

        frame = new_workbook_cache_frame()
        self._mpv_item_wb_frame = frame
        self._mpv_item_wb_mi = mi
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_item_wb_cache open mi_idx=%s",
                mi,
            )
        except Exception:
            pass
        return frame

    def _mpv_try_flush_pending_wb_frames(self) -> None:
        """キャンセル後などに残った pending frame を、ワーカー終了後に閉じる。"""
        self._mpv_clear_wb_worker_if_done()
        pending = getattr(self, "_mpv_item_wb_pending_close", None)
        if not pending:
            return
        if self._mpv_wb_worker_alive():
            try:
                QTimer.singleShot(250, self._mpv_try_flush_pending_wb_frames)
            except Exception:
                pass
            return
        lock = getattr(self, "_mpv_prog_compute_lock", None)
        if lock is not None:
            got = False
            try:
                got = bool(lock.acquire(blocking=False))
            except Exception:
                got = False
            if not got:
                try:
                    QTimer.singleShot(250, self._mpv_try_flush_pending_wb_frames)
                except Exception:
                    pass
                return
            try:
                self._mpv_flush_pending_wb_frames_unlocked()
            finally:
                try:
                    lock.release()
                except Exception:
                    pass
            return
        self._mpv_flush_pending_wb_frames_unlocked()

    def _mpv_item_wb_bind(self, mi_idx: int):
        """項目単位共有フレームを現在スレッドの TLS に載せる context manager。"""
        from contextlib import nullcontext

        from svc.svc_data_agg_extract import bind_workbook_cache_frame

        frame = self._mpv_ensure_item_wb_frame(int(mi_idx))
        if frame is None:
            return nullcontext()
        return bind_workbook_cache_frame(frame)

    def _master_progress_cache_mark_prefix(self, detail_s: str) -> str:
        """raw detail 先頭の [C]/[F] を取り出す（進捗 1 行目＝フェーズ用）。"""
        s = str(detail_s or "").lstrip()
        if s.startswith("[C]"):
            return "[C] "
        if s.startswith("[F]"):
            return "[F] "
        return ""

    @staticmethod
    def _master_progress_strip_cache_mark(text: str) -> str:
        """詳細行から先頭の [C]/[F] を除く（フェーズ側へ移したあとの二重表示防止）。"""
        s = str(text or "").lstrip()
        for prefix in ("[C] ", "[F] ", "[C]", "[F]"):
            if s.startswith(prefix):
                return s[len(prefix) :].lstrip()
        return s

    def _master_progress_resolve_cache_mark(self, detail_s: str, file_index: int) -> str:
        """detail のマーク、なければ同一ファイルの直近マークを返す。"""
        mark = self._master_progress_cache_mark_prefix(detail_s)
        fi = int(file_index or 0)
        if mark:
            self._master_batch_hook_last_cache_mark = mark
            self._master_batch_hook_mark_fi = fi
            return mark
        if fi and fi == int(getattr(self, "_master_batch_hook_mark_fi", 0) or 0):
            return str(getattr(self, "_master_batch_hook_last_cache_mark", "") or "")
        return ""

    def _master_progress_phase_with_file(
        self, phase_head: str, *, mark: str, cur_file: str
    ) -> str:
        """1 行目: [C]/[F] + フェーズ + 現在ファイル名。"""
        head = str(phase_head or "").strip() or "準備中"
        m = str(mark or "")
        if m:
            head = "%s%s" % (m, head)
        fn = str(cur_file or "").strip()
        if fn:
            head = "%s — %s" % (head, fn)
        return head

    def _master_progress_pick_current_file(
        self, detail: str, file_index: int, n_files: int
    ) -> str:
        from svc.svc_data_agg import _batch_hook_resolve_current_file  # noqa: WPS433

        detail_s = str(detail or "")
        row_m = re.search(
            r"行\s*\d+\s*/\s*\d+\s*:\s*(.+?)(?:\s+読込中|（完了）|$)",
            detail_s,
        )
        if row_m:
            return row_m.group(1).strip()
        paths = self._mpv_effective_progress_hook_paths()
        eff_nf = max(1, len(paths)) if paths else max(1, int(n_files))
        eff_fi = int(file_index)
        if paths and 1 <= eff_fi <= len(paths):
            eff_nf = len(paths)
        name = _batch_hook_resolve_current_file(detail_s, eff_fi, paths)
        if name:
            return name
        if eff_nf > 0 and 1 <= eff_fi <= len(paths):
            from pathlib import Path

            return Path(str(paths[int(eff_fi) - 1])).name
        m_colon = re.search(r":\s*([^\s—]+)$", detail_s)
        if m_colon:
            return m_colon.group(1).strip()
        return ""

    def _master_batch_hook_ui_phase(self, sub_phase: int, detail_s: str) -> int:
        """compute_batch の sub 4〜7 を UI 段 done=4..10 に割り当てる。"""
        s = str(detail_s or "")
        if sub_phase == 4:
            if (
                "結合準備中" in s
                or re.search(r"^0/\d+", s.strip())
                or s.strip() == "0 件"
            ):
                return 4
            return 5
        if sub_phase == 5:
            return 6
        if sub_phase == 6:
            if "結合索引" in s:
                return 8
            if re.search(r"結合項目", s) and "候補プール" in s:
                return 7
            return 9
        if sub_phase == 7:
            return 10
        return max(4, min(10, int(sub_phase) + 1))

    def _master_batch_hook_pct(
        self,
        ui_done: int,
        *,
        fi: int,
        nf: int,
        intra: float = 0.0,
    ) -> int:
        tot = len(_MASTER_DEBUG_PROGRESS_PHASES)
        ui = max(1, min(tot, int(ui_done)))
        nf_m = max(1, int(nf))
        fi_m = max(1, min(nf_m, int(fi)))
        batch_lo = 4
        if ui < batch_lo:
            return min(95, int(round(100.0 * ui / tot)))
        batch_idx = ui - batch_lo
        denom = max(1, _MASTER_DEBUG_BATCH_UI_PHASE_COUNT * nf_m)
        step_pos = (fi_m - 1) * _MASTER_DEBUG_BATCH_UI_PHASE_COUNT + batch_idx
        step_pos += min(1.0, max(0.0, float(intra)))
        frac = min(1.0, step_pos / float(denom))
        lo = int(round(100.0 * (batch_lo - 1) / tot))
        hi = 95
        return min(hi, int(round(lo + (hi - lo) * frac)))

    def _master_progress_format_read_start_detail(
        self, *, fi: int, nf: int, raw: str
    ) -> str:
        if nf > 0:
            return "0/%s — 開始" % nf
        r = str(raw or "").strip()
        return r if r else "開始"

    def _master_progress_format_extract_detail(
        self,
        detail_s: str,
        *,
        fi: int,
        nf: int,
        cur_file: str,
    ) -> str:
        raw = str(detail_s or "").strip()
        eff_nf = max(1, int(nf))
        eff_fi = int(fi)
        # ファイル名は 1 行目へ移すため、2 行目には載せない
        _ = str(cur_file or "").strip()
        item_m = re.search(
            r"項目\s*(\d+)\s*/\s*(\d+)\s*:\s*(.+?)(?:\s*（|$)",
            raw,
        )
        if item_m:
            label = self._master_progress_strip_cache_mark(item_m.group(3).strip())
            return "読込 %s/%s · 項目 %s/%s %s" % (
                eff_fi,
                eff_nf,
                item_m.group(1),
                item_m.group(2),
                label,
            )
        row_m = re.search(
            r"行\s*(\d+)\s*/\s*(\d+)\s*:\s*(.+?)\s+読込中",
            raw,
        )
        if row_m:
            return "読込 %s/%s · 行 %s/%s" % (
                eff_fi,
                eff_nf,
                row_m.group(1),
                row_m.group(2),
            )
        row_done = re.search(
            r"行\s*(\d+)\s*/\s*(\d+)\s*:\s*(.+?)（完了）",
            raw,
        )
        if row_done:
            return "読込完了 · 行 %s/%s" % (row_done.group(1), row_done.group(2))
        file_m = re.search(
            r"ファイル\s*(\d+)\s*/\s*(\d+)\s*:\s*(.+?)(?:\s+読込中|（完了）|$)",
            raw,
        )
        if file_m:
            return "読込 %s/%s" % (
                file_m.group(1),
                file_m.group(2),
            )
        if raw:
            return self._master_progress_strip_cache_mark(raw)
        return "読込 %s/%s" % (eff_fi, eff_nf)

    def _master_progress_format_merge_detail(
        self,
        detail_s: str,
        *,
        fi: int,
        nf: int,
        cur_file: str,
    ) -> str:
        raw = str(detail_s or "").strip()
        eff_nf = max(1, int(nf))
        eff_fi = int(fi)
        row_n = re.search(r"（\s*(\d+)\s*行\s*）", raw)
        if row_n:
            return "ファイル %s/%s — %s 行" % (
                eff_fi,
                eff_nf,
                row_n.group(1),
            )
        # ファイル名は 1 行目へ移す
        if cur_file or raw:
            return "ファイル %s/%s" % (eff_fi, eff_nf)
        return "ファイル %s/%s" % (eff_fi, eff_nf)

    def _master_progress_format_join_prep_detail(self, detail_s: str) -> str:
        raw = str(detail_s or "").strip()
        pool_m = re.search(r"候補プール\s*(\d+)\s*行", raw)
        if pool_m:
            return "候補プール %s 行" % pool_m.group(1)
        return raw or "結合準備中"

    def _master_progress_format_join_index_detail(self, detail_s: str) -> str:
        raw = str(detail_s or "").strip()
        pool_m = re.search(r"候補プール\s*(\d+)\s*行", raw)
        if pool_m:
            return "索引構築中（%s 行）" % pool_m.group(1)
        return raw or "索引構築中"

    def _master_progress_format_join_match_detail(
        self,
        detail_s: str,
        *,
        fi: int,
        nf: int,
        cur_file: str,
    ) -> str:
        raw = str(detail_s or "").strip()
        eff_nf = max(1, int(nf))
        eff_fi = int(fi)
        fn = str(cur_file or "").strip()
        if not fn:
            lead_m = re.match(r"^(.+?)\s+結合\s*\d", raw)
            if lead_m:
                fn = lead_m.group(1).strip()
        join_m = re.search(r"結合\s*(\d+)\s*/\s*(\d+)", raw)
        if join_m:
            base = "照合 %s/%s" % (eff_fi, eff_nf)
            if fn:
                base = "%s — %s" % (base, fn)
            return "%s · 結合 %s/%s" % (
                base,
                join_m.group(1),
                join_m.group(2),
            )
        file_m = re.search(r"ファイル\s*(\d+)\s*/\s*(\d+)", raw)
        if file_m:
            eff_fi = int(file_m.group(1))
            eff_nf = max(1, int(file_m.group(2)))
        if fn:
            return "照合 %s/%s — %s" % (eff_fi, eff_nf, fn)
        return "照合 %s/%s" % (eff_fi, eff_nf)

    def _master_progress_format_assemble_detail(
        self,
        detail_s: str,
        *,
        fi: int,
        nf: int,
        cur_file: str,
    ) -> str:
        raw = str(detail_s or "").strip()
        eff_nf = max(1, int(nf))
        eff_fi = int(fi)
        fn = str(cur_file or "").strip()
        row_m = re.search(r"一覧行\s*(\d+)\s*/\s*(\d+)", raw)
        if row_m:
            base = "一覧 %s/%s" % (row_m.group(1), row_m.group(2))
            if fn:
                return "%s — %s" % (base, fn)
            return base
        file_m = re.search(
            r"一覧行を組立\s*(\d+)\s*/\s*(\d+).*?（\s*(\d+)\s*行）",
            raw,
        )
        if file_m:
            if fn:
                return "一覧 %s/%s — %s" % (
                    file_m.group(1),
                    file_m.group(2),
                    fn,
                )
            return "一覧 %s/%s" % (file_m.group(1), file_m.group(2))
        if fn:
            return "一覧 %s/%s — %s" % (eff_fi, eff_nf, fn)
        if raw:
            return "一覧 %s/%s — %s" % (eff_fi, eff_nf, raw)
        return "一覧 %s/%s" % (eff_fi, eff_nf)

    def _master_dbg_batch_progress_hook(self, sub_phase: int, detail: str, *rest: Any) -> None:
        """compute_batch_table_rows からのコールバック（phase 4〜7）。rest は file_index, n_files（任意）。"""
        self._master_poll_cancel(force=True)
        if not (
            getattr(self, "_master_run_progress_active", False)
            or getattr(self, "_debug_progress_locked", False)
        ):
            try:
                dlg = getattr(self, "_run_progress_dlg", None)
                if dlg is None or not dlg.isVisible():
                    return
            except Exception:
                return
        if not (4 <= sub_phase <= 7):
            return
        detail_s = str(detail or "")
        if len(rest) >= 2:
            self._master_batch_hook_last_fi = max(1, int(rest[0]))
            self._master_batch_hook_last_nf = max(1, int(rest[1]))
        else:
            m = re.search(r"（\s*(\d+)\s*/\s*(\d+)\s*）", detail_s)
            if m:
                self._master_batch_hook_last_fi = max(1, int(m.group(1)))
                self._master_batch_hook_last_nf = max(1, int(m.group(2)))
            elif "0 件" in detail_s and sub_phase == 4:
                self._master_batch_hook_last_fi = 1
                self._master_batch_hook_last_nf = 1
        fi = int(getattr(self, "_master_batch_hook_last_fi", 1) or 1)
        nf = max(1, int(getattr(self, "_master_batch_hook_last_nf", 1) or 1))
        ui_done = self._master_batch_hook_ui_phase(sub_phase, detail_s)
        phase_head = _MASTER_DEBUG_PROGRESS_PHASES[ui_done - 1]
        wt = getattr(self, "_master_progress_window_title", "") or ""
        mi_m = re.search(r"項目\s*(\d+)\s*/\s*(\d+)", detail_s)
        row_m = re.search(r"行\s*(\d+)\s*/\s*(\d+)", detail_s)
        join_m = re.search(r"結合\s*(\d+)\s*/\s*(\d+)", detail_s)
        intra = 0.0
        if ui_done == 5 and mi_m:
            inn = max(1, int(mi_m.group(2)))
            intra = min(1.0, max(0.0, int(mi_m.group(1)) / float(inn)))
        elif ui_done == 5 and row_m:
            rnn = max(1, int(row_m.group(2)))
            intra = min(1.0, max(0.0, int(row_m.group(1)) / float(rnn)))
        elif ui_done == 9 and join_m:
            jnn = max(1, int(join_m.group(2)))
            intra = min(1.0, max(0.0, int(join_m.group(1)) / float(jnn)))
        elif ui_done == 10 and row_m:
            rnn = max(1, int(row_m.group(2)))
            intra = min(1.0, max(0.0, int(row_m.group(1)) / float(rnn)))
        pct_ov = self._master_batch_hook_pct(ui_done, fi=fi, nf=nf, intra=intra)
        hook_paths = self._mpv_effective_progress_hook_paths()
        hook_nf = len(hook_paths) if hook_paths else nf
        cur_file = self._master_progress_pick_current_file(detail_s, fi, hook_nf)
        if ui_done == 4:
            show_detail = self._master_progress_format_read_start_detail(
                fi=fi, nf=nf, raw=detail_s
            )
        elif ui_done == 5:
            show_detail = self._master_progress_format_extract_detail(
                detail_s, fi=fi, nf=nf, cur_file=cur_file
            )
        elif ui_done == 6:
            show_detail = self._master_progress_format_merge_detail(
                detail_s, fi=fi, nf=nf, cur_file=cur_file
            )
        elif ui_done == 7:
            show_detail = self._master_progress_format_join_prep_detail(detail_s)
        elif ui_done == 8:
            show_detail = self._master_progress_format_join_index_detail(detail_s)
        elif ui_done == 9:
            show_detail = self._master_progress_format_join_match_detail(
                detail_s, fi=fi, nf=nf, cur_file=cur_file
            )
        else:
            show_detail = self._master_progress_format_assemble_detail(
                detail_s, fi=fi, nf=nf, cur_file=cur_file
            )
        # [C]/[F] + ファイル名は 1 行目。項目進捗でもマークを sticky 維持
        mark = self._master_progress_resolve_cache_mark(detail_s, fi)
        show_detail = self._master_progress_strip_cache_mark(show_detail)
        phase_head = self._master_progress_phase_with_file(
            phase_head, mark=mark, cur_file=cur_file if ui_done in (4, 5, 6) else ""
        )
        self._show_run_progress(
            phase_head,
            ui_done,
            len(_MASTER_DEBUG_PROGRESS_PHASES),
            window_title=wt,
            pct_override=pct_ov,
            detail=show_detail,
            current_file="",
        )
        self._process_events_light()
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        try:
            self._master_poll_cancel(force=True)
        except DataAggCancelled:
            self._master_note_cancel_requested()
            raise

    def _close_run_progress(self, *, cancelled: bool = False) -> None:
        dlg = self._run_progress_dlg
        try:
            if self._run_progress_path is not None:
                self._run_progress_seq += 1
                if cancelled:
                    from svc.data_agg_cancel import write_progress_cancel_status  # noqa: WPS433

                    try:
                        write_progress_cancel_status(self._run_progress_path)
                    except Exception:
                        ipc_file.write_pickle(
                            self._run_progress_path,
                            {
                                "status": "CANCEL",
                                "seq": int(self._run_progress_seq),
                                "phase": "中止",
                                "pct": 5,
                            },
                        )
                else:
                    ipc_file.write_pickle(
                        self._run_progress_path,
                        {
                            "status": "DONE",
                            "seq": int(self._run_progress_seq),
                            "phase_i": 8,
                            "phase": _MASTER_DEBUG_PROGRESS_PHASE_DONE,
                            "done": 8,
                            "total": 8,
                            "pct": 100,
                        },
                    )
            QApplication.processEvents()
            if dlg is not None:
                try:
                    dlg.close()
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            self._set_debug_progress_locked(False)
            self._master_run_progress_active = False
            self._stop_master_cancel_pump_timer()
            self._run_progress_dlg = None
            self._run_progress_path = None
            if not getattr(self, "_continuous_busy", False):
                self._clear_master_run_cancel()
            try:
                self._update_run_buttons_state()
                self._update_clear_buttons()
            except Exception:
                pass

    def _icap(self, lst: list[str]) -> list[str]:
        # #n[...] 形式はキー別に上限済み（svc 側）なので、ここで全体再制限しない。
        if any(re.match(r"^#\d+\[[^\]]*\]", str(x)) for x in lst):
            return list(lst)
        cap = self._max_value_rows()
        if len(lst) == cap + 1 and lst and "省略" in str(lst[-1]):
            return list(lst)
        return _cap_list_capped(lst, cap)

    def _icap_with_tips(
        self, colvals: list[str], tips: list[str | None] | None
    ) -> tuple[list[str], list[str | None]]:
        """値列の上限適用とツールチップ行の整合（svc が省略行付きで返した場合はそのまま）。"""
        tt: list[str | None] = list(tips or [])
        if len(tt) != len(colvals):
            tt = _none_tips(len(colvals))
        if any(re.match(r"^#\d+\[[^\]]*\]", str(x)) for x in colvals):
            return list(colvals), tt
        cap = self._max_value_rows()
        if len(colvals) == cap + 1 and colvals and "省略" in str(colvals[-1]):
            return list(colvals), tt
        if len(colvals) <= cap:
            return list(colvals), tt
        out = _cap_list_capped(colvals, cap)
        out_t: list[str | None] = list(tt[:cap])
        out_t.append(None)
        return out, out_t

    def _slot_summary_row(self, vals: list[str]) -> str:
        parts = []
        for h, v in zip(self._summary_log_headers(), vals):
            if v != "-":
                parts.append("%s=%s" % (h, v))
        return "、".join(parts) if parts else "-"

    def _empty_scenario_state(self) -> dict[str, Any]:
        return {
            "phase_idx": 0,
            "summary_rows": [],
            "value_cols": [],
            "value_col_tooltips": [],
            "summary_phase_labels": [],
            "log": "",
        }

    def _ensure_scenario_state(self, idx: int) -> dict[str, Any]:
        if idx not in self._scenario_snapshots:
            self._scenario_snapshots[idx] = self._empty_scenario_state()
        return self._scenario_snapshots[idx]

    def _persist_scenario_state(self) -> None:
        if self._mode != 0:
            return
        st = self._ensure_scenario_state(self._sc_idx)
        st["phase_idx"] = self._phase_idx
        st["summary_rows"] = copy.deepcopy(self._summary_rows)
        st["value_cols"] = copy.deepcopy(self._value_cols)
        st["value_col_tooltips"] = copy.deepcopy(self._value_col_tooltips)
        st["summary_phase_labels"] = list(self._summary_phase_labels)
        st["log"] = self.log.toPlainText()

    def _load_scenario_state(self, idx: int) -> None:
        st = self._ensure_scenario_state(idx)
        self._phase_idx = int(st["phase_idx"])
        self._summary_rows = copy.deepcopy(st["summary_rows"])
        self._value_cols = copy.deepcopy(st["value_cols"])
        _vt_raw = st.get("value_col_tooltips")
        if isinstance(_vt_raw, list) and len(_vt_raw) == len(self._value_cols):
            self._value_col_tooltips = copy.deepcopy(_vt_raw)
        else:
            self._value_col_tooltips = [
                _none_tips(len(c)) for c in self._value_cols
            ]
        self._summary_phase_labels = list(st["summary_phase_labels"])
        self.log.setPlainText(str(st.get("log", "")))
        self._sync_summary_table_from_lists()
        self._rebuild_value_grid()

    def _sync_summary_table_from_lists(self) -> None:
        try:
            self._ensure_summary_table_columns()
            self.summary_table.setRowCount(0)
            for i, vals in enumerate(self._summary_rows):
                lab = (
                    self._summary_phase_labels[i]
                    if i < len(self._summary_phase_labels)
                    else str(i + 1)
                )
                row = [lab] + self._summary_vals_for_display(vals)
                r = self.summary_table.rowCount()
                self.summary_table.insertRow(r)
                for c, t in enumerate(row):
                    disp = t if c == 0 else _summary_metric_cell_display(str(t))
                    self.summary_table.setItem(r, c, QTableWidgetItem(disp))
            self._fit_summary_table_columns()
        except Exception:
            _logger.exception("summary table sync failed")
            self._log_append("【内部】サマリ表の同期に失敗しました。コンソールに詳細を出力しました。")

    def closeEvent(self, event: QCloseEvent) -> None:
        self._continuous_busy = False
        self._continuous_steps_left = 0
        self._cancel_scenario_link_prefetch(join=False)
        self._bump_mpv_prefetch_cancel()
        self._mpv_close_item_wb_frame()
        self._mpv_try_flush_pending_wb_frames()
        if self._mode == 0:
            self._persist_scenario_state()
        self._scenario_snapshots.clear()
        self._clear_master_item_snapshots()
        super().closeEvent(event)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        top = QHBoxLayout()
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(
            [
                self._d(
                    "MODE_LABEL_SCENARIO",
                    "シナリオフェーズ実行（シナリオ編集から起動想定）",
                ),
                self._d(
                    "MODE_LABEL_MASTER",
                    "マスタ項目ステップ実行（データ集約メインから起動想定）",
                ),
            ]
        )
        self.mode_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
        )
        self.mode_combo.setMinimumContentsLength(0)
        mcw = self._window_int("MODE_COMBO_MAX_WIDTH", 0)
        if mcw > 0:
            self.mode_combo.setMaximumWidth(mcw)
        self._lbl_mode = QLabel(
            "<b>%s</b>" % self._d("LABEL_MODE", "モード").replace("\n", "<br/>")
        )
        self._wrap_desc_label(self._lbl_mode, h_policy=QSizePolicy.Policy.Minimum)
        top.addWidget(self._lbl_mode)
        if self._fixed_mode is not None:
            self.mode_combo.setCurrentIndex(self._mode)
            self.mode_combo.hide()
            self._mode_fixed_label = QLabel(self.mode_combo.currentText())
            self._wrap_desc_label(self._mode_fixed_label)
            top.addWidget(self._mode_fixed_label, 1)
        else:
            self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
            top.addWidget(self.mode_combo, 1)
        root.addLayout(top)

        self.hint = QLabel()
        self._wrap_desc_label(self.hint)
        root.addWidget(self.hint)

        self._debug_main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main = self._debug_main_splitter
        root.addWidget(main, 1)

        self._debug_left_panel = QWidget()
        left = self._debug_left_panel
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 6, 0)
        self.left_title = QLabel()
        self._wrap_desc_label(self.left_title)
        ll.addWidget(self.left_title)

        self._left_col_split = QSplitter(Qt.Orientation.Vertical)
        self.left_table = QTableWidget(0, 1)
        self.left_table.setHorizontalHeaderLabels(
            [self._d("LEFT_TABLE_HEADER_SCENARIO", "シナリオ")]
        )
        self.left_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.left_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.left_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.left_table.verticalHeader().setVisible(False)
        self.left_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.left_table.itemSelectionChanged.connect(self._on_left_sel)
        self.left_table.setMinimumWidth(0)
        self._left_col_split.addWidget(self.left_table)

        self._debug_steps_host = QWidget()
        steps_host = self._debug_steps_host
        sl = QVBoxLayout(steps_host)
        sl.setContentsMargins(0, 0, 0, 0)
        sl.setSpacing(4)
        self._lbl_steps = QLabel(
            self._d(
                "LABEL_PHASE_STEPS_HTML",
                "<b>登録シナリオ</b>（番号＝結果のフェーズ列・サマリ行と対応）",
            )
        )
        self._wrap_desc_label(self._lbl_steps)
        sl.addWidget(self._lbl_steps)
        self.left_steps = QTableWidget(0, 2)
        self.left_steps.setHorizontalHeaderLabels(
            [
                self._d("COND_COL_NO", "番号"),
                self._d("COND_COL_ITEM", "大項目"),
            ]
        )
        self.left_steps.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.left_steps.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.left_steps.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.left_steps.verticalHeader().setVisible(False)
        _ls_h = self.left_steps.horizontalHeader()
        _ls_h.setMinimumSectionSize(self._table_hdr_min_section())
        _ls_h.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        _ls_h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        _ls_h.setDefaultSectionSize(56)
        self.left_steps.itemSelectionChanged.connect(self._on_step_sel)
        self.left_steps.setWordWrap(False)
        self.left_steps.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.left_steps.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.left_steps.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.left_steps.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.left_steps.setMinimumWidth(0)
        sl.addWidget(self.left_steps, 1)
        self._left_col_split.addWidget(steps_host)
        self._left_col_split.setChildrenCollapsible(False)
        ll.addWidget(self._left_col_split, 1)
        self._apply_left_column_splitter_sizes()

        self.left_detail = QLabel()
        self.left_detail.setVisible(False)

        nav = QHBoxLayout()
        self.btn_prev = QPushButton(self._d("BTN_PREV", "前へ"))
        self.btn_next = QPushButton(self._d("BTN_NEXT", "次へ"))
        self.btn_prev.clicked.connect(lambda: self._nav_left(-1))
        self.btn_next.clicked.connect(lambda: self._nav_left(1))
        nav.addStretch(1)
        nav.addWidget(self.btn_prev)
        nav.addWidget(self.btn_next)
        ll.addLayout(nav)

        left.setMinimumWidth(0)
        lmax = self._window_int("LEFT_PANE_MAX_WIDTH", 0)
        if lmax > 0:
            left.setMaximumWidth(lmax)
        main.addWidget(left)

        self._debug_right_panel = QWidget()
        right = self._debug_right_panel
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tab_cond = QWidget()
        self.tab_res = QWidget()
        self.tab_log = QWidget()
        self._build_tab_conditions()
        self._build_tab_results()
        self._build_tab_log()
        self.tabs.addTab(self.tab_cond, self._d("TAB_CONDITIONS", "条件"))
        self.tabs.addTab(self.tab_res, self._d("TAB_RESULTS", "結果"))
        self.tabs.addTab(self.tab_log, self._d("TAB_LOG", "ログ"))
        self.tabs.setMinimumWidth(0)
        rl.addWidget(self.tabs, 1)

        btn_row = QHBoxLayout()
        self.btn_run = QPushButton()
        self.btn_run_all = QPushButton(self._btn_run_all_label())
        self.btn_run_all_master = QPushButton(
            self._d("BTN_RUN_ALL_MASTER_ITEMS", "一括実行")
        )
        self.btn_clear_res = QPushButton(self._d("BTN_CLEAR_RESULTS", "結果クリア"))
        self.btn_clear_log = QPushButton(self._d("BTN_CLEAR_LOG", "ログクリア"))
        self.btn_cancel = QPushButton(self._d("BTN_CANCEL", "閉じる"))
        self.btn_run.clicked.connect(self._on_run)
        self.btn_run_all.clicked.connect(self._on_run_all_continuous)
        self.btn_run_all_master.clicked.connect(self._on_run_all_master_items_continuous)
        self.btn_clear_res.clicked.connect(self._on_clear_results)
        self.btn_clear_log.clicked.connect(self._on_clear_log_only)
        self.btn_cancel.clicked.connect(self._on_close_request)
        btn_row.addWidget(self.btn_run)
        btn_row.addWidget(self.btn_run_all)
        btn_row.addWidget(self.btn_run_all_master)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_clear_res)
        btn_row.addWidget(self.btn_clear_log)
        btn_row.addWidget(self.btn_cancel)
        rl.addLayout(btn_row)

        right.setMinimumWidth(0)
        for btn in (
            self.btn_run,
            self.btn_run_all,
            self.btn_run_all_master,
            self.btn_clear_res,
            self.btn_clear_log,
            self.btn_cancel,
        ):
            btn.setMinimumWidth(0)
            btn.setSizePolicy(
                QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed
            )

        main.addWidget(right)
        self._apply_debug_splitter_sizes(main)
        self._apply_static_debug_tooltips()

    def _apply_debug_splitter_sizes(self, main: QSplitter) -> None:
        """WINDOW.DEFAULT_WIDTH に収まるようスプリッタ幅を設定（SPLITTER_SIZES は超過時スケール）。"""
        wc = self._cfg.get("WINDOW") or {}
        dw = int(wc.get("DEFAULT_WIDTH") or 860)
        margin = int(wc.get("SPLITTER_MARGIN") or 24)
        budget = max(dw - margin, 80)
        raw = wc.get("SPLITTER_SIZES")
        a, b = 260, 560
        if isinstance(raw, list) and len(raw) >= 2:
            try:
                a = max(0, int(raw[0]))
                b = max(0, int(raw[1]))
            except (TypeError, ValueError):
                pass
        total = a + b
        if total <= 0:
            a = budget // 3
            b = budget - a
        elif total > budget:
            scale = budget / total
            a = max(0, int(a * scale))
            b = max(0, int(b * scale))
            if a + b < budget:
                b = budget - a
        lmax = self._window_int("LEFT_PANE_MAX_WIDTH", 0)
        if lmax > 0 and a > lmax:
            b += a - lmax
            a = lmax
        main.setSizes([a, b])

    def _apply_left_column_splitter_sizes(self) -> None:
        """左ペイン内のマスタ項目表と条件ステップエリアの縦分割（WINDOW.LEFT_COLUMN_SPLIT_SIZES）。"""
        wc = self._cfg.get("WINDOW") or {}
        dh = int(wc.get("DEFAULT_HEIGHT") or 600)
        top, bot = max(100, int(dh * 0.38)), max(120, int(dh * 0.42))
        raw = wc.get("LEFT_COLUMN_SPLIT_SIZES")
        if isinstance(raw, list) and len(raw) >= 2:
            try:
                top = max(60, int(raw[0]))
                bot = max(80, int(raw[1]))
            except (TypeError, ValueError):
                pass
        sp = getattr(self, "_left_col_split", None)
        if sp is not None:
            sp.setSizes([top, bot])

    def _build_tab_conditions(self) -> None:
        lay = QVBoxLayout(self.tab_cond)
        self.cond_hint = QLabel()
        self._wrap_desc_label(self.cond_hint)
        lay.addWidget(self.cond_hint)
        self.cond_stack = QStackedWidget()
        self.cond_tree = _DebugCondTreeWidget()
        self.cond_tree.setHeaderLabels(
            [
                self._d("TREE_COL_NO", "番号"),
                self._d("TREE_COL_ITEM", "項目"),
                self._d("TREE_COL_SETTING", "要約"),
            ]
        )
        self._apply_cond_tree_initial_column_widths(self.cond_tree)
        self.master_cond_tree = _DebugCondTreeWidget()
        self.master_cond_tree.setHeaderLabels(
            [
                self._d("TREE_COL_NO", "番号"),
                self._d("TREE_COL_ITEM", "項目"),
                self._d("TREE_COL_SETTING", "要約"),
            ]
        )
        self._apply_cond_tree_initial_column_widths(self.master_cond_tree)
        self.cond_stack.addWidget(self.cond_tree)
        self.cond_stack.addWidget(self.master_cond_tree)
        self.cond_tree.setMinimumWidth(0)
        self.master_cond_tree.setMinimumWidth(0)
        self.cond_stack.setMinimumWidth(0)
        lay.addWidget(self.cond_stack, 1)

    def _apply_summary_table_resize_modes(self) -> None:
        """横の最小幅を抑え、狭いウィンドウでは横スクロールで閲覧する。縦もピクセル単位でスクロール可能にする。"""
        st = self.summary_table
        st.setMinimumWidth(0)
        st.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        st.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        st.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        hdr = st.horizontalHeader()
        smins = self._window_int_list("SUMMARY_TABLE_COL_MIN_WIDTHS")
        hdr.setMinimumSectionSize(
            self._header_global_floor(smins, self._table_hdr_min_section())
        )
        n = st.columnCount()
        for c in range(n):
            hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
        hdr.setDefaultSectionSize(max(48, min(72, hdr.defaultSectionSize())))

    def _build_tab_results(self) -> None:
        lay = QVBoxLayout(self.tab_res)
        self.res_hint = QLabel()
        self._wrap_desc_label(self.res_hint)
        lay.addWidget(self.res_hint)

        cols = self._summary_cols_display()
        self.summary_table = QTableWidget(0, len(cols))
        self.summary_table.setHorizontalHeaderLabels(cols)
        self.summary_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        hdr_s = self.summary_table.horizontalHeader()
        hdr_s.setMinimumHeight(28)
        hdr_s.setTextElideMode(Qt.TextElideMode.ElideNone)
        hdr_s.setDefaultAlignment(
            Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
        )
        self.summary_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._apply_summary_table_resize_modes()
        self.summary_table.setMinimumHeight(self._summary_table_min_height())
        self.summary_table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding
        )
        sum_fold_row = QHBoxLayout()
        self.btn_summary_fold = QToolButton()
        self.btn_summary_fold.setCheckable(True)
        self.btn_summary_fold.setChecked(False)
        self.btn_summary_fold.setText(
            self._d("BTN_SUMMARY_COLLAPSED", "▶ 結果サマリ")
        )
        self.btn_summary_fold.setToolButtonStyle(
            Qt.ToolButtonStyle.ToolButtonTextOnly
        )
        self.btn_summary_fold.toggled.connect(self._on_summary_fold_toggled)
        sum_fold_row.addWidget(self.btn_summary_fold, 0, Qt.AlignmentFlag.AlignLeft)
        sum_fold_row.addStretch(1)
        sum_wrap = QWidget()
        sumw_lay = QVBoxLayout(sum_wrap)
        sumw_lay.setContentsMargins(0, 0, 0, 0)
        sumw_lay.addWidget(self.summary_table)
        lay.addLayout(sum_fold_row)
        lay.addWidget(sum_wrap, 0)
        self._summary_table_wrap = sum_wrap
        sum_wrap.setVisible(False)

        self.values_title = QLabel()
        self._wrap_desc_label(self.values_title)
        lay.addWidget(self.values_title)

        self.value_grid = QTableWidget(0, 0)
        self.value_grid.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.value_grid.verticalHeader().setVisible(True)
        self.value_grid.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.value_grid.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.value_grid.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.value_grid.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.value_grid.setMinimumWidth(0)
        self.value_grid.setMinimumHeight(self._value_grid_min_height())
        self.value_grid.setTextElideMode(Qt.TextElideMode.ElideNone)
        self._value_grid_delegate = _ValueGridNoElideDelegate(self.value_grid)
        self.value_grid.setItemDelegate(self._value_grid_delegate)
        _vgh = _ValueGridPhaseHeader(Qt.Orientation.Horizontal, self.value_grid)
        self.value_grid.setHorizontalHeader(_vgh)
        _vgh.setTextElideMode(Qt.TextElideMode.ElideNone)
        _vgh.sectionResized.connect(self._on_value_grid_section_resized)
        lay.addWidget(self.value_grid, 1)
        self._ensure_summary_table_columns()

    def _style_results_table_header_rows(self) -> None:
        """結果サマリ・結果一覧グリッドの列・行見出しを薄いグレーにする（一覧の角は別途同期）。"""
        sheet = (
            "QHeaderView::section {"
            "  background-color: %s;"
            "  color: #222;"
            "  padding: 3px;"
            "  border: 1px solid #d0d0d0;"
            "}"
            % _DEBUG_RESULTS_HEADER_BG
        )
        try:
            self.summary_table.horizontalHeader().setStyleSheet(sheet)
            self.summary_table.verticalHeader().setStyleSheet(sheet)
            self._apply_value_grid_header_and_corner_stylesheet()
        except Exception:
            pass
        self._apply_values_title_snapshot_tint()

    def _master_snapshot_browseable(self) -> bool:
        return self._mode == 1 and self._master_snapshot_priority_active()

    def _master_showing_row_snapshot(self) -> bool:
        if not self._master_snapshot_browseable():
            return False
        mi = int(self._mi_idx)
        if mi in self._master_item_snapshots:
            return True
        if getattr(self, "_master_snapshot_browse_after_cancel", False):
            lr = self.left_steps.currentRow()
            if lr >= 0:
                cancel_mi = int(getattr(self, "_master_cancel_mi", mi))
                cancel_step = int(getattr(self, "_master_cancel_step", 0))
                if (
                    not self._master_item_has_no_scenario_at_mi(mi)
                    and mi < cancel_mi
                    and self._master_step_snapshot_key(mi, lr)
                    in self._master_step_snapshots
                ):
                    return True
                if (
                    not self._master_item_has_no_scenario_at_mi(mi)
                    and mi == cancel_mi
                    and lr < cancel_step
                    and self._master_step_snapshot_key(mi, lr)
                    in self._master_step_snapshots
                ):
                    return True
        return False

    def _value_grid_corner_background_for_state(self) -> str:
        if self._master_snapshot_browseable():
            return _DEBUG_RESULTS_SNAPSHOT_TINT_BG
        return _DEBUG_RESULTS_HEADER_BG

    def _apply_value_grid_header_and_corner_stylesheet(self) -> None:
        """結果一覧: 通常は列見出し・行番号とも薄灰。スナップショット表示中は列見出し（票の項目名）と行番号列を薄青、角は閲覧可能時のみ薄青。"""
        hdr = _DEBUG_RESULTS_HEADER_BG
        corner = self._value_grid_corner_background_for_state()
        snap_tint = self._master_showing_row_snapshot()
        h_hdr_bg = _DEBUG_RESULTS_SNAPSHOT_TINT_BG if snap_tint else hdr
        v_hdr_bg = _DEBUG_RESULTS_SNAPSHOT_TINT_BG if snap_tint else hdr
        h_sheet = (
            "QHeaderView::section {"
            "  background-color: %s;"
            "  color: #222;"
            "  padding: 3px;"
            "  border: 1px solid #d0d0d0;"
            "}"
            % h_hdr_bg
        )
        v_sheet = (
            "QHeaderView::section {"
            "  background-color: %s;"
            "  color: #222;"
            "  padding: 3px;"
            "  border: 1px solid #d0d0d0;"
            "}"
            % v_hdr_bg
        )
        corner_sheet = (
            "QTableCornerButton::section {"
            "  background-color: %s;"
            "  border: 1px solid #d0d0d0;"
            "}"
            % corner
        )
        try:
            self.value_grid.horizontalHeader().setStyleSheet(h_sheet)
            self.value_grid.verticalHeader().setStyleSheet(v_sheet)
            self.value_grid.setStyleSheet(corner_sheet)
        except Exception:
            pass

    def _apply_values_title_snapshot_tint(self) -> None:
        """結果一覧タイトル行（QLabel）は背景色を付けない（スナップショット示唆は列見出し・行番号列。左一覧は登録行スタイルのみ）。"""
        try:
            self.values_title.setStyleSheet("")
        except Exception:
            pass

    def _master_left_row_run_active(self) -> bool:
        """マスタ一覧: 実行中（連続／ステップ／進捗ロック）か。"""
        if self._mode != 1:
            return False
        return bool(
            getattr(self, "_continuous_busy", False)
            or getattr(self, "_master_step_loop_busy", False)
            or getattr(self, "_debug_progress_locked", False)
            or self._run_progress_dialog_blocking()
        )

    def _master_left_row_executing(self, row_idx: int) -> bool:
        """登録行のうち、現在実行中のマスタ項目行か。"""
        return self._master_left_row_run_active() and int(row_idx) == int(self._mi_idx)

    def _master_left_row_selected(self, row_idx: int) -> bool:
        """登録行のうち、現在選択中のマスタ項目行か（実行中でなくても濃色）。"""
        return int(row_idx) == int(self._mi_idx)

    def _apply_master_left_registered_row_style(self) -> None:
        """マスタ一覧: 登録行は薄ベージュ。選択中（実行中含む）は濃いベージュ＋項目名は青太字。"""
        if self._mode != 1:
            return
        try:
            mit = self._master_table_items()
            nr = self.left_table.rowCount()
            nc = self.left_table.columnCount()
            clear_bg = QBrush()
            default_fg = QBrush()
            beige = QBrush(_DEBUG_MASTER_REGISTERED_ROW_BG)
            active_beige = QBrush(_DEBUG_MASTER_ACTIVE_ROW_BG)
            blue = QBrush(_DEBUG_MASTER_REGISTERED_NAME_COLOR)
            for ri in range(nr):
                reg = (
                    ri < len(mit)
                    and self._master_active_count_for_item(mit[ri]) > 0
                )
                selected = reg and self._master_left_row_selected(ri)
                for ci in range(nc):
                    it = self.left_table.item(ri, ci)
                    if it is None:
                        continue
                    if reg:
                        it.setBackground(active_beige if selected else beige)
                        if ci == 1:
                            f = it.font()
                            f.setBold(True)
                            it.setFont(f)
                            it.setForeground(blue)
                        elif ci == 2:
                            f = it.font()
                            f.setBold(False)
                            it.setFont(f)
                            it.setForeground(default_fg)
                            it.setTextAlignment(
                                Qt.AlignmentFlag.AlignRight
                                | Qt.AlignmentFlag.AlignVCenter
                            )
                        else:
                            f = it.font()
                            f.setBold(False)
                            it.setFont(f)
                            it.setForeground(default_fg)
                    else:
                        it.setBackground(clear_bg)
                        f = it.font()
                        f.setBold(False)
                        it.setFont(f)
                        it.setForeground(default_fg)
        except Exception:
            pass

    def _refresh_master_snapshot_chrome(self) -> None:
        """スナップショット閲覧状態に応じて結果一覧の角・列／行見出しを更新し、マスタ一覧の登録行スタイルを再適用する。"""
        self._apply_value_grid_header_and_corner_stylesheet()
        self._apply_values_title_snapshot_tint()
        self._apply_master_left_registered_row_style()

    def _build_tab_log(self) -> None:
        lay = QVBoxLayout(self.tab_log)
        lay.setSpacing(2)
        lay.setContentsMargins(0, 0, 0, 0)
        self.log_intro = QLabel(
            self._d(
                "LOG_INTRO_HTML",
                "<b>ログ</b>：実行の要約、<b>EVENT</b> 行（§10.8：時刻・理由コード・シナリオID・参照パス・詳細）。",
            )
        )
        self._wrap_desc_label(self.log_intro)
        self.log_intro.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.log_intro)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(
            "font-family: 'MS Gothic', 'Yu Gothic UI', 'Consolas', monospace; "
            "font-size: 11px;"
        )
        lay.addWidget(self.log, 1)

    def _scenario_slots(self) -> list[dict[str, Any] | None]:
        return self._current_scenario()["slots"]

    def _on_mode_changed(self) -> None:
        if self._fixed_mode is not None:
            return
        if self._mode == 0:
            self._persist_scenario_state()
        self._mode = self.mode_combo.currentIndex()
        if self._mode != 0:
            self._scenario_snapshots.clear()
        self._apply_mode()
        self._full_reset(False)
        self._rebuild_after_reset()

    def _apply_mode(self) -> None:
        mvr = self._max_value_rows()
        if self._mode == 0:
            self.left_title.setText(self._d("LEFT_TITLE_SCENARIO_HTML", "<b>登録シナリオ</b>"))
            self.btn_run.setText(self._d("BTN_RUN_SCENARIO", "シナリオ実行"))
            sc_hint = self._d("HINT_SCENARIO_HTML", "") or (
                "<b>フェーズステップ</b>：実行済み行に薄いグレイ。"
                " シナリオ切替では結果・ログを保持。各フェーズの取得は最大<b>%d件</b>で打ち切り。"
                " キャンセル＝同一シナリオの結果のみクリア（ログは残る）。"
                % mvr
            )
            if not self._live_items:
                sc_hint = self._d(
                    "EMPTY_DEBUG_HINT_HTML",
                    "<p style='color:#444;'><b>ヒント</b>：試作デモデータは同梱していません。"
                    " メインで項目・取得シナリオを設定するか、シナリオ編集の「デバッグ」から開いてください。</p>",
                ) + sc_hint
            self.hint.setText(sc_hint)
            self.res_hint.setText(
                self._d("RES_HINT_SCENARIO_HTML", "")
                or "<b>結果サマリ</b>：列見出しは2行表示。先頭列はシナリオ名（マスタ）／条件ステップ（シナリオ）。"
            )
            self.values_title.setText(
                self._d("VALUES_TITLE_SCENARIO_HTML", "")
                or "<b>結果一覧</b>：フェーズ列は右へ増加。行番号は左端ヘッダのみ。"
            )
            self.btn_run_all_master.setVisible(False)
            self.btn_run_all.setText(self._btn_run_all_label())
            self._refresh_master_snapshot_chrome()
        else:
            self.left_title.setText(
                self._d("LEFT_TITLE_MASTER_HTML", "<b>項目一覧</b>")
            )
            self.btn_run.setText(self._d("BTN_RUN_MASTER", "シナリオ実行"))
            self.btn_run_all.setText(self._btn_run_all_label())
            self.btn_run_all_master.setVisible(True)
            mpd = self._master_preview_display_rows()
            ma_hint = self._d("HINT_MASTER_HTML", "") or (
                "<b>マスタ項目ステップ</b>：最終項目まで実行すると周回せず待機します（スナップショット閲覧が可能なときは左上コーナーが薄青）。"
                " 本番経路の一覧を再計算し、表示上限<b>%d行</b>で表示。"
                % (mpd,)
            )
            if not self._live_items:
                ma_hint = self._d(
                    "EMPTY_DEBUG_HINT_HTML",
                    "<p style='color:#444;'><b>ヒント</b>：試作デモデータは同梱していません。"
                    " メインで項目を定義してから「デバッグ」を開いてください。</p>",
                ) + ma_hint
            extra_hint = self._d(
                "HINT_MASTER_ROW_GAP_HTML",
                "",
            )
            if extra_hint:
                ma_hint += extra_hint
            self.hint.setText(ma_hint)
            self._update_master_res_hint()
            self._update_values_title_master()
        self._update_debug_window_title()
        self._apply_mode_dependent_tooltips()

    def _update_values_title_master(self) -> None:
        if self._mode != 1:
            return
        m = self._current_master()
        fmt = self._d(
            "VALUES_TITLE_MASTER_FMT",
            "<b>結果一覧</b>　項目名：<b>%s</b>　— 本番と同一の抽出ロジック（マスタシートへは書込まない）",
        )
        title = fmt % m["title"]
        suffix = self._master_values_title_rows_suffix()
        if suffix:
            title = "%s　%s" % (title, suffix)
        self.values_title.setText(title)
        self._refresh_master_snapshot_chrome()
        self._set_tip(
            self.values_title,
            "TIP_VALUES_TITLE",
            "結果一覧グリッドの見出しです。表示中の項目名などが含まれます。",
        )

    def _live_item_for_scenario_index(self, idx: int) -> dict[str, Any]:
        if 0 <= idx < len(self._live_items):
            return self._live_items[idx]
        return {}

    def _full_reset(self, keep_selection: bool) -> None:
        self._cancel_scenario_link_prefetch(join=False)
        self._bump_mpv_prefetch_cancel()
        self._mpv_close_item_wb_frame()
        self._mpv_try_flush_pending_wb_frames()
        self._mpv_invalidate_final_table_rows()
        self._scenario_bundle_caches.clear()
        self._master_sparse_notice_shown = False
        self._mpv_extract_cache.clear()
        self._mpv_colvals_cache.clear()
        self._mpv_progress_rows_cache = None
        self._mpv_last_valid_table_rows = []
        self._mpv_last_stats_files_read = 0
        self._mpv_last_stats_read_rows = 0
        self._mpv_last_stats_scan_cap_hit = False
        self._mpv_join_compute_busy = 0
        self._mpv_progress_rows_step_cache.clear()
        self._mpv_progress_rows_by_mi.clear()
        self._mpv_frozen_snapshots.clear()
        self._mpv_progress_row_peak_by_mi.clear()
        self._mpv_join_search_pool_seed = None
        self._mpv_join_search_pool_seed_paths_count = -1
        self._mpv_join_pool_by_mi.clear()
        self._mpv_row_file_paths_by_mi.clear()
        self._mpv_final_table_rows = None
        self._last_master_completed_mi_idx = None
        self._mpv_display_mi_idx = None
        self._mpv_deferred_value_grid_mi = None
        self._mpv_column_fit_pending = False
        self._mpv_final_grid_applied = False
        if self._mode == 1:
            if self._live_items:
                self._master_items_override = build_master_items_live(
                    self._live_items,
                    self._debug_scan_paths,
                    self._max_value_rows(),
                    preload_values=False,
                )
            else:
                self._master_items_override = _empty_debug_master_items()
        else:
            self._master_items_override = None
        self._phase_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._last_master_active_count = 0
        self._master_exec_armed = False
        self._master_global_row_idx = 0
        self._summary_rows.clear()
        self._value_cols.clear()
        self._value_col_tooltips.clear()
        self._value_col_spans.clear()
        self._summary_phase_labels.clear()
        self._master_full_continuous_allowed = True
        self._master_step_pass_complete = False
        if not keep_selection:
            self._sc_idx = 0
            self._mi_idx = 0
        self.log.clear()
        self.summary_table.setRowCount(0)
        self._reset_value_grid()
        self._reload_left_table()
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._rebuild_value_grid()
        self._paint_result_highlights()
        self._update_run_buttons_state()
        self._update_clear_buttons()
        if self._mode == 0:
            self._scenario_snapshots.clear()
            self._load_scenario_state(self._sc_idx)

    def _rebuild_after_reset(self) -> None:
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._rebuild_value_grid()
        self._paint_result_highlights()
        self._update_run_buttons_state()
        self._update_clear_buttons()
        if self._mode == 0:
            self._load_scenario_state(self._sc_idx)

    def _current_scenario(self) -> dict[str, Any]:
        return self._scenarios_data[self._sc_idx]

    def _master_table_items(self) -> list[dict[str, Any]]:
        if self._master_items_override is not None:
            return self._master_items_override
        return _empty_debug_master_items()

    def _current_master(self) -> dict[str, Any]:
        items = self._master_table_items()
        return items[self._mi_idx]

    def _scenario_progress_window_title(self) -> str:
        """進捗ダイアログのタイトルバー用。マスタ時は項目名と実行中シナリオ名を含める。"""
        if self._mode == 1:
            try:
                m = self._current_master()
                mt = str(m.get("title") or "").strip()
                acts = self._active_slot_indices
                if self._master_step_idx < len(acts):
                    si = acts[self._master_step_idx]
                    scenarios = m.get("scenarios") or []
                    if isinstance(scenarios, list) and si < len(scenarios):
                        sc = scenarios[si]
                        if isinstance(sc, dict):
                            st = str(sc.get("title") or "").strip()
                            if mt and st:
                                return "%s: %s" % (mt, st)
                            if st:
                                return st
                if mt:
                    return mt
            except Exception:
                pass
            return "-"
        scen = self._scenario_for_dry_run or {}
        for key in ("name", "id"):
            v = str(scen.get(key) or "").strip()
            if v:
                return v
        try:
            t = str(self._current_scenario().get("title") or "").strip()
            if t:
                return t
        except Exception:
            pass
        return "-"

    def _scenario_wants_file_progress(self, phase_gi: int, n_paths: int) -> bool:
        """シナリオモードは常時進捗表示し、連携・結合フェーズでは常にファイル単位更新も出す。"""
        _ = n_paths
        return int(phase_gi) in (3, 4)

    def _show_scenario_step_progress_start(
        self,
        phase_gi: int,
        phase_label: str,
        *,
        detail: str = "",
    ) -> None:
        wt = self._scenario_progress_window_title()
        show_detail = str(detail or "").strip()
        if not show_detail:
            show_detail = str(phase_label or "").strip()
        self._show_run_progress(
            "シナリオステップ実行中",
            max(1, int(phase_gi) + 1),
            max(1, len(self._active_slot_indices or [])),
            window_title=wt,
            detail=show_detail,
        )

    def _scenario_file_progress_phase_message(self, phase_gi: int) -> str:
        return (
            _SCENARIO_PROGRESS_PHASE_MSGLINK
            if int(phase_gi) == 3
            else _SCENARIO_PROGRESS_PHASE_MSGJOIN
        )

    def _scenario_make_file_progress_hook(
        self, phase_msg: str
    ) -> Callable[[int, int], None]:
        wt = self._scenario_progress_window_title()

        def hook(done: int, total: int) -> None:
            n = max(1, int(total))
            d = max(0, min(n, int(done)))
            self._show_run_progress(
                phase_msg,
                d,
                n,
                window_title=wt,
                detail="ファイル %s/%s" % (d, n) if n > 0 else "",
            )

        return hook

    def _reload_left_table(self) -> None:
        self.left_table.blockSignals(True)
        try:
            if self._mode == 0:
                self.left_table.setColumnCount(1)
                self.left_table.setHorizontalHeaderLabels(
                    [self._d("LEFT_TABLE_HEADER_SCENARIO", "シナリオ")]
                )
                self.left_table.setRowCount(len(self._scenarios_data))
                dn = _ne_detail_name_cfg()
                dcell = _ne_detail_cell_cfg()
                for i, s in enumerate(self._scenarios_data):
                    it = QTableWidgetItem(s["title"])
                    src = s.get("source")
                    if isinstance(src, dict):
                        it.setToolTip(
                            _normalize_tooltip_text(
                                scenario_source_tooltip_plain(src, dn, detail_cell_cfg=dcell)
                            )
                        )
                    self.left_table.setItem(i, 0, it)
                self.left_table.selectRow(self._sc_idx)
                _lh0 = self.left_table.horizontalHeader()
                _lh0.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            else:
                self.left_table.setColumnCount(3)
                self.left_table.setHorizontalHeaderLabels(
                    [
                        self._d("LEFT_TABLE_COL_INDEX", "#"),
                        self._d("LEFT_TABLE_HEADER_MASTER", "項目名"),
                        self._d("LEFT_TABLE_COL_ELAPSED", "時間(秒)"),
                    ]
                )
                mit = self._master_table_items()
                self.left_table.setRowCount(len(mit))
                for i, m in enumerate(mit):
                    self.left_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                    self.left_table.setItem(
                        i, 1, QTableWidgetItem(str(m.get("title") or ""))
                    )
                    reg = self._master_active_count_for_item(m) > 0
                    it_el = QTableWidgetItem(
                        self._master_format_elapsed_sec(
                            self._master_item_elapsed_sec.get(int(i))
                        )
                        if reg
                        else ""
                    )
                    it_el.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                    self.left_table.setItem(i, 2, it_el)
                self.left_table.selectRow(self._mi_idx)
                _lh = self.left_table.horizontalHeader()
                _lh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
                _lh.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
                _lh.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        finally:
            self.left_table.blockSignals(False)
        if self._mode == 1:
            self._apply_master_left_registered_row_style()
        self._update_left_detail()

    def _update_left_detail(self) -> None:
        if self._mode == 1:
            self._update_values_title_master()

    def _on_left_sel(self) -> None:
        r = self.left_table.currentRow()
        if r < 0:
            return
        if self._mode == 0:
            if r != self._sc_idx:
                self._cancel_scenario_link_prefetch(join=False)
                self._persist_scenario_state()
                self._sc_idx = r
                self._load_scenario_state(self._sc_idx)
        else:
            if r != self._mi_idx:
                self._bump_mpv_prefetch_cancel()
                self._mpv_deferred_value_grid_mi = None
                # 項目切替時は workbook キャッシュも破棄（離脱キャプチャを経ない操作）
                self._mpv_close_item_wb_frame()
                self._mi_idx = r
                self._master_sparse_notice_shown = False
                # 選択マスタ項目の先頭シナリオ（アクティブスロットの先頭）へ。一覧の行番号ではない。
                self._master_step_idx = 0
                self._master_session_start_step = 0
                self._master_exec_armed = False
                self._master_global_row_idx = self._master_global_row_base_for_mi(self._mi_idx)
        self._reload_conditions()
        if self._mode == 0:
            self._rebuild_active_slots()
            self._rebuild_left_steps()
            self._paint_left_steps_executed()
            self._rebuild_value_grid()
        else:
            self._recompute_active_slot_indices()
            self._rebuild_left_steps()
            self._paint_left_steps_executed()
            if getattr(self, "_master_snapshot_browse_after_cancel", False):
                self._apply_master_browse_snapshot_for_mi(r)
            elif self._master_snapshot_priority_active() and r in self._master_item_snapshots:
                self._apply_master_item_snapshot(r)
            else:
                self._sync_summary_table_from_lists()
                self._rebuild_value_grid()
            if getattr(self, "_master_snapshot_browse_after_cancel", False) or (
                self._master_snapshot_priority_active()
                and r in self._master_item_snapshots
            ):
                self._apply_selected_master_step_snapshot_if_any()
        if self._mode == 1:
            self._update_left_detail()
        self._paint_result_highlights()
        self._update_run_buttons_state()
        self._update_clear_buttons()
        if self._mode == 1:
            self._apply_master_left_registered_row_style()

    def _clear_master_item_snapshots(self) -> None:
        self._master_item_snapshots.clear()
        self._master_item_snapshot_done.clear()
        self._master_step_snapshots.clear()
        self._master_snapshot_browse_after_cancel = False
        self._master_cancel_mi = 0
        self._master_cancel_step = 0
        self._refresh_master_snapshot_chrome()

    def _master_snapshot_priority_active(self) -> bool:
        if self._mode != 1:
            return False
        items = self._master_table_items()
        if not items:
            return False
        if getattr(self, "_master_snapshot_browse_after_cancel", False):
            return bool(
                self._master_item_snapshots or self._master_step_snapshots
            )
        return len(self._master_item_snapshot_done) >= len(items)

    def _master_has_step_snapshots_for_mi(self, mi: int) -> bool:
        mi_i = int(mi)
        return any(k[0] == mi_i for k in self._master_step_snapshots)

    def _master_item_has_no_scenario_at_mi(self, mi: int) -> bool:
        items = self._master_table_items()
        if mi < 0 or mi >= len(items):
            return True
        return self._master_active_count_for_item(items[mi]) <= 0

    def _master_first_scenario_registered_mi(self) -> int:
        """シナリオが登録されている先頭マスタ項目の行番号。"""
        for i, m in enumerate(self._master_table_items()):
            if self._master_active_count_for_item(m) > 0:
                return int(i)
        return 0

    def _apply_master_empty_display(self) -> None:
        self._summary_rows.clear()
        self._summary_phase_labels.clear()
        self._value_cols.clear()
        self._value_col_tooltips.clear()
        self._value_col_spans.clear()
        self._sync_summary_table_from_lists()
        self._reset_value_grid()
        self._mpv_join_table_active = False
        self._mpv_join_table_ncols = 0
        self._mpv_apply_item_stats_snapshot(None)
        self._update_values_title_master()
        self._refresh_master_snapshot_chrome()
        self._paint_result_highlights()

    def _apply_master_browse_snapshot_for_mi(self, mi: int) -> None:
        """キャンセル後閲覧: 完了項目・ステップのみスナップショット、それ以外は空欄。"""
        mi_i = int(mi)
        if self._master_item_has_no_scenario_at_mi(mi_i):
            self._apply_master_empty_display()
            return
        cancel_mi = int(getattr(self, "_master_cancel_mi", mi_i))
        if mi_i > cancel_mi:
            self._apply_master_empty_display()
            return
        if mi_i < cancel_mi:
            if mi_i in self._master_item_snapshots:
                self._apply_master_item_snapshot(mi_i)
            else:
                self._apply_master_empty_display()
            return
        if not self._apply_selected_master_step_snapshot_if_any():
            self._apply_master_empty_display()

    def _master_snapshot_browse_step_for_mi(self, mi: int) -> int:
        """閲覧開始時に選択するステップ行（その項目で最後に完了したステップ）。"""
        mi_i = int(mi)
        cancel_mi = int(getattr(self, "_master_cancel_mi", 0))
        cancel_step = int(getattr(self, "_master_cancel_step", 0))
        if mi_i < cancel_mi:
            best = -1
            for k in self._master_step_snapshots:
                if k[0] == mi_i and int(k[1]) > best:
                    best = int(k[1])
            return max(0, best)
        if mi_i == cancel_mi:
            return max(0, cancel_step - 1)
        return 0

    def _focus_master_snapshot_browse_first_item(self) -> int:
        """キャンセル後閲覧: シナリオ登録済みの先頭マスタ項目を選択する。"""
        browse_mi = self._master_first_scenario_registered_mi()
        if int(self._mi_idx) != browse_mi:
            self._bump_mpv_prefetch_cancel()
            self._mpv_deferred_value_grid_mi = None
            self._mi_idx = browse_mi
            self._master_step_idx = 0
            self._master_session_start_step = 0
            self._master_exec_armed = False
            self._master_global_row_idx = self._master_global_row_base_for_mi(browse_mi)
            if self.left_table.rowCount() > browse_mi:
                self.left_table.blockSignals(True)
                try:
                    self.left_table.selectRow(browse_mi)
                finally:
                    self.left_table.blockSignals(False)
            self._update_left_detail()
            self._reload_conditions()
            self._recompute_active_slot_indices()
            self._rebuild_left_steps()
            self._paint_left_steps_executed()
        step_row = self._master_snapshot_browse_step_for_mi(browse_mi)
        if self.left_steps.rowCount() > 0 and step_row >= 0:
            wr = min(step_row, self.left_steps.rowCount() - 1)
            self.left_steps.blockSignals(True)
            try:
                self.left_steps.selectRow(wr)
            finally:
                self.left_steps.blockSignals(False)
        return browse_mi

    def _flush_deferred_master_value_grid_if_mi(self, mi: int) -> None:
        dm = getattr(self, "_mpv_deferred_value_grid_mi", None)
        if dm is None:
            return
        if int(dm) != int(mi):
            return
        self._mpv_deferred_value_grid_mi = None
        self._rebuild_value_grid()

    def _capture_master_leave_item(self, completed_mi: int, *, empty: bool = False) -> None:
        if completed_mi < 0:
            return
        # 項目単位 workbook キャッシュは離脱時に破棄（次項目へ持ち越さない）
        wb_mi = getattr(self, "_mpv_item_wb_mi", None)
        if wb_mi is not None and int(wb_mi) == int(completed_mi):
            self._mpv_close_item_wb_frame()
            self._mpv_try_flush_pending_wb_frames()
        if empty:
            self._master_item_snapshots[completed_mi] = {"empty": True}
        else:
            gh: list[str] = []
            gr: list[list[str]] = []
            used_step_cache = False
            if self._scenario_for_dry_run and self._debug_scan_paths:
                items = list((self._scenario_for_dry_run or {}).get("items") or [])
                gh = [
                    str(it.get("name") or it.get("id") or ("項目_%s" % i))
                    for i, it in enumerate(items)
                ]
                mi_saved = int(self._mi_idx)
                step_saved = int(self._master_step_idx)
                try:
                    self._mi_idx = int(completed_mi)
                    self._rebuild_active_slots()
                    n_act = len(self._active_slot_indices or [])
                    if n_act > 0:
                        gr = self._mpv_table_rows_for_step_snapshot(
                            int(completed_mi),
                            max(0, int(n_act) - 1),
                        )
                        used_step_cache = self._mpv_table_rows_have_data(gr)
                finally:
                    self._mi_idx = mi_saved
                    self._master_step_idx = step_saved
                    self._rebuild_active_slots()
            if not used_step_cache:
                gh = []
                nc = self.value_grid.columnCount()
                for c in range(nc):
                    hi = self.value_grid.horizontalHeaderItem(c)
                    gh.append("" if hi is None else str(hi.text()))
                gr = []
                for r in range(self.value_grid.rowCount()):
                    row: list[str] = []
                    for c in range(nc):
                        it = self.value_grid.item(r, c)
                        row.append("" if it is None else str(it.text()))
                    gr.append(row)
            self._master_item_snapshots[completed_mi] = {
                "empty": False,
                "summary_rows": copy.deepcopy(self._summary_rows),
                "summary_phase_labels": copy.deepcopy(self._summary_phase_labels),
                "value_cols": copy.deepcopy(self._value_cols),
                "value_col_tooltips": copy.deepcopy(self._value_col_tooltips),
                "value_col_spans": copy.deepcopy(self._value_col_spans),
                "grid_headers": gh,
                "grid_rows": gr,
                "item_stats": self._mpv_item_stats_for_snapshot(),
            }
        self._master_item_snapshot_done.add(completed_mi)
        self._refresh_master_snapshot_chrome()

    def _master_step_snapshot_key(self, mi: int, step_row: int) -> tuple[int, int]:
        return (int(mi), int(step_row))

    def _mpv_stringify_table_rows(
        self, rows: list[list[Any]], *, ncols: int
    ) -> list[list[str]]:
        out: list[list[str]] = []
        for r in rows:
            rr = list(r)
            if len(rr) < ncols:
                rr.extend([None] * (ncols - len(rr)))
            out.append(["" if v is None else str(v) for v in rr[:ncols]])
        return out

    def _mpv_table_rows_have_data(self, rows: list[list[str]]) -> bool:
        return bool(rows) and any(
            any(str(v).strip() for v in row) for row in rows
        )

    def _mpv_publish_table_rows(
        self,
        rows: list[list[Any]],
        *,
        mi_idx: int | None = None,
    ) -> None:
        if not rows:
            return
        copied = [list(r) for r in rows]
        key = self._mpv_progress_cache_key()
        self._mpv_progress_rows_cache = (key, copied)
        self._mpv_last_valid_table_rows = [list(r) for r in copied]
        self._mpv_grid = [list(r) for r in copied]
        if mi_idx is not None:
            self._mpv_display_mi_idx = int(mi_idx)

    def _mpv_table_rows_for_step_snapshot(
        self, mi: int, step_row: int
    ) -> list[list[str]]:
        """ステップ snapshot 用。compute の table_rows のみ（extract overlay 禁止）。"""
        if not (self._scenario_for_dry_run and self._debug_scan_paths):
            return []
        items = list((self._scenario_for_dry_run or {}).get("items") or [])
        ncols = len(items)
        if ncols <= 0 or int(mi) < 0 or int(step_row) < 0:
            return []
        n_pick = int(step_row) + 1
        rows: list[list[Any]] | None = None
        if n_pick > 0:
            rows = self._mpv_rows_from_step_cache_n_pick(n_pick)
        if not rows:
            cache_ent = self._mpv_progress_rows_cache
            if isinstance(cache_ent, tuple) and len(cache_ent) == 2 and cache_ent[1]:
                rows = [list(r) for r in cache_ent[1]]
        if not rows and self._mpv_last_valid_table_rows:
            rows = [list(r) for r in self._mpv_last_valid_table_rows]
        return self._mpv_stringify_table_rows(list(rows or []), ncols=ncols)

    def _enter_master_snapshot_browse_after_cancel(self) -> None:
        if self._mode != 1:
            return
        self._mpv_close_item_wb_frame()
        self._mpv_try_flush_pending_wb_frames()
        if not (self._master_item_snapshots or self._master_step_snapshots):
            return
        self._master_cancel_mi = int(self._mi_idx)
        self._master_cancel_step = int(self._master_step_idx)
        self._master_snapshot_browse_after_cancel = True
        browse_mi = self._focus_master_snapshot_browse_first_item()
        self._apply_master_browse_snapshot_for_mi(browse_mi)
        self._refresh_master_snapshot_chrome()
        self._paint_result_highlights()
        self._update_run_buttons_state()
        self._update_clear_buttons()

    def _capture_master_step_snapshot(
        self,
        mi: int,
        step_row: int,
        *,
        colvals: list[str] | None = None,
    ) -> None:
        del colvals  # snapshot は table_rows のみ。extract 主値は使わない。
        if mi < 0 or step_row < 0:
            return
        headers: list[str] = []
        grid_rows: list[list[str]] = []
        if self._scenario_for_dry_run and self._debug_scan_paths:
            items = list((self._scenario_for_dry_run or {}).get("items") or [])
            headers = [
                str(it.get("name") or it.get("id") or ("項目_%s" % i))
                for i, it in enumerate(items)
            ]
            n_pick = int(step_row) + 1
            n_act = len(self._active_slot_indices or [])
            if n_pick > 0 and n_act > 0:
                n_pick_eff = min(n_pick, n_act)
                if not self._mpv_rows_from_step_cache_n_pick(n_pick_eff):
                    self._mpv_ensure_step_n_pick_cached(
                        n_pick=n_pick_eff,
                        progress_hook=self._mpv_master_dbg_progress_hook_or_none(),
                        probe_caller="mpv_step_snapshot",
                    )
                self._mpv_sync_progress_cache_from_step_n_pick(n_pick_eff)
            grid_rows = self._mpv_table_rows_for_step_snapshot(int(mi), int(step_row))
            if grid_rows and self._mpv_table_rows_have_data(grid_rows):
                any_rows = [
                    [None if str(v).strip() == "" else v for v in row]
                    for row in grid_rows
                ]
                self._mpv_publish_table_rows(any_rows, mi_idx=int(mi))
            else:
                grid_rows = []
        else:
            nc = self.value_grid.columnCount()
            for c in range(nc):
                hi = self.value_grid.horizontalHeaderItem(c)
                headers.append("" if hi is None else str(hi.text()))
            for r in range(self.value_grid.rowCount()):
                row: list[str] = []
                for c in range(nc):
                    it = self.value_grid.item(r, c)
                    row.append("" if it is None else str(it.text()))
                grid_rows.append(row)
        prev = self._master_step_snapshots.get(self._master_step_snapshot_key(mi, step_row))
        if not self._mpv_table_rows_have_data(grid_rows):
            if isinstance(prev, dict) and prev.get("grid_rows"):
                return
        self._master_step_snapshots[self._master_step_snapshot_key(mi, step_row)] = {
            "summary_rows": copy.deepcopy(self._summary_rows),
            "summary_phase_labels": copy.deepcopy(self._summary_phase_labels),
            "value_cols": copy.deepcopy(self._value_cols),
            "value_col_tooltips": copy.deepcopy(self._value_col_tooltips),
            "value_col_spans": copy.deepcopy(self._value_col_spans),
            "grid_headers": headers,
            "grid_rows": grid_rows,
            "item_stats": self._mpv_item_stats_for_snapshot(),
        }

    def _apply_master_step_snapshot(self, mi: int, step_row: int) -> bool:
        snap = self._master_step_snapshots.get(self._master_step_snapshot_key(mi, step_row))
        if snap is None:
            return False
        self._summary_rows = copy.deepcopy(snap["summary_rows"])
        self._summary_phase_labels = copy.deepcopy(snap["summary_phase_labels"])
        self._value_cols = copy.deepcopy(snap["value_cols"])
        self._value_col_tooltips = copy.deepcopy(snap["value_col_tooltips"])
        self._value_col_spans = copy.deepcopy(snap["value_col_spans"])
        self._sync_summary_table_from_lists()
        self._apply_value_grid_from_snapshot(
            list(snap["grid_headers"]), list(snap["grid_rows"])
        )
        any_rows = [
            [None if str(v).strip() == "" else v for v in row]
            for row in snap.get("grid_rows") or []
        ]
        if any_rows:
            self._mpv_publish_table_rows(any_rows, mi_idx=int(mi))
            self._mpv_grid = [list(r) for r in any_rows]
        self._mpv_join_table_active = bool(snap.get("grid_headers"))
        self._mpv_join_table_ncols = len(snap.get("grid_headers") or [])
        self._mpv_apply_item_stats_snapshot(snap.get("item_stats"))
        self._update_values_title_master()
        self._refresh_master_snapshot_chrome()
        return True

    def _apply_selected_master_step_snapshot_if_any(self) -> bool:
        if self._mode != 1:
            return False
        lr = self.left_steps.currentRow()
        if lr < 0:
            return False
        mi = int(self._mi_idx)
        step_row = int(lr)
        if getattr(self, "_master_snapshot_browse_after_cancel", False):
            if self._master_item_has_no_scenario_at_mi(mi):
                self._apply_master_empty_display()
                return True
            cancel_mi = int(getattr(self, "_master_cancel_mi", mi))
            if mi > cancel_mi:
                self._apply_master_empty_display()
                return True
            if mi < cancel_mi:
                key = self._master_step_snapshot_key(mi, step_row)
                if key in self._master_step_snapshots:
                    return self._apply_master_step_snapshot(mi, step_row)
                if mi in self._master_item_snapshots:
                    self._apply_master_item_snapshot(mi)
                    return True
                self._apply_master_empty_display()
                return True
            if step_row >= int(getattr(self, "_master_cancel_step", 0)):
                self._apply_master_empty_display()
                return True
        return self._apply_master_step_snapshot(mi, step_row)

    def _apply_value_grid_from_snapshot(self, headers: list[str], rows: list[list[str]]) -> None:
        self._mpv_join_table_active = False
        self._mpv_join_table_ncols = 0
        disp_headers = self._decorate_debug_grid_headers([str(h) for h in headers])
        nc = len(disp_headers)
        nr = len(rows)
        self._value_grid_note_structure(disp_headers)
        self.value_grid.clear()
        self.value_grid.setColumnCount(max(0, nc))
        self.value_grid.setRowCount(max(0, nr))
        if nc > 0:
            self.value_grid.setHorizontalHeaderLabels(disp_headers)
        hdr_v = self.value_grid.horizontalHeader()
        hdr_v.setVisible(True)
        hdr_v.setMinimumHeight(32)
        for r in range(nr):
            self.value_grid.setVerticalHeaderItem(r, QTableWidgetItem(str(r + 1)))
        for r in range(nr):
            row = rows[r] if r < len(rows) else []
            for c in range(nc):
                tx = str(row[c]) if c < len(row) else ""
                cell = QTableWidgetItem(tx)
                cell.setToolTip(_normalize_tooltip_text(tx))
                self.value_grid.setItem(r, c, cell)
        for c in range(nc):
            hdr_v.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
        self._fit_value_grid_columns()
        self._paint_result_highlights()

    def _apply_master_item_snapshot(self, mi: int) -> None:
        snap = self._master_item_snapshots.get(mi)
        if snap is None:
            return
        if snap.get("empty"):
            self._summary_rows.clear()
            self._summary_phase_labels.clear()
            self._value_cols.clear()
            self._value_col_tooltips.clear()
            self._value_col_spans.clear()
            self._sync_summary_table_from_lists()
            self._reset_value_grid()
            self._paint_result_highlights()
            self._refresh_master_snapshot_chrome()
            return
        self._summary_rows = copy.deepcopy(snap["summary_rows"])
        self._summary_phase_labels = copy.deepcopy(snap["summary_phase_labels"])
        self._value_cols = copy.deepcopy(snap["value_cols"])
        self._value_col_tooltips = copy.deepcopy(snap["value_col_tooltips"])
        self._value_col_spans = copy.deepcopy(snap["value_col_spans"])
        self._sync_summary_table_from_lists()
        self._apply_value_grid_from_snapshot(
            list(snap["grid_headers"]), list(snap["grid_rows"])
        )
        self._mpv_apply_item_stats_snapshot(snap.get("item_stats"))
        self._update_values_title_master()
        self._refresh_master_snapshot_chrome()

    def _recompute_active_slot_indices(self) -> None:
        self._active_slot_indices.clear()
        if self._mode == 0:
            for i, slot in enumerate(self._scenario_slots()):
                if slot is not None and (i < 3 or bool(slot.get("defined", True))):
                    self._active_slot_indices.append(i)
        else:
            m = self._current_master()
            for si, sc in enumerate(m["scenarios"]):
                slot = sc.get("slot")
                if slot is not None and bool(slot.get("defined", True)):
                    self._active_slot_indices.append(si)
        self._active_slot_indices = self._active_slot_indices[: self._max_phase_slots()]

    def _refresh_master_nav_lock_state(self) -> None:
        if self._mode != 1:
            return
        locked = bool(getattr(self, "_debug_progress_locked", False))
        locked = locked or bool(getattr(self, "_continuous_busy", False))
        locked = locked or self._run_progress_dialog_blocking()
        try:
            self.left_table.setEnabled(not locked)
            self.btn_prev.setEnabled(not locked)
            self.btn_next.setEnabled(not locked)
            if self.mode_combo.isVisible():
                self.mode_combo.setEnabled(not locked)
        except Exception:
            pass
        self._apply_master_left_registered_row_style()

    def _rebuild_active_slots(self) -> None:
        self._recompute_active_slot_indices()
        self._sync_summary_table_from_lists()

    def _nav_left(self, delta: int) -> None:
        if self._mode == 0:
            idx = max(0, min(len(self._scenarios_data) - 1, self._sc_idx + delta))
            self.left_table.selectRow(idx)
        else:
            idx = max(0, min(len(self._master_table_items()) - 1, self._mi_idx + delta))
            self.left_table.selectRow(idx)

    def _display_step_no(self, row_in_left: int) -> int:
        if self._mode == 0:
            return row_in_left + 1
        return self._master_session_start_step + row_in_left + 1

    def _scenario_join_def_item_names(self) -> list[str]:
        """結合キー定義のマスタ項目名（重複は除く）。"""
        item = self._live_item_for_scenario_index(self._sc_idx)
        srcs = item.get("sources") or []
        s0 = srcs[0] if srcs and isinstance(srcs[0], dict) else None
        if not isinstance(s0, dict):
            return []
        p = source_ui_block(s0) or {}
        raw_jdefs = p.get("join_defs")
        jdefs: list[Any] = raw_jdefs if isinstance(raw_jdefs, list) else []
        seen: set[str] = set()
        out: list[str] = []
        for jd in jdefs:
            if not isinstance(jd, dict):
                continue
            nm = str(jd.get("item") or "").strip()
            if not nm or nm in seen:
                continue
            seen.add(nm)
            out.append(nm)
        return out

    def _join_phase_value_row_index(self) -> int | None:
        """セル座標系で結合キー（スロット index 4）に対応するサマリ／値列の行番号。"""
        if self._scenario_source_kind() != "cell":
            return None
        for r, gi in enumerate(self._active_slot_indices):
            if gi == 4:
                return r
        return None

    def _append_scenario_join_columns_if_needed(
        self,
        expanded: list[tuple[str, list[str], list[str | None] | None]],
    ) -> list[tuple[str, list[str], list[str | None] | None]]:
        """結合設定があるのに #n[項目] 展開で列が付いていない場合、結合項目名列を補完する。"""
        if self._mode != 0:
            return expanded
        names = self._scenario_join_def_item_names()
        if not names:
            return expanded
        hdrs = {h for h, _, _ in expanded}
        max_r = max((len(col) for _, col, _ in expanded), default=0)
        if max_r <= 0:
            max_r = 1
        join_row = self._join_phase_value_row_index()
        colvals: list[str] | None = None
        if join_row is not None and join_row < len(self._value_cols):
            colvals = list(self._value_cols[join_row])
        parsed = _parse_hash_bracket_column_values(colvals) if colvals else {}
        j_prefix = self._d("VALUE_COL_JOIN_PREFIX", "結合")
        out = list(expanded)
        for jname in names:
            if jname in hdrs:
                continue
            hdr = "%s:%s" % (j_prefix, jname)
            if hdr in hdrs:
                continue
            vals = parsed.get(jname) or []
            col: list[str] = []
            for r in range(max_r):
                col.append(str(vals[r]) if r < len(vals) else "—")
            if not vals and not colvals:
                col = ["—"] * max_r
            out.append((hdr, col, None))
            hdrs.add(hdr)
        return out

    def _rebuild_left_steps(self) -> None:
        self.left_steps.blockSignals(True)
        dn = _ne_detail_name_cfg()
        dcell = _ne_detail_cell_cfg()
        try:
            self.left_steps.setRowCount(0)
            if self._mode == 0:
                self.left_steps.setColumnCount(2)
                self.left_steps.setHorizontalHeaderLabels(
                    [
                        self._d("COND_COL_NO", "番号"),
                        self._d("COND_COL_ITEM", "大項目"),
                    ]
                )
                _ls_h = self.left_steps.horizontalHeader()
                _ls_h.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
                _ls_h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
                slots = self._scenario_slots()
                keys = self._cond_keys()
                cs = self._current_scenario()
                src0 = cs.get("source")
                for _li, gi in enumerate(self._active_slot_indices):
                    s = slots[gi]
                    assert s is not None
                    r = self.left_steps.rowCount()
                    self.left_steps.insertRow(r)
                    ckey = keys[gi] if gi < len(keys) else str(gi + 1)
                    it0 = QTableWidgetItem(str(self._display_step_no(r)))
                    it1 = QTableWidgetItem(ckey)
                    for it in (it0, it1):
                        it.setTextAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
                    if isinstance(src0, dict):
                        row_tip = scenario_source_tooltip_plain(src0, dn, detail_cell_cfg=dcell)
                    else:
                        row_tip = _format_condition_step_tooltip(ckey, s)
                    it0.setToolTip(_normalize_tooltip_text(row_tip))
                    it1.setToolTip(_normalize_tooltip_text(row_tip))
                    self.left_steps.setItem(r, 0, it0)
                    self.left_steps.setItem(r, 1, it1)
            else:
                self.left_steps.setColumnCount(3)
                self.left_steps.setHorizontalHeaderLabels(
                    [
                        self._d("COND_COL_NO", "番号"),
                        self._d("COND_COL_ITEM", "大項目"),
                        self._d("COND_COL_ELAPSED", "時間(秒)"),
                    ]
                )
                _ls_h = self.left_steps.horizontalHeader()
                _ls_h.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
                _ls_h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
                _ls_h.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
                m = self._current_master()
                mi = int(self._mi_idx)
                for _li, si in enumerate(self._active_slot_indices):
                    sc = m["scenarios"][si]
                    slot = sc["slot"]
                    assert slot is not None
                    r = self.left_steps.rowCount()
                    self.left_steps.insertRow(r)
                    step_txt = str(sc["title"] or "シナリオ")
                    it0 = QTableWidgetItem(str(self._display_step_no(r)))
                    it1 = QTableWidgetItem(step_txt)
                    it2 = QTableWidgetItem(
                        self._master_format_elapsed_sec(
                            self._master_step_elapsed_sec.get((mi, int(r)))
                        )
                    )
                    it2.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                    for it in (it0, it1, it2):
                        it.setTextAlignment(
                            Qt.AlignmentFlag.AlignTop
                            | (
                                Qt.AlignmentFlag.AlignRight
                                if it is it2
                                else Qt.AlignmentFlag.AlignLeft
                            )
                        )
                    src1 = sc.get("source")
                    if isinstance(src1, dict):
                        row_tip = scenario_source_tooltip_plain(src1, dn, detail_cell_cfg=dcell)
                    else:
                        row_tip = _format_condition_step_tooltip(step_txt, slot)
                    for it in (it0, it1, it2):
                        it.setToolTip(_normalize_tooltip_text(row_tip))
                    self.left_steps.setItem(r, 0, it0)
                    self.left_steps.setItem(r, 1, it1)
                    self.left_steps.setItem(r, 2, it2)
        finally:
            self.left_steps.blockSignals(False)
        self.left_steps.resizeRowsToContents()
        if self.left_steps.rowCount() > 0:
            wr = self._phase_idx if self._mode == 0 else self._master_step_idx
            self.left_steps.selectRow(min(wr, self.left_steps.rowCount() - 1))

    def _paint_left_steps_executed(self) -> None:
        for r in range(self.left_steps.rowCount()):
            for c in range(self.left_steps.columnCount()):
                it = self.left_steps.item(r, c)
                if not it:
                    continue
                it.setBackground(QBrush())

    def _on_step_sel(self) -> None:
        if self._mode == 1:
            self._apply_selected_master_step_snapshot_if_any()
            if getattr(self, "_master_snapshot_browse_after_cancel", False):
                self._update_run_buttons_state()
        self._paint_result_highlights()

    def _summary_row_for_left_row(self, left_row: int) -> int:
        if self._mode == 0:
            return left_row
        return self._master_session_start_step + left_row

    def _value_col_for_left_row(self, left_row: int) -> int:
        if self._mode == 0:
            return left_row
        return self._master_session_start_step + left_row

    def _value_col_span_for_left_row(self, left_row: int) -> tuple[int, int]:
        if (
            self._mode == 1
            and self._mpv_join_table_active
            and self._mpv_join_table_ncols > 0
        ):
            return (0, self._mpv_join_table_ncols - 1)
        vcol = self._value_col_for_left_row(left_row)
        if 0 <= vcol < len(self._value_col_spans):
            return self._value_col_spans[vcol]
        return (vcol, vcol)

    def _paint_result_highlights(self) -> None:
        for r in range(self.summary_table.rowCount()):
            for c in range(self.summary_table.columnCount()):
                it = self.summary_table.item(r, c)
                if it:
                    it.setBackground(QBrush())

        for r in range(self.value_grid.rowCount()):
            for c in range(self.value_grid.columnCount()):
                it = self.value_grid.item(r, c)
                if it:
                    it.setBackground(QBrush())

        if not (
            self._mode == 1
            and getattr(self, "_master_step_pass_complete", False)
        ):
            _phase_br = QBrush(_DEBUG_SUMMARY_PHASE_COL_BG)
            for r in range(self.summary_table.rowCount()):
                it0 = self.summary_table.item(r, 0)
                if it0 is not None:
                    it0.setBackground(_phase_br)

        self._style_results_table_header_rows()
        self._sync_value_grid_phase_dividers()

    def _value_grid_phase_start_columns(self) -> frozenset[int]:
        return phase_start_columns_from_spans(
            list(self._value_col_spans),
            int(self.value_grid.columnCount()),
            scenario_mode=self._mode == 0,
        )

    def _sync_value_grid_phase_dividers(self) -> None:
        starts = self._value_grid_phase_start_columns()
        dele = getattr(self, "_value_grid_delegate", None)
        if isinstance(dele, _ValueGridNoElideDelegate):
            dele.phase_start_cols = starts
        hdr = self.value_grid.horizontalHeader()
        if isinstance(hdr, _ValueGridPhaseHeader):
            hdr.phase_start_cols = starts
        try:
            self.value_grid.viewport().update()
            hdr.viewport().update()
        except Exception:
            pass

    def _tree_paint_parent_item_column(self, top: QTreeWidgetItem, col: int = 1) -> None:
        """親行の項目名列に背景色は付けない。"""
        if 0 <= col < top.columnCount():
            top.setBackground(col, QBrush())

    def _apply_cond_tree_item_align(self, it: QTreeWidgetItem) -> None:
        ta = Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        for c in range(it.columnCount()):
            it.setTextAlignment(c, ta)

    def _tree_lines_for_slot(self, slot: dict[str, Any]) -> list[str]:
        """1 行＝1 表示行（§3.1.3.1）。editor_lines 優先、なければ第3列要約を分割。"""
        el = slot.get("editor_lines")
        if isinstance(el, list) and el:
            out: list[str] = []
            for x in el:
                for s in str(x).split("\n"):
                    t = s.strip()
                    if t:
                        out.append(t)
            if out:
                return out
        txt = _slot_third_column(slot)
        return [s.strip() for s in txt.split("\n") if s.strip()]

    def _tree_add_slot_children(self, top: QTreeWidgetItem, slot: dict[str, Any]) -> None:
        lines = self._tree_lines_for_slot(slot)
        for ln in lines[1:]:
            ch = QTreeWidgetItem(["", "", ln])
            self._apply_cond_tree_item_align(ch)
            top.addChild(ch)
        if self._scenario_source_kind() == "name_extract":
            return
        # editor_lines があるときは同一内容が details にも出ることが多く二重になるため付けない
        el = slot.get("editor_lines")
        if isinstance(el, list) and el:
            return
        for k, v in slot.get("details", []):
            ch = QTreeWidgetItem(["", "", "%s：%s" % (k, v)])
            self._apply_cond_tree_item_align(ch)
            top.addChild(ch)

    def _reload_conditions(self) -> None:
        dn = _ne_detail_name_cfg()
        dcell = _ne_detail_cell_cfg()
        if self._mode == 0:
            self.cond_stack.setCurrentIndex(0)
            self.cond_hint.setText(
                self._d(
                    "COND_HINT_SCENARIO_HTML",
                    "<b>条件</b>：番号／項目／要約。親行の背景は<b>条件項目列のみ</b>（§3.1.3.2）。",
                )
            )
            self.cond_tree.clear()
            slots = self._scenario_slots()
            n = 0
            for gi, key in enumerate(self._cond_keys()):
                slot = slots[gi]
                if slot is None:
                    continue
                if gi >= 3 and not bool(slot.get("defined", True)):
                    continue
                n += 1
                lines = self._tree_lines_for_slot(slot)
                first_line = lines[0] if lines else ""
                top = QTreeWidgetItem([str(n), key, first_line])
                self._apply_cond_tree_item_align(top)
                self._tree_paint_parent_item_column(top, 1)
                step_tip = _format_condition_step_tooltip(key, slot)
                for col in range(3):
                    top.setToolTip(col, _normalize_tooltip_text(step_tip))
                self._tree_add_slot_children(top, slot)
                self.cond_tree.addTopLevelItem(top)
            self.cond_tree.collapseAll()
        else:
            self.cond_stack.setCurrentIndex(1)
            self.cond_hint.setText(
                self._d(
                    "COND_HINT_MASTER_HTML",
                    "<b>条件</b>：登録シナリオごと。親行の背景は<b>シナリオ名列のみ</b>。",
                )
            )
            self.master_cond_tree.clear()
            m = self._current_master()
            n = 0
            for si, sc in enumerate(m["scenarios"]):
                slot = sc.get("slot")
                if slot is None:
                    continue
                n += 1
                lines = self._tree_lines_for_slot(slot)
                first_line = lines[0] if lines else ""
                src_m = sc.get("source")
                # 親行は要約1行のみ（全文はツールチップ）。scenario_source_tooltip_plain を親に載せると子と二重になる。
                top = QTreeWidgetItem([str(n), sc["title"], first_line])
                self._apply_cond_tree_item_align(top)
                self._tree_paint_parent_item_column(top, 1)
                step_txt = str(sc.get("title") or "シナリオ")
                if isinstance(src_m, dict):
                    row_tip = scenario_source_tooltip_plain(
                        src_m, dn, detail_cell_cfg=dcell
                    )
                else:
                    row_tip = _format_condition_step_tooltip(step_txt, slot)
                for col in range(3):
                    top.setToolTip(col, _normalize_tooltip_text(row_tip))
                self._tree_add_slot_children(top, slot)
                self.master_cond_tree.addTopLevelItem(top)
            self.master_cond_tree.collapseAll()
        self._fit_cond_tree_columns(self.cond_tree)
        self._fit_cond_tree_columns(self.master_cond_tree)

    def _mpv_progress_cache_key(self) -> tuple[Any, ...]:
        sp = self._debug_scan_paths
        head = tuple(str(sp[i]) for i in range(min(5, len(sp))))
        return (
            self._mi_idx,
            self._master_step_idx,
            id(self._scenario_for_dry_run),
            len(sp),
            head,
            tuple(int(x) for x in (self._active_slot_indices or [])),
        )

    def _mpv_effective_master_step_for_preview(self) -> int:
        """進捗用 batch の master_step_idx。全項目完了待機で step が 0 に戻ると n_pick=0 キャッシュが選ばれ
        結合列が空に見えるため、_master_step_pass_complete かつ step==0 のときは len(active) を返す。"""
        act = self._active_slot_indices or []
        step = int(self._master_step_idx)
        if (
            act
            and getattr(self, "_master_step_pass_complete", False)
            and step == 0
        ):
            return len(act)
        return step

    def _mpv_progress_n_pick(self) -> int:
        """scenario_for_stepped_preview の取り込み本数 min(有効 step_idx, len(active))。"""
        act = self._active_slot_indices or []
        if not act:
            return 0
        return min(int(self._mpv_effective_master_step_for_preview()), len(act))

    def _mpv_note_progress_row_peak(self, mi_idx: int, row_count: int) -> None:
        if row_count <= 0:
            return
        cur = int(self._mpv_progress_row_peak_by_mi.get(int(mi_idx), 0))
        if int(row_count) > cur:
            self._mpv_progress_row_peak_by_mi[int(mi_idx)] = int(row_count)

    def _mpv_step_cached_rows_acceptable(
        self,
        rows: list[list[Any]],
        *,
        mi_idx: int,
        n_pick: int,
    ) -> bool:
        """品名など多段階の途中キャッシュ（行数が急減）を表示に使わない。"""
        if not rows:
            return False
        peak = int(self._mpv_progress_row_peak_by_mi.get(int(mi_idx), 0))
        n = len(rows)
        if peak > 0 and n < peak and int(n_pick) < len(self._active_slot_indices or []):
            if n < max(peak // 2, peak - 32):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress skip=thin_step_cache "
                        "mi_idx=%s n_pick=%s rows=%s peak=%s",
                        mi_idx,
                        n_pick,
                        n,
                        peak,
                    )
                except Exception:
                    pass
                return False
        if self._mpv_current_item_has_join_defs(int(mi_idx)):
            prior_peak = self._mpv_prior_peak_rows_before_mi(int(mi_idx))
            if not self._mpv_join_result_usable(
                list(rows),
                mi_idx=int(mi_idx),
                prior_peak_rows=prior_peak,
            ):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress skip=empty_join_col_cache "
                        "mi_idx=%s n_pick=%s rows=%s",
                        mi_idx,
                        n_pick,
                        n,
                    )
                except Exception:
                    pass
                return False
        return True

    def _mpv_anchor_file_path_for_seed(self, mi_idx: int) -> str:
        """seed pool 用: 横断 join の side（錨）ファイルパスを scan_paths から解決。"""
        from svc.svc_data_agg import (  # noqa: WPS433
            _file_path_matches_filter_specs,
            _join_comparison_side_file_filter_specs,
        )

        scen = self._scenario_for_dry_run or {}
        items = list(scen.get("items") or [])
        if mi_idx < 0 or mi_idx >= len(items):
            return ""
        host = items[int(mi_idx)]
        if not isinstance(host, dict):
            return ""
        headers = self._mpv_preview_headers()
        side_specs = _join_comparison_side_file_filter_specs(host, items, headers)
        if not side_specs:
            return ""
        for p in self._debug_scan_paths or []:
            if _file_path_matches_filter_specs(str(p), side_specs):
                return str(p)
        return ""

    def _mpv_is_single_slot_active(self) -> bool:
        return len(self._active_slot_indices or []) == 1

    def _mpv_rows_from_step_cache_n_pick(self, n_pick: int) -> list[list[Any]] | None:
        sk = self._mpv_progress_step_cache_key(int(n_pick))
        cached = self._mpv_progress_rows_step_cache.get(sk)
        if cached is None or not cached:
            return None
        if not self._mpv_step_cached_rows_acceptable(
            cached,
            mi_idx=int(self._mi_idx),
            n_pick=int(n_pick),
        ):
            return None
        return [list(r) for r in cached]

    def _mpv_warmup_single_slot_progress_cache(self, mi_idx: int | None = None) -> None:
        """単一スロット項目: n_pick=1 を先読み（dedup 時は prefetch 無効でもキュー投入）。"""
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_should_warmup_single_slot,
        )

        act = self._active_slot_indices or []
        if len(act) != 1:
            return
        mi = int(self._mi_idx if mi_idx is None else mi_idx)
        if not master_preview_should_warmup_single_slot(
            has_join_defs=self._mpv_current_item_has_join_defs(mi)
        ):
            return
        sk1 = self._mpv_progress_step_cache_key_for(
            1,
            mi_idx=mi,
            active_slot_indices=list(act),
            scenario_id=id(self._scenario_for_dry_run),
            scan_paths=list(self._debug_scan_paths or []),
        )
        if sk1 in self._mpv_progress_rows_step_cache:
            return
        self._mpv_maybe_enqueue_progress_prefetch(
            next_master_step_override=1,
            force=True,
        )
        self._mpv_single_slot_prefetch_pending_sk = sk1
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_progress warmup=single_slot mi_idx=%s "
                "prefetch_pending=1",
                mi,
            )
        except Exception:
            pass

    def _mpv_is_single_slot_prefetch_pending(self) -> bool:
        sk = getattr(self, "_mpv_single_slot_prefetch_pending_sk", None)
        if sk is None:
            return False
        if sk in self._mpv_progress_rows_step_cache:
            self._mpv_single_slot_prefetch_pending_sk = None
            return False
        return True

    def _mpv_clear_single_slot_prefetch_pending(self, sk: tuple[Any, ...]) -> None:
        if getattr(self, "_mpv_single_slot_prefetch_pending_sk", None) == sk:
            self._mpv_single_slot_prefetch_pending_sk = None

    def _mpv_should_defer_join_value_grid_rebuild(self) -> bool:
        """結合項目: step キャッシュ完成前は連続実行中のグリッド再構築を遅延する。"""
        if self._mode != 1 or not self._mpv_current_item_has_join_defs():
            return False
        busy = bool(
            getattr(self, "_continuous_busy", False)
            or getattr(self, "_master_step_loop_busy", False)
            or int(getattr(self, "_mpv_join_compute_busy", 0) or 0) > 0
        )
        if not busy:
            return False
        act = self._active_slot_indices or []
        if not act:
            return False
        n_act = len(act)
        step_idx = int(self._master_step_idx)
        n_pick = 1 if self._mpv_is_single_slot_active() else max(
            1, min(step_idx + 1, n_act)
        )
        if self._mpv_rows_from_step_cache_n_pick(int(n_pick)):
            return False
        return True

    def _mpv_try_join_step0_display_rows(
        self, key: tuple[Any, ...]
    ) -> list[list[Any]] | None:
        """
        結合項目 step0 の表示用行。disk compute は行わず前 table_rows / step キャッシュのみ。
        該当しない場合は None（通常の mpv_progress へ）。
        """
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_step0_should_skip_progress_compute,
        )

        act = self._active_slot_indices or []
        if not act:
            return None
        has_join = self._mpv_current_item_has_join_defs(int(self._mi_idx))
        step_idx = int(self._master_step_idx)
        n_pick = max(1, int(self._mpv_progress_n_pick()))
        cached = self._mpv_rows_from_step_cache_n_pick(n_pick)
        if not master_preview_join_step0_should_skip_progress_compute(
            has_join_defs=has_join,
            master_step_idx=step_idx,
            has_step_cache=bool(cached),
        ):
            return None
        if cached:
            rows = [list(r) for r in cached]
            self._mpv_display_mi_idx = int(self._mi_idx)
            self._mpv_progress_rows_cache = (key, rows)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress reuse=join_step0_step_cache "
                    "mi_idx=%s n_pick=%s rows=%s",
                    self._mi_idx,
                    n_pick,
                    len(rows),
                )
            except Exception:
                pass
            return rows
        if self._mpv_last_valid_table_rows:
            rows_lv = [list(r) for r in self._mpv_last_valid_table_rows]
            disp_mi = getattr(self, "_last_master_completed_mi_idx", None)
            if disp_mi is None:
                disp_mi = int(self._mi_idx)
            self._mpv_display_mi_idx = int(disp_mi)
            self._mpv_progress_rows_cache = (key, rows_lv)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress reuse=join_step0_last_valid "
                    "mi_idx=%s display_mi=%s rows=%s",
                    self._mi_idx,
                    disp_mi,
                    len(rows_lv),
                )
            except Exception:
                pass
            return rows_lv
        prior_tbl = self._mpv_best_prior_table_rows_for_seed(
            int(self._mi_idx),
            n_pick=int(n_pick),
        )
        if prior_tbl:
            prows, src_mi = prior_tbl
            if prows:
                rows_prior = [list(r) for r in prows]
                self._mpv_display_mi_idx = int(src_mi)
                self._mpv_progress_rows_cache = (key, rows_prior)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress reuse=join_step0_prior_table "
                        "mi_idx=%s src_mi=%s rows=%s",
                        self._mi_idx,
                        src_mi,
                        len(rows_prior),
                    )
                except Exception:
                    pass
                return rows_prior
        fb_mi = getattr(self, "_last_master_completed_mi_idx", None)
        if fb_mi is not None:
            ent = self._mpv_progress_rows_by_mi.get(int(fb_mi))
            if ent is not None and ent[1]:
                rows_fb = [list(r) for r in ent[1]]
                self._mpv_display_mi_idx = int(fb_mi)
                self._mpv_progress_rows_cache = (key, rows_fb)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress reuse=join_step0_last_completed_mi "
                        "mi_idx=%s fb_mi=%s rows=%s",
                        self._mi_idx,
                        fb_mi,
                        len(rows_fb),
                    )
                except Exception:
                    pass
                return rows_fb
        self._mpv_display_mi_idx = int(self._mi_idx)
        self._mpv_progress_rows_cache = (key, [])
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_progress skip=join_step0_compute "
                "mi_idx=%s step_idx=%s",
                self._mi_idx,
                step_idx,
            )
        except Exception:
            pass
        return []

    def _mpv_try_single_slot_step0_rows(
        self, key: tuple[Any, ...]
    ) -> list[list[Any]] | None:
        """single_slot step0: n_pick=0 の compute を避け、先読み or 前項目行を返す。"""
        if not self._mpv_is_single_slot_active():
            return None
        if int(self._master_step_idx) != 0:
            return None
        if int(self._mpv_progress_n_pick()) != 0:
            return None
        if self._mpv_current_item_has_join_defs():
            return None

        rows = self._mpv_rows_from_step_cache_n_pick(1)
        if rows:
            self._mpv_display_mi_idx = int(self._mi_idx)
            self._mpv_progress_rows_cache = (key, rows)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress reuse=single_slot_warmup "
                    "mi_idx=%s step_idx=0 rows=%s",
                    self._mi_idx,
                    len(rows),
                )
            except Exception:
                pass
            return rows

        fb_mi = getattr(self, "_last_master_completed_mi_idx", None)
        if fb_mi is not None:
            ent = self._mpv_progress_rows_by_mi.get(int(fb_mi))
            if ent is not None and ent[1]:
                rows = [list(r) for r in ent[1]]
                self._mpv_display_mi_idx = int(fb_mi)
                self._mpv_progress_rows_cache = (key, rows)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress reuse=single_slot_prev_mi "
                        "mi_idx=%s step_idx=0 fb_mi=%s rows=%s",
                        self._mi_idx,
                        fb_mi,
                        len(rows),
                    )
                except Exception:
                    pass
                return rows

        cur_mi = int(self._mi_idx)
        by_mi = getattr(self, "_mpv_progress_rows_by_mi", None) or {}
        for m in sorted(by_mi.keys(), reverse=True):
            if m <= cur_mi:
                ent = by_mi.get(m)
                if ent and ent[1]:
                    rows = [list(r) for r in ent[1]]
                    self._mpv_display_mi_idx = int(m)
                    self._mpv_progress_rows_cache = (key, rows)
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_progress reuse=single_slot_best_cached "
                            "mi_idx=%s step_idx=0 pick_mi=%s rows=%s",
                            self._mi_idx,
                            m,
                            len(rows),
                        )
                    except Exception:
                        pass
                    return rows

        self._mpv_display_mi_idx = int(self._mi_idx)
        self._mpv_progress_rows_cache = (key, [])
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_progress skip=single_slot_step0_compute "
                "mi_idx=%s step_idx=0",
                self._mi_idx,
            )
        except Exception:
            pass
        return []

    def _mpv_wait_single_slot_n_pick1_cache(
        self, *, max_wait_ms: int = 120_000
    ) -> list[list[Any]] | None:
        """先読み完了を短時間ポール（UI イベントは処理する）。"""
        if not self._mpv_is_single_slot_active():
            return None
        if int(max_wait_ms) <= 0:
            return None
        t0 = time.perf_counter()
        deadline = t0 + max(0, int(max_wait_ms)) / 1000.0
        while time.perf_counter() < deadline:
            rows = self._mpv_rows_from_step_cache_n_pick(1)
            if rows:
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress prefetch_wait_hit "
                        "mi_idx=%s rows=%s elapsed_ms=%s",
                        self._mi_idx,
                        len(rows),
                        int((time.perf_counter() - t0) * 1000),
                    )
                except Exception:
                    pass
                self._mpv_clear_single_slot_prefetch_pending(
                    self._mpv_progress_step_cache_key(1)
                )
                return rows
            self._process_events_light()
            from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

            try:
                self._master_poll_cancel(force=True)
            except DataAggCancelled:
                self._master_note_cancel_requested()
                raise
            time.sleep(0.05)
        if self._mpv_is_single_slot_prefetch_pending():
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress prefetch_wait_miss "
                    "mi_idx=%s wait_ms=%s",
                    self._mi_idx,
                    int(max_wait_ms),
                )
            except Exception:
                pass
        return None

    def _mpv_ensure_single_slot_n_pick1_cached(
        self,
        *,
        progress_hook: Any = None,
        frozen_capture_out: dict[str, Any] | None = None,
        wait_async_ms: int = 0,
    ) -> list[list[Any]] | None:
        """single_slot: n_pick=1 の progress 行を最大1回だけ確保する。"""
        if not self._mpv_is_single_slot_active():
            return None
        return self._mpv_ensure_step_n_pick_cached(
            n_pick=1,
            progress_hook=progress_hook,
            frozen_capture_out=frozen_capture_out,
            wait_async_ms=wait_async_ms,
            probe_caller="mpv_single_slot_n_pick1",
        )

    def _mpv_try_colvals_from_step_cache(
        self, *, mi_idx: int, n_pick: int
    ) -> list[str] | None:
        """extract を避け、段階キャッシュから当該列を切り出す。"""
        sk = self._mpv_progress_step_cache_key_for(
            int(n_pick),
            mi_idx=int(mi_idx),
            active_slot_indices=list(self._active_slot_indices or []),
            scenario_id=id(self._scenario_for_dry_run),
            scan_paths=list(self._debug_scan_paths or []),
        )
        cached = self._mpv_progress_rows_step_cache.get(sk)
        if cached is None:
            return None
        if not self._mpv_step_cached_rows_acceptable(
            cached, mi_idx=int(mi_idx), n_pick=int(n_pick)
        ):
            return None
        col: list[str] = []
        for rr in cached[: self._max_value_rows()]:
            v = rr[int(mi_idx)] if int(mi_idx) < len(rr) else None
            col.append("" if v is None else str(v))
        while col and (not str(col[-1]).strip()):
            col.pop()
        if not col:
            return None
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_colvals_from_step_cache mi_idx=%s n_pick=%s col_count=%s",
                mi_idx,
                n_pick,
                len(col),
            )
        except Exception:
            pass
        return col

    def _mpv_master_dbg_progress_hook_or_none(self) -> Any:
        if bool(
            getattr(self, "_master_run_progress_active", False)
            or getattr(self, "_debug_progress_locked", False)
        ):
            return self._master_dbg_batch_progress_hook
        try:
            _pd = getattr(self, "_run_progress_dlg", None)
            if _pd is not None and _pd.isVisible():
                return self._master_dbg_batch_progress_hook
        except Exception:
            pass
        return None

    def _mpv_show_join_compute_progress(self) -> None:
        """結合項目の同期 compute 直前に進捗フェーズを進める（長時間 phase0 固定を防ぐ）。"""
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_sync_compute_progress,
        )

        phase, done = master_preview_join_sync_compute_progress()
        sub_total = len(_MASTER_DEBUG_PROGRESS_PHASES)
        wt = (
            getattr(self, "_master_progress_window_title", None)
            or self._scenario_progress_window_title()
        )
        hook_paths = self._mpv_effective_progress_hook_paths()
        file_total = len(hook_paths) if hook_paths else 0
        self._show_run_progress(
            phase,
            done,
            sub_total,
            window_title=wt,
            detail="0/%s — 開始" % file_total if file_total else "開始",
        )
        self._process_events_light()

    def _mpv_sync_progress_cache_from_step_n_pick(self, n_pick: int) -> bool:
        """step キャッシュを _mpv_progress_rows_cache に同期（結合項目完了後の空表示防止）。"""
        rows = self._mpv_rows_from_step_cache_n_pick(int(n_pick))
        if not rows:
            return False
        key = self._mpv_progress_cache_key()
        copied = [list(r) for r in rows]
        self._mpv_publish_table_rows(copied, mi_idx=int(self._mi_idx))
        n_act = len(self._active_slot_indices or [])
        step_for_mi = min(int(n_pick), n_act) if n_act > 0 else int(n_pick)
        self._mpv_progress_rows_by_mi[int(self._mi_idx)] = (step_for_mi, copied)
        if copied and n_act > 0 and step_for_mi >= n_act:
            self._last_master_completed_mi_idx = int(self._mi_idx)
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_progress sync=step_cache mi_idx=%s n_pick=%s rows=%s",
                self._mi_idx,
                n_pick,
                len(copied),
            )
        except Exception:
            pass
        return True

    def _mpv_try_join_step_cache_fallback_rows(
        self, *, n_pick: int
    ) -> list[list[Any]] | None:
        """結合項目で compute が空のとき step キャッシュから復元。"""
        if not self._mpv_current_item_has_join_defs():
            return None
        for pick in (int(n_pick), len(self._active_slot_indices or [])):
            if pick <= 0:
                continue
            rows = self._mpv_rows_from_step_cache_n_pick(pick)
            if rows:
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress reuse=join_step_cache_fallback "
                        "mi_idx=%s step_idx=%s n_pick=%s rows=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        pick,
                        len(rows),
                    )
                except Exception:
                    pass
                return rows
        prior = self._mpv_best_prior_table_rows_for_seed(
            int(self._mi_idx),
            n_pick=int(n_pick),
        )
        if prior:
            prows, src_mi = prior
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress reuse=join_prior_table_fallback "
                    "mi_idx=%s src_mi=%s rows=%s",
                    self._mi_idx,
                    src_mi,
                    len(prows),
                )
            except Exception:
                pass
            return prows
        return None

    def _master_mpv_compute_lock_acquire_ui(self) -> None:
        """GUI スレッド: ロック待ち中も Qt イベントとキャンセルを処理する。"""
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        lock = self._mpv_prog_compute_lock
        if not self._master_on_ui_thread():
            while True:
                if lock.acquire(blocking=False):
                    return
                if self._master_cancel_pending():
                    raise DataAggCancelled()
                chk = self._master_run_cancel_check()
                if chk is not None:
                    try:
                        chk(force=True)
                    except DataAggCancelled:
                        self._master_note_cancel_requested()
                        raise
                time.sleep(0.01)
            return
        while True:
            if lock.acquire(blocking=False):
                return
            if self._master_cancel_pending():
                raise DataAggCancelled()
            self._process_events_for_master_cancel()
            try:
                self._master_poll_cancel(force=True)
            except DataAggCancelled:
                self._master_note_cancel_requested()
                raise
            time.sleep(0.01)

    def _master_mpv_compute_lock_release_ui(self) -> None:
        self._mpv_prog_compute_lock_release()

    def _mpv_prog_compute_lock_release(self) -> None:
        try:
            self._mpv_clear_wb_worker_if_done()
            self._mpv_flush_pending_wb_frames_unlocked()
        except Exception:
            pass
        self._mpv_prog_compute_lock.release()

    def _mpv_ensure_step_n_pick_cached(
        self,
        *,
        n_pick: int,
        progress_hook: Any = None,
        frozen_capture_out: dict[str, Any] | None = None,
        wait_async_ms: int = 0,
        probe_caller: str = "mpv_step_n_pick",
    ) -> list[list[Any]] | None:
        """指定 n_pick の progress 行を最大1回だけ確保（single/multi スロット共通）。"""
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        self._master_raise_if_cancelled()
        n_pick_i = int(n_pick)
        n_act = len(self._active_slot_indices or [])
        if n_act <= 0 or n_pick_i <= 0:
            return None
        n_pick_i = min(n_pick_i, n_act)
        sk = self._mpv_progress_step_cache_key(n_pick_i)
        cached = self._mpv_progress_rows_step_cache.get(sk)
        need_frozen_capture = (
            frozen_capture_out is not None
            and self._mpv_frozen_columns_enabled()
            and int(self._mi_idx) not in self._mpv_frozen_snapshots
        )
        if (
            cached is not None
            and self._mpv_step_cached_rows_acceptable(
                cached,
                mi_idx=int(self._mi_idx),
                n_pick=n_pick_i,
            )
            and not need_frozen_capture
        ):
            if self._mpv_is_single_slot_active() and n_pick_i == 1:
                self._mpv_clear_single_slot_prefetch_pending(sk)
            return [list(r) for r in cached]
        if wait_async_ms > 0 and self._mpv_is_single_slot_active() and n_pick_i == 1:
            waited = self._mpv_wait_single_slot_n_pick1_cache(max_wait_ms=int(wait_async_ms))
            if waited:
                return waited
        if (
            self._mpv_is_single_slot_active()
            and n_pick_i == 1
            and probe_caller != "mpv_progress_prefetch"
        ):
            from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                master_preview_single_slot_sync_wait_ms,
            )

            if self._mpv_is_single_slot_prefetch_pending():
                sync_wait = master_preview_single_slot_sync_wait_ms(
                    prefetch_pending=True
                )
                if sync_wait > 0:
                    waited_sync = self._mpv_wait_single_slot_n_pick1_cache(
                        max_wait_ms=int(sync_wait)
                    )
                    if waited_sync:
                        return waited_sync
        if (
            self._mpv_is_single_slot_active()
            and n_pick_i == 1
            and probe_caller != "mpv_progress_prefetch"
            and self._mpv_is_single_slot_prefetch_pending()
        ):
            waited_long = self._mpv_wait_single_slot_n_pick1_cache(max_wait_ms=120_000)
            if waited_long:
                return waited_long
        while (
            self._mpv_is_single_slot_active()
            and n_pick_i == 1
            and probe_caller != "mpv_progress_prefetch"
            and self._mpv_is_single_slot_prefetch_pending()
        ):
            hit_pf = self._mpv_rows_from_step_cache_n_pick(1)
            if hit_pf:
                return [list(r) for r in hit_pf]
            self._master_raise_if_cancelled(force_poll=True)
            self._process_events_for_master_cancel()
            time.sleep(0.05)
        self._master_mpv_compute_lock_acquire_ui()
        try:
            cached = self._mpv_progress_rows_step_cache.get(sk)
            if (
                cached is not None
                and self._mpv_step_cached_rows_acceptable(
                    cached,
                    mi_idx=int(self._mi_idx),
                    n_pick=n_pick_i,
                )
                and not need_frozen_capture
            ):
                return [list(r) for r in cached]
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress step_sync_compute "
                    "mi_idx=%s n_pick=%s caller=%s",
                    self._mi_idx,
                    n_pick_i,
                    probe_caller,
                )
            except Exception:
                pass
            join_busy = self._mpv_current_item_has_join_defs()
            if join_busy:
                self._mpv_begin_join_compute()
            try:
                rows = self._mpv_compute_progress_table_rows(
                    mi_idx=int(self._mi_idx),
                    master_step_idx=n_pick_i,
                    active_slot_indices=list(self._active_slot_indices or []),
                    scenario_base=self._scenario_for_dry_run or {},
                    scan_paths=list(self._debug_scan_paths or []),
                    n_pick=n_pick_i,
                    use_max_sources=False,
                    progress_hook=progress_hook,
                    probe_caller=probe_caller,
                    frozen_capture_out=frozen_capture_out,
                )
            except DataAggCancelled:
                self._master_note_cancel_requested()
                raise
            finally:
                if join_busy:
                    self._mpv_end_join_compute()
            self._master_raise_if_cancelled()
            self._mpv_store_step_cache(
                sk,
                rows,
                mi_idx=int(self._mi_idx),
                master_step_idx=n_pick_i,
            )
            if (
                self._mpv_is_single_slot_active()
                and n_pick_i == 1
                and probe_caller == "mpv_single_slot_n_pick1"
            ):
                self._mpv_clear_single_slot_prefetch_pending(sk)
            return rows
        finally:
            self._master_mpv_compute_lock_release_ui()

    def _mpv_resolve_master_step_colvals(self, si: int) -> list[str]:
        """
        マスタステップの取得値列。結合項目はキャッシュ or 同期 compute（進捗付き）。
        非 join の step0 は compute_then_colvals（extract_first による二重走査を避ける）。
        それ以外の非 join は先読み／progress キャッシュ、だめなら extract。
        """
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_colvals_should_call_progress_batch,
            master_preview_join_requires_sync_compute_before_colvals,
            master_preview_should_warmup_single_slot,
            master_preview_step0_should_block_wait_n_pick1,
            master_preview_step0_wait_async_ms,
        )

        self._master_raise_if_cancelled()

        if not self._scenario_for_dry_run or not self._debug_scan_paths:
            m = self._current_master()
            slot = m["scenarios"][si]["slot"]
            assert slot is not None
            return self._icap(list(slot.get("values_prod", slot["values_column"])))

        has_join = self._mpv_current_item_has_join_defs()
        single = self._mpv_is_single_slot_active()
        step_idx = int(self._master_step_idx)
        n_act = len(self._active_slot_indices or [])
        n_pick_after = min(step_idx + 1, n_act) if n_act > 0 else 0

        if single and step_idx == 0 and master_preview_should_warmup_single_slot(
            has_join_defs=has_join
        ):
            self._mpv_warmup_single_slot_progress_cache()

        col_from_prog: list[str] = []
        cache_hit = False

        if single and n_pick_after > 0:
            hit = self._mpv_try_colvals_from_step_cache(
                mi_idx=int(self._mi_idx), n_pick=n_pick_after
            )
            if hit is not None:
                col_from_prog = hit
                cache_hit = True
            elif (
                step_idx == 0
                and not has_join
                and self._mpv_is_single_slot_prefetch_pending()
            ):
                from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                    master_preview_single_slot_sync_wait_ms,
                )

                self._mpv_wait_single_slot_n_pick1_cache(
                    max_wait_ms=master_preview_single_slot_sync_wait_ms(
                        prefetch_pending=True
                    )
                )
                hit = self._mpv_try_colvals_from_step_cache(
                    mi_idx=int(self._mi_idx), n_pick=n_pick_after
                )
                if hit is not None:
                    col_from_prog = hit
                    cache_hit = True
            elif (
                step_idx == 0
                and master_preview_step0_should_block_wait_n_pick1(has_join_defs=has_join)
            ):
                wait_ms = master_preview_step0_wait_async_ms(has_join_defs=has_join)
                hook0 = self._mpv_master_dbg_progress_hook_or_none()
                self._mpv_ensure_step_n_pick_cached(
                    n_pick=1,
                    progress_hook=hook0,
                    wait_async_ms=wait_ms,
                    probe_caller="mpv_single_slot_step0_wait",
                )
                hit = self._mpv_try_colvals_from_step_cache(
                    mi_idx=int(self._mi_idx), n_pick=1
                )
                if hit is not None:
                    col_from_prog = hit
                    cache_hit = True

        if (
            not col_from_prog
            and step_idx > 0
            and master_preview_colvals_should_call_progress_batch(
                master_step_idx=step_idx,
                can_use_progress_cache=self._mpv_can_colvals_from_progress(),
            )
        ):
            try:
                t_prog_col = time.perf_counter()
                prog_rows_now = self._mpv_progress_batch_rows()
                for rr in prog_rows_now[: self._max_value_rows()]:
                    v = rr[self._mi_idx] if self._mi_idx < len(rr) else None
                    col_from_prog.append("" if v is None else str(v))
                while col_from_prog and (not str(col_from_prog[-1]).strip()):
                    col_from_prog.pop()
                cache_hit = bool(col_from_prog)
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_colvals_from_progress mi_idx=%s step_idx=%s si=%s "
                    "row_count=%s col_count=%s elapsed_ms=%s",
                    self._mi_idx,
                    step_idx,
                    si,
                    len(prog_rows_now),
                    len(col_from_prog),
                    int((time.perf_counter() - t_prog_col) * 1000),
                )
            except Exception:
                col_from_prog = []

        if master_preview_join_requires_sync_compute_before_colvals(
            has_join_defs=has_join,
            cache_hit=cache_hit,
        ):
            # 結合項目の凍結キャプチャは pool 構造が表行と合わず列が空になるため行わない。
            fcap_out: dict[str, Any] | None = None
            hook = self._mpv_master_dbg_progress_hook_or_none()
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_colvals strategy=join_sync_compute "
                    "mi_idx=%s step_idx=%s si=%s n_pick=%s",
                    self._mi_idx,
                    step_idx,
                    si,
                    n_pick_after,
                )
            except Exception:
                pass
            self._mpv_show_join_compute_progress()
            self._master_raise_if_cancelled()
            self._mpv_ensure_step_n_pick_cached(
                n_pick=n_pick_after,
                progress_hook=hook,
                frozen_capture_out=fcap_out,
                wait_async_ms=0,
                probe_caller="mpv_join_step_colvals",
            )
            self._master_raise_if_cancelled()
            self._mpv_sync_progress_cache_from_step_n_pick(n_pick_after)
            hit = self._mpv_try_colvals_from_step_cache(
                mi_idx=int(self._mi_idx), n_pick=n_pick_after
            )
            if hit is not None:
                col_from_prog = hit
                cache_hit = True

        if col_from_prog:
            return self._icap(col_from_prog)

        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_colvals_from_progress skip=no_progress_cache "
                "mi_idx=%s step_idx=%s si=%s",
                self._mi_idx,
                step_idx,
                si,
            )
        except Exception:
            pass

        # #4: step0 非 join は extract_first せず、先に progress compute → 列はキャッシュから
        if step_idx == 0 and not has_join and n_pick_after > 0:
            hook0 = self._mpv_master_dbg_progress_hook_or_none()
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_colvals strategy=compute_then_colvals "
                    "mi_idx=%s step_idx=%s si=%s n_pick=%s",
                    self._mi_idx,
                    step_idx,
                    si,
                    n_pick_after,
                )
            except Exception:
                pass
            self._master_raise_if_cancelled()
            self._mpv_ensure_step_n_pick_cached(
                n_pick=n_pick_after,
                progress_hook=hook0,
                wait_async_ms=0,
                probe_caller="mpv_step0_compute_then_colvals",
            )
            self._master_raise_if_cancelled()
            self._mpv_sync_progress_cache_from_step_n_pick(n_pick_after)
            hit = self._mpv_try_colvals_from_step_cache(
                mi_idx=int(self._mi_idx), n_pick=n_pick_after
            )
            if hit is not None:
                return self._icap(hit)

        t_extract = time.perf_counter()
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_extract_start mi_idx=%s step_idx=%s si=%s title=%s",
                self._mi_idx,
                step_idx,
                si,
                str(
                    (self._current_master().get("scenarios") or [])[si].get("title")
                    or ""
                ),
            )
        except Exception:
            pass
        out = self._icap(self._mpv_extract_colvals(self._mi_idx, si))
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_extract_end mi_idx=%s step_idx=%s si=%s col_count=%s elapsed_ms=%s",
                self._mi_idx,
                step_idx,
                si,
                len(out),
                int((time.perf_counter() - t_extract) * 1000),
            )
        except Exception:
            pass
        return out

    def _mpv_progress_step_cache_key(self, n_pick: int) -> tuple[Any, ...]:
        sp = self._debug_scan_paths
        head = tuple(str(sp[i]) for i in range(min(5, len(sp))))
        return (
            int(self._mi_idx),
            int(n_pick),
            id(self._scenario_for_dry_run),
            len(sp),
            head,
            tuple(int(x) for x in self._active_slot_indices),
        )

    def _mpv_frozen_columns_enabled(self) -> bool:
        from core import core_env  # noqa: WPS433

        return core_env.data_agg_master_frozen_columns_enabled()

    def _mpv_preview_headers(self) -> list[str]:
        from svc.svc_data_agg_scenario import output_table_headers_for_scenario  # noqa: WPS433

        return output_table_headers_for_scenario(self._scenario_for_dry_run or {})

    def _debug_carry_empty_target_names(self) -> set[str]:
        """
        前置保持(carry_empty)対象項目名。
        シナリオ編集からの起動は scenario_for_dry_run が無いため live_items を使う。
        マスタは dry_run の items を優先する。
        """
        dry = list((self._scenario_for_dry_run or {}).get("items") or [])
        live = list(self._live_items or [])
        if self._mode == 0:
            items = live or dry
        else:
            items = dry or live
        return carry_empty_target_names_from_items(items)

    def _decorate_debug_grid_headers(self, headers: list[str]) -> list[str]:
        """デバッグ結果一覧用: 前置保持項目の見出しに「・」を付ける。"""
        return decorate_debug_carry_empty_headers(
            headers, self._debug_carry_empty_target_names()
        )

    def _mpv_preview_compute_paths(self) -> list[str]:
        """compute_batch 内部 filter と同じ絞り込み後パス（凍結検証用）。"""
        return preview_compute_file_paths(
            self._scenario_for_dry_run or {},
            list(self._debug_scan_paths or []),
        )

    def _mpv_current_item_has_join_defs(self, mi_idx: int | None = None) -> bool:
        """現在（または指定）マスタ項目に結合定義があるか。"""
        idx = int(self._mi_idx if mi_idx is None else mi_idx)
        items = list((self._scenario_for_dry_run or {}).get("items") or [])
        if idx < 0 or idx >= len(items):
            return False
        it = items[idx]
        if not isinstance(it, dict):
            return False
        from svc.svc_data_agg import _item_join_defs_list  # noqa: WPS433

        return bool(_item_join_defs_list(it))

    def _mpv_master_item_label(self, mi_idx: int) -> str:
        items = list((self._scenario_for_dry_run or {}).get("items") or [])
        if mi_idx < 0 or mi_idx >= len(items):
            return ""
        it = items[mi_idx]
        if not isinstance(it, dict):
            return ""
        return str(it.get("name") or it.get("id") or "").strip()

    def _mpv_frozen_context_for_mi(self, mi_idx: int) -> tuple[dict[str, Any] | None, int | None]:
        """次項目 compute 用: (frozen_prior, frozen_through_mi)。不適格時は (None, None)。"""
        if self._mpv_current_item_has_join_defs(int(mi_idx)):
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_frozen skip mi_idx=%s reason=join_item",
                    mi_idx,
                )
            except Exception:
                pass
            return None, None
        if not self._mpv_frozen_columns_enabled():
            return None, None
        if int(mi_idx) <= 0:
            return None, None
        headers = self._mpv_preview_headers()
        paths = self._mpv_preview_compute_paths()
        from svc.data_agg_master_preview import best_frozen_snapshot_for_mi  # noqa: WPS433

        snap, through = best_frozen_snapshot_for_mi(
            self._mpv_frozen_snapshots,
            int(mi_idx),
            headers=headers,
            file_paths=paths,
        )
        if snap is None or through is None:
            try:
                scan_n = len(self._debug_scan_paths or [])
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_frozen skip mi_idx=%s reason=no_snapshot "
                    "expected_through=%s paths_filter=%s scan_paths=%s snapshots=%s",
                    mi_idx,
                    int(mi_idx) - 1,
                    len(paths),
                    scan_n,
                    len(self._mpv_frozen_snapshots),
                )
            except Exception:
                pass
            return None, None
        if int(through) < int(mi_idx) - 1:
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_frozen apply carried_through=%s mi_idx=%s "
                    "paths_filter=%s snap_paths=%s",
                    through,
                    mi_idx,
                    len(paths),
                    snap.get("paths_count"),
                )
            except Exception:
                pass
        return snap, int(through)

    def _mpv_store_frozen_snapshot(self, cap: dict[str, Any]) -> None:
        if cap.get("version") != FROZEN_SNAPSHOT_VERSION:
            return
        through = cap.get("through_mi")
        if not isinstance(through, int) or through < 0:
            return
        stored = dict(cap)
        self._mpv_frozen_snapshots[int(through)] = stored
        try:
            rbk = stored.get("rows_by_key") or {}
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_frozen stored through_mi=%s keys=%s paths=%s snapshots=%s",
                through,
                len(rbk) if isinstance(rbk, dict) else 0,
                stored.get("paths_count"),
                len(self._mpv_frozen_snapshots),
            )
        except Exception:
            pass

    def _mpv_one_shot_eligible(self) -> bool:
        from core import core_env

        if not core_env.data_agg_master_progress_one_shot_enabled():
            return False
        if self._mode != 1:
            return False
        return master_preview_one_shot_eligible(
            self._scenario_for_dry_run or {},
            int(self._mi_idx),
            list(self._active_slot_indices or []),
        )

    def _mpv_progress_step_cache_key_for(
        self,
        n_pick: int,
        *,
        mi_idx: int,
        active_slot_indices: list[int],
        scenario_id: int,
        scan_paths: list[str],
    ) -> tuple[Any, ...]:
        head = tuple(str(scan_paths[i]) for i in range(min(5, len(scan_paths))))
        return (
            int(mi_idx),
            int(n_pick),
            scenario_id,
            len(scan_paths),
            head,
            tuple(int(x) for x in active_slot_indices),
        )

    def _mpv_prior_peak_rows_before_mi(self, before_mi: int) -> int:
        peaks = getattr(self, "_mpv_progress_row_peak_by_mi", {}) or {}
        best = 0
        for mi, peak in peaks.items():
            if int(mi) < int(before_mi):
                best = max(best, int(peak))
        if best > 0:
            return best
        by_mi = getattr(self, "_mpv_progress_rows_by_mi", {}) or {}
        for mi in range(int(before_mi)):
            ent = by_mi.get(mi)
            if ent and ent[1]:
                best = max(best, len(ent[1]))
        return best

    def _mpv_completed_table_rows_for_prior_mi(
        self,
        prior_mi: int,
        *,
        by_mi: dict[int, tuple[int, list[list[Any]]]] | None = None,
    ) -> list[list[Any]] | None:
        """前項目 mi の全スロット完了相当 table_rows。呼び出し側 n_pick は使わない。"""
        rows_by_mi: dict[int, tuple[int, list[list[Any]]]]
        if by_mi is not None:
            rows_by_mi = by_mi
        else:
            raw_by = getattr(self, "_mpv_progress_rows_by_mi", None)
            rows_by_mi = raw_by if isinstance(raw_by, dict) else {}
        mi_saved = int(self._mi_idx)
        step_saved = int(self._master_step_idx)
        try:
            self._mi_idx = int(prior_mi)
            self._rebuild_active_slots()
            n_act = len(self._active_slot_indices or [])
            if n_act > 0:
                prev = rows_by_mi.get(int(prior_mi))
                if prev is not None and prev[1]:
                    po = int(prev[0])
                    if po >= n_act:
                        return [list(r) for r in prev[1]]
                rows = self._mpv_rows_from_step_cache_n_pick(n_act)
                if rows:
                    return [list(r) for r in rows]
            ent = rows_by_mi.get(int(prior_mi))
            if ent and ent[1]:
                return [list(r) for r in ent[1]]
            return None
        finally:
            self._mi_idx = mi_saved
            self._master_step_idx = step_saved
            self._rebuild_active_slots()

    def _mpv_best_prior_table_rows_for_seed(
        self,
        before_mi: int,
        *,
        n_pick: int | None = None,
    ) -> tuple[list[list[Any]], int] | None:
        """before_mi より前の項目で行数最大の table_rows を返す (rows, source_mi)。同数なら mi 大を優先。"""
        del n_pick  # 前項目参照では現項目の n_pick を使わない
        by_mi = getattr(self, "_mpv_progress_rows_by_mi", {}) or {}
        peaks = getattr(self, "_mpv_progress_row_peak_by_mi", {}) or {}
        best_mi: int | None = None
        best_rows = 0
        for mi in range(int(before_mi)):
            peak = int(peaks.get(mi, 0))
            ent = by_mi.get(mi)
            nrow = peak
            if ent and ent[1]:
                nrow = max(nrow, len(ent[1]))
            if nrow > best_rows or (
                nrow == best_rows
                and nrow > 0
                and (best_mi is None or int(mi) > int(best_mi))
            ):
                best_rows = nrow
                best_mi = mi
        if best_mi is None or best_rows <= 0:
            return None
        rows = self._mpv_completed_table_rows_for_prior_mi(int(best_mi), by_mi=by_mi)
        if rows:
            return ([list(r) for r in rows], int(best_mi))
        return None

    def _mpv_prior_rows_for_stacked_join_seed(
        self,
        before_mi: int,
        join_defs: list[dict],
        *,
        n_pick: int | None = None,
    ) -> tuple[list[list[Any]], int] | None:
        """積み上げ join: 結合比較列が埋まった前段 table_rows を優先して返す (rows, source_mi)。"""
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_target_headers,
            master_preview_stacked_seed_join_targets_fill_ratio,
            master_preview_stacked_seed_usable,
        )

        del n_pick  # 前項目参照では現項目の n_pick を使わない
        targets = master_preview_join_target_headers(list(join_defs or []))
        headers = self._mpv_preview_headers()
        if not targets or not headers:
            return self._mpv_best_prior_table_rows_for_seed(int(before_mi))
        by_mi = getattr(self, "_mpv_progress_rows_by_mi", {}) or {}
        prior_join_mi = int(before_mi) - 1
        if prior_join_mi >= 0 and self._mpv_current_item_has_join_defs(prior_join_mi):
            prior_rows = self._mpv_completed_table_rows_for_prior_mi(
                prior_join_mi, by_mi=by_mi
            )
            if prior_rows and master_preview_stacked_seed_usable(
                prior_rows, headers, targets
            ):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_stacked_seed_prior_pick "
                        "mi_idx=%s src_mi=%s rows=%s fill_ratio=%.4f targets=%s "
                        "reason=immediate_prior_join",
                        before_mi,
                        prior_join_mi,
                        len(prior_rows),
                        master_preview_stacked_seed_join_targets_fill_ratio(
                            prior_rows, headers, targets
                        ),
                        targets,
                    )
                except Exception:
                    pass
                return ([list(r) for r in prior_rows], int(prior_join_mi))
        peaks = getattr(self, "_mpv_progress_row_peak_by_mi", {}) or {}
        best_mi: int | None = None
        best_ratio = -1.0
        best_rows = 0
        mi_saved = int(self._mi_idx)
        step_saved = int(self._master_step_idx)
        try:
            for mi in range(int(before_mi)):
                peak = int(peaks.get(mi, 0))
                ent = by_mi.get(mi)
                nrow = peak
                if ent and ent[1]:
                    nrow = max(nrow, len(ent[1]))
                if nrow <= 0:
                    continue
                rows = self._mpv_completed_table_rows_for_prior_mi(int(mi), by_mi=by_mi)
                if not rows:
                    continue
                ratio = master_preview_stacked_seed_join_targets_fill_ratio(
                    rows, headers, targets
                )
                if ratio > best_ratio or (
                    abs(ratio - best_ratio) < 1e-9
                    and nrow > best_rows
                ) or (
                    abs(ratio - best_ratio) < 1e-9
                    and nrow == best_rows
                    and (best_mi is None or int(mi) > int(best_mi))
                ):
                    best_ratio = ratio
                    best_rows = nrow
                    best_mi = int(mi)
            if best_mi is None:
                return self._mpv_best_prior_table_rows_for_seed(int(before_mi))
            rows_out = self._mpv_completed_table_rows_for_prior_mi(
                int(best_mi), by_mi=by_mi
            )
            if not rows_out:
                return self._mpv_best_prior_table_rows_for_seed(int(before_mi))
            if not master_preview_stacked_seed_usable(
                rows_out, headers, targets
            ):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_stacked_seed_prior_weak "
                        "mi_idx=%s src_mi=%s fill_ratio=%.4f targets=%s",
                        before_mi,
                        best_mi,
                        best_ratio,
                        targets,
                    )
                except Exception:
                    pass
                return ([list(r) for r in rows_out], int(best_mi))
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_stacked_seed_prior_pick "
                    "mi_idx=%s src_mi=%s rows=%s fill_ratio=%.4f targets=%s",
                    before_mi,
                    best_mi,
                    len(rows_out),
                    best_ratio,
                    targets,
                )
            except Exception:
                pass
            return ([list(r) for r in rows_out], int(best_mi))
        finally:
            self._mi_idx = mi_saved
            self._master_step_idx = step_saved
            self._rebuild_active_slots()

    def _mpv_join_item_complete(
        self,
        *,
        mi_idx: int,
        master_step_idx: int,
    ) -> bool:
        act = self._active_slot_indices or []
        if not act:
            return True
        return int(master_step_idx) >= len(act)

    def _mpv_join_result_usable(
        self,
        rows: list[list[Any]],
        *,
        mi_idx: int,
        prior_peak_rows: int,
    ) -> bool:
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_compute_rows_acceptable,
            master_preview_join_result_usable,
        )

        row_ok = master_preview_join_compute_rows_acceptable(
            new_rows=len(rows or []),
            prior_peak_rows=int(prior_peak_rows),
            item_complete=True,
        )
        return master_preview_join_result_usable(
            rows=list(rows or []),
            col_idx=int(mi_idx),
            row_count_acceptable=row_ok,
        )

    def _mpv_coalesce_join_compute_rows(
        self,
        *,
        mi_idx: int,
        rows: list[list[Any]],
        master_step_idx: int,
        n_pick: int,
    ) -> list[list[Any]]:
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_compute_rows_acceptable,
        )

        if not self._mpv_current_item_has_join_defs(int(mi_idx)):
            return rows
        item_complete = self._mpv_join_item_complete(
            mi_idx=int(mi_idx),
            master_step_idx=int(master_step_idx),
        )
        prior_peak = self._mpv_prior_peak_rows_before_mi(int(mi_idx))
        row_ok = master_preview_join_compute_rows_acceptable(
            new_rows=len(rows or []),
            prior_peak_rows=prior_peak,
            item_complete=item_complete,
        )
        if row_ok:
            if not item_complete:
                return rows
            if self._mpv_join_result_usable(
                list(rows or []),
                mi_idx=int(mi_idx),
                prior_peak_rows=prior_peak,
            ):
                return rows
            fb_rows = self._mpv_try_join_step_cache_fallback_rows(n_pick=int(n_pick))
            if fb_rows and self._mpv_join_result_usable(
                fb_rows,
                mi_idx=int(mi_idx),
                prior_peak_rows=prior_peak,
            ):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_join_coalesce fallback=step_cache "
                        "mi_idx=%s new_rows=%s fb_rows=%s prior_peak=%s reason=empty_host_col",
                        mi_idx,
                        len(rows or []),
                        len(fb_rows),
                        prior_peak,
                    )
                except Exception:
                    pass
                return fb_rows
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_join_coalesce keep=empty_host_col "
                    "mi_idx=%s rows=%s prior_peak=%s",
                    mi_idx,
                    len(rows or []),
                    prior_peak,
                )
            except Exception:
                pass
            return rows
        fb_rows = self._mpv_try_join_step_cache_fallback_rows(n_pick=int(n_pick))
        if fb_rows and self._mpv_join_result_usable(
            fb_rows,
            mi_idx=int(mi_idx),
            prior_peak_rows=prior_peak,
        ):
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_join_coalesce fallback=step_cache "
                    "mi_idx=%s new_rows=%s fb_rows=%s prior_peak=%s",
                    mi_idx,
                    len(rows or []),
                    len(fb_rows),
                    prior_peak,
                )
            except Exception:
                pass
            return fb_rows
        prior = self._mpv_best_prior_table_rows_for_seed(
            int(mi_idx),
            n_pick=int(n_pick),
        )
        if prior and self._mpv_current_item_has_join_defs(int(mi_idx)) and rows:
            return rows
        if prior:
            prows, src_mi = prior
            if self._mpv_join_result_usable(
                prows,
                mi_idx=int(mi_idx),
                prior_peak_rows=prior_peak,
            ):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_join_coalesce fallback=prior_table "
                        "mi_idx=%s new_rows=%s src_mi=%s fb_rows=%s prior_peak=%s",
                        mi_idx,
                        len(rows or []),
                        src_mi,
                        len(prows),
                        prior_peak,
                    )
                except Exception:
                    pass
                return prows
        return rows

    def _mpv_apply_join_item_debug_diag(
        self,
        dd: dict[str, Any],
        *,
        mi_idx: int,
        join_item: bool,
        n_pick: int,
        n_act: int,
        master_step_idx: int,
        probe_caller: str,
    ) -> None:
        """結合項目の __debug_diag: 本番経路・seed 方針（連鎖結合時のみ seed 利用）。"""
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_chain_targets_prior_item,
            master_preview_join_search_skip_seed,
            master_preview_should_use_join_search_seed_pool,
            master_preview_should_use_prior_join_pool_as_seed,
            master_preview_should_use_stacked_join,
        )
        from svc.data_agg_master_preview import (  # noqa: WPS433
            table_rows_to_join_search_seed_pool,
        )
        from svc.svc_data_agg import _item_join_defs_list  # noqa: WPS433

        if (
            n_act > 0
            and int(n_pick) >= n_act
            and int(master_step_idx) >= n_act
            and not join_item
            and probe_caller not in ("mpv_final_display", "mpv_production_parity")
        ):
            dd["preview_use_production_table_rows"] = True
        else:
            dd["preview_use_production_table_rows"] = False

        if not join_item:
            return

        prior_mi = int(mi_idx) - 1
        prior_name = self._mpv_master_item_label(prior_mi) if prior_mi >= 0 else ""
        join_defs: list[dict] = []
        items = list((self._scenario_for_dry_run or {}).get("items") or [])
        if 0 <= int(mi_idx) < len(items):
            it = items[int(mi_idx)]
            if isinstance(it, dict):
                join_defs = list(_item_join_defs_list(it))

        seed_pool = self._mpv_join_search_pool_seed or []
        seed_rows = len(seed_pool)
        file_count = len(self._debug_scan_paths or [])
        prior_had_join = (
            self._mpv_current_item_has_join_defs(prior_mi) if prior_mi >= 0 else False
        )
        chain_targets = master_preview_join_chain_targets_prior_item(
            prior_item_name=prior_name,
            join_defs=join_defs,
        )
        use_chain_seed = master_preview_should_use_join_search_seed_pool(
            chain_targets_prior=chain_targets,
            seed_pool_rows=seed_rows,
        )
        use_prior_seed = master_preview_should_use_prior_join_pool_as_seed(
            prior_mi_had_join=prior_had_join,
            seed_pool_rows=seed_rows,
            file_count=file_count,
        )
        use_prior_table_seed = False
        prior_seed_pool: list[dict[str, Any]] = []
        headers = self._mpv_preview_headers()
        join_targets = [
            str(d.get("item") or "").strip()
            for d in join_defs
            if str(d.get("item") or "").strip()
        ]
        prior_tbl = self._mpv_prior_rows_for_stacked_join_seed(
            int(mi_idx),
            join_defs,
            n_pick=int(n_pick),
        )
        if prior_tbl:
            prows, src_mi = prior_tbl
            if master_preview_should_use_stacked_join(prior_table_rows=len(prows)):
                row_fps = self._mpv_row_file_paths_for_stacked_seed(
                    len(prows),
                    prows=prows,
                    headers=headers,
                    source_mi=int(src_mi),
                )
                prior_seed_pool = table_rows_to_join_search_seed_pool(
                    headers,
                    prows,
                    row_file_paths=row_fps,
                    stacked_join=True,
                )
                from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                    master_preview_stacked_seed_usable,
                )

                seed_ok = master_preview_stacked_seed_usable(
                    prows,
                    headers,
                    join_targets,
                )
                use_prior_table_seed = bool(prior_seed_pool) and seed_ok
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_join_seed_prior_table mi_idx=%s "
                        "src_mi=%s prior_rows=%s pool_rows=%s stacked=1 seed_ok=%s",
                        mi_idx,
                        src_mi,
                        len(prows),
                        len(prior_seed_pool),
                        seed_ok,
                    )
                except Exception:
                    pass
                if prior_seed_pool and seed_ok:
                    dd["preview_anchor_row_keys"] = [
                        [str(r.get("__file_path") or ""), int(r.get("__iter_index", 0))]
                        for r in prior_seed_pool
                    ]
                elif prior_seed_pool and not seed_ok:
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_join_seed_fallback=production_table_rows "
                            "mi_idx=%s src_mi=%s targets=%s",
                            mi_idx,
                            src_mi,
                            join_targets,
                        )
                    except Exception:
                        pass
        skip_seed = master_preview_join_search_skip_seed(
            chain_targets_prior=chain_targets,
            use_prior_pool_seed=use_prior_seed,
            use_chain_pool_seed=use_chain_seed,
            use_prior_table_seed=use_prior_table_seed,
        )
        dd["join_search_skip_seed"] = bool(skip_seed)
        if use_prior_table_seed and prior_seed_pool:
            dd["join_search_seed_pool"] = prior_seed_pool
            dd["join_search_seed_from_table_rows"] = True
            dd["master_preview_stacked_join"] = True
            dd["preview_extract_item_allowlist"] = [int(mi_idx)]
            dd["master_preview_join_read_full_files"] = False
        elif prior_seed_pool and not use_prior_table_seed:
            dd.pop("join_search_seed_pool", None)
            dd.pop("master_preview_stacked_join", None)
            dd.pop("join_search_seed_from_table_rows", None)
            dd["preview_use_production_table_rows"] = True
            dd["master_preview_join_read_full_files"] = True
        elif not skip_seed and seed_pool:
            dd.pop("master_preview_stacked_join", None)
            dd["join_search_seed_pool"] = [
                dict(r) for r in seed_pool if isinstance(r, dict)
            ]
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_join_seed_apply mi_idx=%s chain=%s "
                    "prior_seed=%s rows=%s caller=%s",
                    mi_idx,
                    chain_targets,
                    use_prior_seed,
                    seed_rows,
                    probe_caller,
                )
            except Exception:
                pass
            dd["master_preview_join_read_full_files"] = True
        else:
            dd.pop("join_search_seed_pool", None)
            dd.pop("master_preview_stacked_join", None)
            dd["master_preview_join_read_full_files"] = True

    def _mpv_compute_progress_table_rows(
        self,
        *,
        mi_idx: int,
        master_step_idx: int,
        active_slot_indices: list[int],
        scenario_base: dict[str, Any],
        scan_paths: list[str],
        n_pick: int,
        use_max_sources: bool,
        progress_hook: Any,
        probe_caller: str,
        frozen_capture_out: dict[str, Any] | None = None,
    ) -> list[list[Any]]:
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        frozen_prior, frozen_through = self._mpv_frozen_context_for_mi(int(mi_idx))
        cap_acc: list[dict[str, Any]] | None = (
            [] if frozen_capture_out is not None else None
        )
        join_item = self._mpv_current_item_has_join_defs(int(mi_idx))
        anchor_prior = self._mpv_best_prior_table_rows_for_seed(
            int(mi_idx),
            n_pick=int(n_pick),
        )
        scen = scenario_for_stepped_preview(
            scenario_base,
            mi_idx=int(mi_idx),
            master_step_idx=int(master_step_idx),
            active_slot_indices=list(active_slot_indices),
            use_max_sources_for_current_item=bool(use_max_sources),
            carry_forward_completed_items=bool(join_item and anchor_prior),
            frozen_through_mi=frozen_through,
            frozen_prior=frozen_prior,
            frozen_capture_out=frozen_capture_out,
            frozen_capture_acc=cap_acc,
        )
        join_pool_out: list[dict[str, Any]] = []
        dd = scen.get("__debug_diag")
        if isinstance(dd, dict):
            cap_jf = self._master_debug_join_max_files()
            if cap_jf > 0:
                dd["master_preview_join_max_files"] = int(cap_jf)
            cap_mf = self._master_debug_max_files()
            if cap_mf > 0:
                dd["master_preview_max_files"] = int(cap_mf)
            base_items = list((scenario_base or {}).get("items") or [])
            if base_items:
                dd["preview_join_topology_items"] = copy.deepcopy(base_items)
            dd["join_search_pool_out"] = join_pool_out
            from svc.svc_data_agg import (  # noqa: WPS433
                master_preview_extract_item_allowlist,
            )

            allow_ix = master_preview_extract_item_allowlist(
                scenario_base, mi_idx=int(mi_idx)
            )
            if allow_ix is not None:
                dd["preview_extract_item_allowlist"] = list(allow_ix)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_cross_join_extract_allowlist "
                        "mi_idx=%s indices=%s caller=%s",
                        mi_idx,
                        allow_ix,
                        probe_caller,
                    )
                except Exception:
                    pass
            n_act = len(active_slot_indices or [])
            # 結合項目: 本番経路 + seed 方針（連鎖時のみ前項目 pool を seed に）
            self._mpv_apply_join_item_debug_diag(
                dd,
                mi_idx=int(mi_idx),
                join_item=bool(join_item),
                n_pick=int(n_pick),
                n_act=int(n_act),
                master_step_idx=int(master_step_idx),
                probe_caller=str(probe_caller),
            )
        if frozen_prior is not None and frozen_through is not None:
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_frozen compute mi_idx=%s through=%s caller=%s",
                    mi_idx,
                    frozen_through,
                    probe_caller,
                )
            except Exception:
                pass
        hook_paths_prev = getattr(self, "_mpv_progress_hook_paths", None)
        try:
            from svc.svc_data_agg import filter_file_paths_for_master_preview  # noqa: WPS433

            base_items = list((scenario_base or {}).get("items") or [])
            dd_hook = scen.get("__debug_diag")
            self._mpv_progress_hook_paths = list(
                filter_file_paths_for_master_preview(
                    list(scan_paths or []),
                    base_items,
                    dd_hook if isinstance(dd_hook, dict) else None,
                )
            )
            hook_q: queue.SimpleQueue | None = None
            eff_hook = progress_hook
            if self._master_should_offthread_compute(str(probe_caller or "")):
                hook_q = queue.SimpleQueue()
                eff_hook = self._master_bridge_progress_hook(progress_hook, hook_q)

            def _do_preview_compute() -> tuple[
                list[str], list[list[Any]], list[list[Any]], int
            ]:
                iter_ctx: list[dict[str, Any]] = []
                with self._mpv_item_wb_bind(int(mi_idx)):
                    result = run_preview_compute(
                        scen,
                        scan_paths,
                        max_primary_rows=self._master_preview_display_rows(),
                        max_table_rows=self._master_preview_display_rows(),
                        progress_hook=eff_hook,
                        probe_caller=probe_caller,
                        cancel_check=self._master_run_cancel_check(),
                        iteration_contexts_out=iter_ctx,
                    )
                if iter_ctx:
                    dd_local = scen.get("__debug_diag")
                    if isinstance(dd_local, dict):
                        dd_local["_mpv_iteration_contexts"] = list(iter_ctx)
                return result

            try:
                if hook_q is not None:
                    _h, table_rows, _ev, _jt = self._master_run_blocking_with_ui_pump(
                        _do_preview_compute,
                        progress_hook=progress_hook,
                        hook_q=hook_q,
                    )
                else:
                    _h, table_rows, _ev, _jt = _do_preview_compute()
            except DataAggCancelled:
                self._master_note_cancel_requested()
                raise
        finally:
            self._mpv_progress_hook_paths = hook_paths_prev
        rows_out = [list(r) for r in table_rows]
        rows_out = self._mpv_coalesce_join_compute_rows(
            mi_idx=int(mi_idx),
            rows=rows_out,
            master_step_idx=int(master_step_idx),
            n_pick=int(n_pick),
        )
        iter_ctx: list[dict[str, Any]] = []
        if isinstance(dd, dict):
            raw_ctx = dd.get("_mpv_iteration_contexts")
            if isinstance(raw_ctx, list):
                iter_ctx = [c for c in raw_ctx if isinstance(c, dict)]
        n_act_paths = len(active_slot_indices or [])
        item_complete_paths = n_act_paths <= 0 or int(master_step_idx) >= n_act_paths
        if item_complete_paths and not join_item:
            self._mpv_store_row_file_paths_for_mi(
                int(mi_idx),
                rows_out,
                iteration_contexts=iter_ctx or None,
            )
        self._mpv_note_item_stats(scen)
        if frozen_capture_out is not None and not join_item:
            self._mpv_store_frozen_snapshot(frozen_capture_out)
        if join_item and join_pool_out:
            from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                master_preview_join_compute_rows_acceptable,
            )

            item_complete = self._mpv_join_item_complete(
                mi_idx=int(mi_idx),
                master_step_idx=int(master_step_idx),
            )
            prior_peak = self._mpv_prior_peak_rows_before_mi(int(mi_idx))
            seed_ok = self._mpv_join_result_usable(
                rows_out,
                mi_idx=int(mi_idx),
                prior_peak_rows=prior_peak,
            ) or (
                not item_complete
                and master_preview_join_compute_rows_acceptable(
                    new_rows=len(rows_out),
                    prior_peak_rows=prior_peak,
                    item_complete=False,
                )
            )
            if seed_ok:
                pool_copy = [dict(r) for r in join_pool_out if isinstance(r, dict)]
                self._mpv_join_pool_by_mi[int(mi_idx)] = pool_copy
                self._mpv_join_search_pool_seed = list(pool_copy)
                self._mpv_join_search_pool_seed_paths_count = len(scan_paths)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_join_pool_seed stored mi_idx=%s rows=%s paths=%s",
                        mi_idx,
                        len(self._mpv_join_search_pool_seed),
                        self._mpv_join_search_pool_seed_paths_count,
                    )
                except Exception:
                    pass
            else:
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_join_pool_seed skip=unacceptable "
                        "mi_idx=%s emit_rows=%s prior_peak=%s pool_out=%s",
                        mi_idx,
                        len(rows_out),
                        prior_peak,
                        len(join_pool_out),
                    )
                except Exception:
                    pass
        return rows_out

    def _mpv_store_step_cache(
        self,
        sk: tuple[Any, ...],
        rows: list[list[Any]],
        *,
        mi_idx: int,
        master_step_idx: int,
    ) -> None:
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_compute_rows_acceptable,
        )

        n_act = len(self._active_slot_indices or [])
        item_complete = n_act <= 0 or int(master_step_idx) >= n_act
        join_item = self._mpv_current_item_has_join_defs(int(mi_idx))
        if join_item and rows:
            prior_peak = self._mpv_prior_peak_rows_before_mi(int(mi_idx))
            if not master_preview_join_compute_rows_acceptable(
                new_rows=len(rows),
                prior_peak_rows=prior_peak,
                item_complete=item_complete,
            ):
                prev = self._mpv_progress_rows_step_cache.get(sk)
                if prev:
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_progress skip=thin_join_step_cache "
                            "mi_idx=%s step_idx=%s new_rows=%s prev_rows=%s prior_peak=%s",
                            mi_idx,
                            master_step_idx,
                            len(rows),
                            len(prev),
                            prior_peak,
                        )
                    except Exception:
                        pass
                    return
            if item_complete and not self._mpv_join_result_usable(
                rows,
                mi_idx=int(mi_idx),
                prior_peak_rows=prior_peak,
            ):
                prev = self._mpv_progress_rows_step_cache.get(sk)
                if prev:
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_progress skip=empty_host_join_cache "
                            "mi_idx=%s step_idx=%s new_rows=%s prev_rows=%s",
                            mi_idx,
                            master_step_idx,
                            len(rows),
                            len(prev),
                        )
                    except Exception:
                        pass
                    return
        if not rows:
            prev = self._mpv_progress_rows_step_cache.get(sk)
            if prev:
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress skip=empty_step_cache_overwrite "
                        "mi_idx=%s step_idx=%s prev_rows=%s",
                        mi_idx,
                        master_step_idx,
                        len(prev),
                    )
                except Exception:
                    pass
                return
        self._mpv_progress_rows_step_cache[sk] = [list(r) for r in rows]
        self._mpv_progress_rows_by_mi[int(mi_idx)] = (
            int(master_step_idx),
            list(rows),
        )
        # 結合項目は seed の __file_path を維持。品名等は全スロット完了時のみ行パスを確定。
        if not join_item and item_complete:
            self._mpv_store_row_file_paths_for_mi(int(mi_idx), list(rows))
        self._mpv_note_progress_row_peak(int(mi_idx), len(rows))
        if rows and item_complete:
            if join_item:
                prior_peak = self._mpv_prior_peak_rows_before_mi(int(mi_idx))
                if self._mpv_join_result_usable(
                    rows,
                    mi_idx=int(mi_idx),
                    prior_peak_rows=prior_peak,
                ):
                    self._last_master_completed_mi_idx = int(mi_idx)
            else:
                self._last_master_completed_mi_idx = int(mi_idx)

    def _mpv_schedule_step_cache_backfill(
        self,
        *,
        cancel_gen: int,
        mi_idx: int,
        n_act: int,
        scenario_base: dict[str, Any],
        active_copy: list[int],
        paths_copy: list[str],
    ) -> None:
        """一括 compute 後、未到達 n_pick の段階行をバックグラウンドでキャッシュする。"""
        if n_act < 2:
            return
        if self._mpv_master_run_is_continuous():
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress backfill=skip reason=continuous_run "
                    "mi_idx=%s n_act=%s",
                    mi_idx,
                    n_act,
                )
            except Exception:
                pass
            return
        scen_id = id(scenario_base)
        jobs: list[tuple[tuple[Any, ...], int]] = []
        for k in range(1, n_act):
            sk = self._mpv_progress_step_cache_key_for(
                k,
                mi_idx=mi_idx,
                active_slot_indices=active_copy,
                scenario_id=scen_id,
                scan_paths=paths_copy,
            )
            if sk not in self._mpv_progress_rows_step_cache:
                jobs.append((sk, k))

        if not jobs:
            return

        def _worker() -> None:
            for sk, k in jobs:
                if cancel_gen != self._mpv_prefetch_cancel_gen:
                    return
                if sk in self._mpv_progress_rows_step_cache:
                    continue
                if not self._mpv_prog_compute_lock.acquire(blocking=False):
                    return
                try:
                    if cancel_gen != self._mpv_prefetch_cancel_gen:
                        continue
                    if sk in self._mpv_progress_rows_step_cache:
                        continue
                    rows = self._mpv_compute_progress_table_rows(
                        mi_idx=mi_idx,
                        master_step_idx=k,
                        active_slot_indices=active_copy,
                        scenario_base=scenario_base,
                        scan_paths=paths_copy,
                        n_pick=k,
                        use_max_sources=False,
                        progress_hook=None,
                        probe_caller="mpv_progress_backfill",
                    )
                except Exception:
                    _logger.exception("mpv progress backfill failed n_pick=%s", k)
                    continue
                finally:
                    self._mpv_prog_compute_lock_release()
                dlg = self
                QTimer.singleShot(
                    0,
                    lambda cg=cancel_gen, key=sk, r=rows, kk=k, m=mi_idx: dlg._mpv_backfill_apply(
                        cg, key, r, kk, m
                    ),
                )

        threading.Thread(
            target=_worker,
            daemon=True,
            name="mpv_prog_backfill",
        ).start()

    def _mpv_backfill_apply(
        self,
        cancel_gen: int,
        sk: tuple[Any, ...],
        rows: list[list[Any]],
        n_pick: int,
        mi_idx: int,
    ) -> None:
        if cancel_gen != self._mpv_prefetch_cancel_gen:
            return
        if sk in self._mpv_progress_rows_step_cache:
            return
        self._mpv_progress_rows_step_cache[sk] = [list(r) for r in rows]
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_progress backfill=done mi_idx=%s n_pick=%s rows=%s",
                mi_idx,
                n_pick,
                len(rows),
            )
        except Exception:
            pass

    def _mpv_can_colvals_from_progress(self) -> bool:
        """進捗表が既にキャッシュされているときだけ列切り出し（未キャッシュ時の compute_batch 起動はしない）。"""
        if not self._scenario_for_dry_run or not self._debug_scan_paths:
            return False
        n_act = len(self._active_slot_indices or [])
        if n_act <= 0:
            return False
        mi = int(self._mi_idx)
        n_pick = self._mpv_progress_n_pick()
        if n_pick > 0:
            sk = self._mpv_progress_step_cache_key(n_pick)
            cached = self._mpv_progress_rows_step_cache.get(sk)
            if cached and self._mpv_step_cached_rows_acceptable(
                cached, mi_idx=mi, n_pick=n_pick
            ):
                return True
        if n_act == 1:
            sk1 = self._mpv_progress_step_cache_key(1)
            cached1 = self._mpv_progress_rows_step_cache.get(sk1)
            if cached1 and self._mpv_step_cached_rows_acceptable(
                cached1, mi_idx=mi, n_pick=1
            ):
                return True
        key = self._mpv_progress_cache_key()
        cache = self._mpv_progress_rows_cache
        if cache is not None and cache[0] == key and cache[1]:
            return True
        ent = self._mpv_progress_rows_by_mi.get(mi)
        if ent and ent[1] and int(ent[0]) >= n_act:
            return True
        if self._mpv_one_shot_eligible():
            sk_full = self._mpv_progress_step_cache_key(n_act)
            cached_f = self._mpv_progress_rows_step_cache.get(sk_full)
            if cached_f and self._mpv_step_cached_rows_acceptable(
                cached_f, mi_idx=mi, n_pick=n_act
            ):
                return True
        return False

    def _mpv_master_run_is_continuous(self) -> bool:
        return bool(
            getattr(self, "_continuous_busy", False)
            or getattr(self, "_master_step_loop_busy", False)
        )

    def _bump_mpv_prefetch_cancel(self) -> None:
        self._mpv_prefetch_cancel_gen += 1
        self._mpv_single_slot_prefetch_pending_sk = None

    def _ensure_mpv_prefetch_worker(self) -> None:
        with self._mpv_prefetch_worker_lock:
            if self._mpv_prefetch_worker_started:
                return
            self._mpv_prefetch_worker_started = True
            threading.Thread(
                target=self._mpv_prefetch_worker_loop,
                daemon=True,
                name="mpv_prog_prefetch",
            ).start()

    def _mpv_prefetch_worker_loop(self) -> None:
        while True:
            job = self._mpv_prefetch_q.get()
            if job is None:
                break
            if len(job) >= 11:
                (
                    cancel_gen,
                    sk_next,
                    scenario_base,
                    mi_idx,
                    next_master_step,
                    active_copy,
                    paths_copy,
                    max_pr,
                    max_tr,
                    use_max_sources,
                    schedule_backfill,
                ) = job
            else:
                (
                    cancel_gen,
                    sk_next,
                    scenario_base,
                    mi_idx,
                    next_master_step,
                    active_copy,
                    paths_copy,
                    max_pr,
                    max_tr,
                ) = job
                use_max_sources = False
                schedule_backfill = False
            if cancel_gen != self._mpv_prefetch_cancel_gen:
                continue
            dlg = self
            if dlg._master_cancel_pending():
                continue
            if sk_next in self._mpv_progress_rows_step_cache:
                continue
            if not dlg._mpv_prog_compute_lock.acquire(blocking=True, timeout=2.5):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress prefetch=requeue_busy "
                        "n_pick=%s",
                        sk_next[1] if len(sk_next) > 1 else "?",
                    )
                except Exception:
                    pass
                try:
                    dlg._mpv_prefetch_q.put_nowait(job)
                except queue.Full:
                    pass
                time.sleep(0.03)
                continue
            out_rows: list[list[Any]] = []
            apply_after = False
            try:
                if cancel_gen != dlg._mpv_prefetch_cancel_gen:
                    pass
                elif dlg._master_cancel_pending():
                    pass
                elif sk_next in dlg._mpv_progress_rows_step_cache:
                    pass
                else:
                    out_rows = dlg._mpv_compute_progress_table_rows(
                        mi_idx=mi_idx,
                        master_step_idx=next_master_step,
                        active_slot_indices=active_copy,
                        scenario_base=scenario_base,
                        scan_paths=paths_copy,
                        n_pick=int(next_master_step),
                        use_max_sources=bool(use_max_sources),
                        progress_hook=None,
                        probe_caller="mpv_progress_prefetch",
                    )
                    apply_after = True
                    if (
                        apply_after
                        and cancel_gen == dlg._mpv_prefetch_cancel_gen
                        and sk_next not in dlg._mpv_progress_rows_step_cache
                    ):
                        copied = [list(r) for r in out_rows]
                        dlg._mpv_progress_rows_step_cache[sk_next] = copied
                        dlg._mpv_note_progress_row_peak(int(mi_idx), len(copied))
                        dlg._mpv_clear_single_slot_prefetch_pending(sk_next)
                        try:
                            _data_agg_probe_log.info(
                                "[DATA_AGG_DIAG] mpv_progress prefetch=done "
                                "n_pick=%s rows=%s one_shot=%s sync_cache=1",
                                sk_next[1] if len(sk_next) > 1 else "?",
                                len(copied),
                                schedule_backfill,
                            )
                        except Exception:
                            pass
            except Exception:
                _logger.exception("mpv progress prefetch failed")
            finally:
                dlg._mpv_prog_compute_lock_release()
            if apply_after:
                dlg = self
                QTimer.singleShot(
                    0,
                    lambda cg=cancel_gen,
                    k=sk_next,
                    r=out_rows,
                    m=mi_idx,
                    na=next_master_step,
                    sb=schedule_backfill,
                    sc=scenario_base,
                    ac=active_copy,
                    pc=paths_copy: dlg._mpv_prefetch_apply(
                        cg, k, r, m, na, sb, sc, ac, pc
                    ),
                )

    def _mpv_prefetch_apply(
        self,
        cancel_gen: int,
        sk: tuple[Any, ...],
        rows: list[list[Any]],
        mi_idx: int,
        next_master_step: int,
        schedule_backfill: bool,
        scenario_base: dict[str, Any],
        active_copy: list[int],
        paths_copy: list[str],
    ) -> None:
        if cancel_gen != self._mpv_prefetch_cancel_gen:
            return
        cached = self._mpv_progress_rows_step_cache.get(sk)
        if cached is None:
            self._mpv_progress_rows_step_cache[sk] = [list(r) for r in rows]
            self._mpv_clear_single_slot_prefetch_pending(sk)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress prefetch=done n_pick=%s rows=%s "
                    "one_shot=%s sync_cache=0",
                    sk[1] if len(sk) > 1 else "?",
                    len(rows),
                    schedule_backfill,
                )
            except Exception:
                pass
        else:
            self._mpv_clear_single_slot_prefetch_pending(sk)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress prefetch=done_already n_pick=%s "
                    "rows=%s",
                    sk[1] if len(sk) > 1 else "?",
                    len(cached),
                )
            except Exception:
                pass
        if schedule_backfill:
            n_act = len(active_copy)
            self._mpv_schedule_step_cache_backfill(
                cancel_gen=cancel_gen,
                mi_idx=int(mi_idx),
                n_act=n_act,
                scenario_base=scenario_base,
                active_copy=active_copy,
                paths_copy=paths_copy,
            )
            return
        tail = sk[-1] if isinstance(sk, tuple) and sk else ()
        n_act = len(tail) if isinstance(tail, tuple) else 0
        sk_parts: Sequence[Any] = sk
        try:
            done_np = int(sk_parts[1]) if len(sk_parts) > 1 else 0
        except (TypeError, ValueError, IndexError):
            done_np = 0
        if n_act > 0 and 0 < done_np < n_act:
            self._mpv_maybe_enqueue_progress_prefetch(
                next_master_step_override=done_np + 1
            )

    def _mpv_request_progress_prefetch_debounced(self) -> None:
        from core import core_env

        if not core_env.data_agg_master_progress_prefetch_enabled():
            return
        if self._mode != 1:
            return
        self._mpv_prefetch_debounce_timer.start(self._mpv_prefetch_debounce_ms)

    def _mpv_prefetch_debounced_fire(self) -> None:
        self._mpv_maybe_enqueue_progress_prefetch()

    def _mpv_maybe_enqueue_progress_prefetch(
        self,
        *,
        next_master_step_override: int | None = None,
        force: bool = False,
    ) -> None:
        from core import core_env

        if not force and not core_env.data_agg_master_progress_prefetch_enabled():
            return
        if self._mode != 1:
            return
        if not self._scenario_for_dry_run or not self._debug_scan_paths:
            return
        if self._mpv_current_item_has_join_defs(int(self._mi_idx)):
            return
        act = self._active_slot_indices or []
        n_act = len(act)
        if n_act <= 0:
            return
        one_shot = self._mpv_one_shot_eligible()
        if one_shot:
            next_master_step = n_act
            use_max_sources = True
            schedule_backfill = not self._mpv_master_run_is_continuous()
        elif next_master_step_override is not None:
            next_master_step = int(next_master_step_override)
            use_max_sources = False
            schedule_backfill = False
        else:
            step_idx = int(self._master_step_idx)
            next_master_step = step_idx + 1
            use_max_sources = False
            schedule_backfill = False
        if next_master_step > n_act or next_master_step < 1:
            return
        sk_next = self._mpv_progress_step_cache_key(next_master_step)
        if sk_next in self._mpv_progress_rows_step_cache:
            return
        cancel_gen = self._mpv_prefetch_cancel_gen
        scenario_base = copy.deepcopy(self._scenario_for_dry_run or {})
        mi_idx = int(self._mi_idx)
        active_copy = list(act)
        paths_copy = list(self._debug_scan_paths)
        max_pr = self._master_preview_display_rows()
        max_tr = self._master_preview_display_rows()
        job = (
            cancel_gen,
            sk_next,
            scenario_base,
            mi_idx,
            next_master_step,
            active_copy,
            paths_copy,
            max_pr,
            max_tr,
            use_max_sources,
            schedule_backfill,
        )
        try:
            self._mpv_prefetch_q.put_nowait(job)
        except queue.Full:
            try:
                self._mpv_prefetch_q.get_nowait()
            except queue.Empty:
                pass
            try:
                self._mpv_prefetch_q.put_nowait(job)
            except queue.Full:
                pass
        self._ensure_mpv_prefetch_worker()

    def _mpv_invalidate_final_table_rows(self) -> None:
        self._mpv_final_table_rows = None
        self._mpv_final_grid_applied = False

    def _mpv_should_defer_column_fit(self) -> bool:
        return bool(
            getattr(self, "_continuous_busy", False)
            or getattr(self, "_master_step_loop_busy", False)
        )

    def _mpv_flush_deferred_column_fit_if_needed(self) -> None:
        if not getattr(self, "_mpv_column_fit_pending", False):
            return
        self._mpv_column_fit_pending = False
        self._fit_value_grid_columns()

    def _mpv_finalize_target_mi(self) -> int | None:
        """連続実行完了時の最終 compute 対象 mi（最後に到達した項目を優先）。"""
        disp = self._mpv_last_completed_mi_for_display()
        fb = getattr(self, "_last_master_completed_mi_idx", None)
        if fb is not None and int(fb) >= 0:
            if disp is None or int(fb) >= int(disp):
                return int(fb)
        if disp is not None:
            return int(disp)
        if fb is not None and int(fb) >= 0:
            return int(fb)
        return None

    def _mpv_finalize_step_cache_acceptable(self) -> bool:
        """連続実行完了時に step キャッシュで最終表示できるか。"""
        done_mi = self._mpv_finalize_target_mi()
        if done_mi is None:
            return False
        mi_saved = int(self._mi_idx)
        step_saved = int(self._master_step_idx)
        try:
            self._mi_idx = int(done_mi)
            self._rebuild_active_slots()
            n_act = len(self._active_slot_indices or [])
            if n_act <= 0:
                ent = self._mpv_progress_rows_by_mi.get(int(done_mi))
                return bool(ent and ent[1])
            rows = self._mpv_rows_from_step_cache_n_pick(n_act)
            if rows and self._mpv_current_item_has_join_defs(int(done_mi)):
                prior_peak = self._mpv_prior_peak_rows_before_mi(int(done_mi))
                if not self._mpv_join_result_usable(
                    rows,
                    mi_idx=int(done_mi),
                    prior_peak_rows=prior_peak,
                ):
                    return False
            if rows:
                return True
            ent = self._mpv_progress_rows_by_mi.get(int(done_mi))
            return bool(
                ent
                and ent[1]
                and int(ent[0]) >= n_act
                and self._mpv_step_cached_rows_acceptable(
                    ent[1],
                    mi_idx=int(done_mi),
                    n_pick=n_act,
                )
                and (
                    not self._mpv_current_item_has_join_defs(int(done_mi))
                    or self._mpv_join_result_usable(
                        ent[1],
                        mi_idx=int(done_mi),
                        prior_peak_rows=self._mpv_prior_peak_rows_before_mi(
                            int(done_mi)
                        ),
                    )
                )
            )
        finally:
            self._mi_idx = mi_saved
            self._master_step_idx = step_saved
            self._rebuild_active_slots()

    def _mpv_last_completed_mi_for_display(self) -> int | None:
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_compute_rows_acceptable,
        )

        by_mi = getattr(self, "_mpv_progress_rows_by_mi", {}) or {}
        fb = getattr(self, "_last_master_completed_mi_idx", None)
        best_mi: int | None = None

        def _mi_usable(mi: int, ent: tuple[int, list[list[Any]]] | None) -> bool:
            if not ent or not ent[1]:
                return False
            prior_peak = self._mpv_prior_peak_rows_before_mi(int(mi))
            if not self._mpv_current_item_has_join_defs(int(mi)):
                return master_preview_join_compute_rows_acceptable(
                    new_rows=len(ent[1]),
                    prior_peak_rows=prior_peak,
                    item_complete=True,
                )
            return self._mpv_join_result_usable(
                ent[1],
                mi_idx=int(mi),
                prior_peak_rows=prior_peak,
            )

        scan_from = int(fb) if fb is not None and int(fb) >= 0 else -1
        for mi in range(scan_from, -1, -1):
            ent = by_mi.get(mi)
            if _mi_usable(mi, ent):
                best_mi = int(mi)
                break
        if best_mi is None:
            for mi in sorted(by_mi.keys(), reverse=True):
                ent = by_mi.get(int(mi))
                if _mi_usable(int(mi), ent):
                    best_mi = int(mi)
                    break
        if best_mi is not None:
            return best_mi
        peaks = getattr(self, "_mpv_progress_row_peak_by_mi", {}) or {}
        best_peak = 0
        for mi, peak in peaks.items():
            p = int(peak)
            ent = by_mi.get(int(mi))
            if p > best_peak and ent and ent[1]:
                best_peak = p
                best_mi = int(mi)
        if best_mi is not None:
            return best_mi
        if fb is not None and int(fb) >= 0:
            return int(fb)
        return None

    def _mpv_build_final_table_rows(self, *, force_recompute: bool = False) -> list[list[Any]]:
        """デバッグ完了時: 最後に完了した項目で全列フル compute（行順は compute_batch の file/iter ソート）。"""
        if (
            not force_recompute
            and isinstance(getattr(self, "_mpv_final_table_rows", None), list)
            and self._mpv_final_table_rows
        ):
            return list(self._mpv_final_table_rows)
        done_mi = self._mpv_finalize_target_mi()
        if (
            done_mi is None
            or not self._scenario_for_dry_run
            or not self._debug_scan_paths
        ):
            self._mpv_final_table_rows = []
            return []
        mi_saved = int(self._mi_idx)
        step_saved = int(self._master_step_idx)
        rows: list[list[Any]] = []
        try:
            self._mi_idx = int(done_mi)
            self._rebuild_active_slots()
            act = list(self._active_slot_indices or [])
            if not act:
                ent = self._mpv_progress_rows_by_mi.get(int(done_mi))
                if ent and ent[1]:
                    rows = [list(r) for r in ent[1]]
            else:
                n_act = len(act)
                cached_full = self._mpv_rows_from_step_cache_n_pick(n_act)
                global_peak = max(
                    (
                        int(v)
                        for v in getattr(
                            self, "_mpv_progress_row_peak_by_mi", {}
                        ).values()
                    ),
                    default=0,
                )
                if (
                    cached_full
                    and global_peak >= 10
                    and len(cached_full) < max(global_peak // 4, 2)
                ):
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_final_table_rows skip=thin_cache "
                            "done_mi=%s cached_rows=%s global_peak=%s",
                            done_mi,
                            len(cached_full),
                            global_peak,
                        )
                    except Exception:
                        pass
                    cached_full = None
                if cached_full and self._mpv_current_item_has_join_defs(int(done_mi)):
                    prior_peak = self._mpv_prior_peak_rows_before_mi(int(done_mi))
                    if not self._mpv_join_result_usable(
                        cached_full,
                        mi_idx=int(done_mi),
                        prior_peak_rows=prior_peak,
                    ):
                        try:
                            _data_agg_probe_log.info(
                                "[DATA_AGG_DIAG] mpv_final_table_rows skip=empty_host_cache "
                                "done_mi=%s cached_rows=%s",
                                done_mi,
                                len(cached_full),
                            )
                        except Exception:
                            pass
                        cached_full = None
                ent = self._mpv_progress_rows_by_mi.get(int(done_mi))
                if cached_full:
                    rows = cached_full
                elif (
                    ent
                    and ent[1]
                    and int(ent[0]) >= n_act
                    and self._mpv_step_cached_rows_acceptable(
                        ent[1],
                        mi_idx=int(done_mi),
                        n_pick=n_act,
                    )
                    and (
                        not self._mpv_current_item_has_join_defs(int(done_mi))
                        or self._mpv_join_result_usable(
                            ent[1],
                            mi_idx=int(done_mi),
                            prior_peak_rows=self._mpv_prior_peak_rows_before_mi(
                                int(done_mi)
                            ),
                        )
                    )
                ):
                    rows = [list(r) for r in ent[1]]
                else:
                    recomputed = self._mpv_ensure_step_n_pick_cached(
                        n_pick=n_act,
                        progress_hook=None,
                        probe_caller="mpv_final_display",
                    )
                    rows = [list(r) for r in (recomputed or [])]
                    sk_full = self._mpv_progress_step_cache_key(n_act)
                    self._mpv_store_step_cache(
                        sk_full,
                        rows,
                        mi_idx=int(done_mi),
                        master_step_idx=n_act,
                    )
        finally:
            self._mi_idx = mi_saved
            self._master_step_idx = step_saved
            self._rebuild_active_slots()
        ordered = self._mpv_apply_aggregation_row_order([list(r) for r in rows])
        self._mpv_final_table_rows = ordered
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_final_table_rows mi=%s rows=%s recompute=%s",
                done_mi,
                len(ordered),
                force_recompute,
            )
        except Exception:
            pass
        return list(self._mpv_final_table_rows)

    def _mpv_apply_final_result_grid(self, *, force_recompute: bool = False) -> None:
        """完了後の結果一覧を従来どおり file/iter 順の結合表で表示する。"""
        if self._mode != 1 or not self._scenario_for_dry_run:
            return
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_finalize_should_force_recompute,
        )

        cache_ok = self._mpv_finalize_step_cache_acceptable()
        effective_force = (
            force_recompute
            and master_preview_finalize_should_force_recompute(
                step_cache_hit=cache_ok
            )
        )
        if (
            getattr(self, "_mpv_final_grid_applied", False)
            and isinstance(getattr(self, "_mpv_final_table_rows", None), list)
            and self._mpv_final_table_rows
            and not effective_force
        ):
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_finalize_skip_duplicate_grid_apply "
                    "rows=%s cache_ok=%s",
                    len(self._mpv_final_table_rows),
                    cache_ok,
                )
            except Exception:
                pass
            self._mpv_flush_deferred_column_fit_if_needed()
            return
        self._mpv_build_final_table_rows(force_recompute=effective_force)
        self._mpv_progress_rows_cache = None
        self._mpv_display_mi_idx = None
        self._mpv_show_merged_current = False
        self._mpv_join_table_active = False
        self._mpv_join_table_ncols = 0
        self._mpv_final_grid_applied = True
        self._rebuild_value_grid()
        self._mpv_flush_deferred_column_fit_if_needed()

    def _mpv_apply_aggregation_row_order(
        self, rows: list[list[Any]]
    ) -> list[list[Any]]:
        """デバッグ結果一覧を本番集約と同じ縦順（paths 順 + excel sort_keys）に揃える。"""
        if self._mode != 1 or not rows or not self._scenario_for_dry_run:
            return rows
        headers = self._mpv_preview_headers()
        if not headers:
            return rows
        from svc.svc_data_agg import apply_master_preview_table_row_order  # noqa: WPS433

        return apply_master_preview_table_row_order(
            self._scenario_for_dry_run or {},
            headers,
            [list(r) for r in rows],
        )

    def _mpv_progress_batch_rows(self) -> list[list[Any]]:
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        try:
            rows = self._mpv_progress_batch_rows_impl()
            return self._mpv_apply_aggregation_row_order(rows)
        except DataAggCancelled:
            cache = getattr(self, "_mpv_progress_rows_cache", None)
            if cache is not None and cache[1]:
                return [list(r) for r in cache[1]]
            last = list(getattr(self, "_mpv_last_valid_table_rows", None) or [])
            return [list(r) for r in last]
        finally:
            if not getattr(self, "_master_abort_in_progress", False):
                self._mpv_request_progress_prefetch_debounced()

    def _mpv_progress_batch_rows_impl(self) -> list[list[Any]]:
        if self._master_on_ui_thread():
            return self._master_run_blocking_with_ui_pump(
                self._mpv_progress_batch_rows_impl_core
            )
        return self._mpv_progress_batch_rows_impl_core()

    def _mpv_progress_batch_rows_impl_core(self) -> list[list[Any]]:
        """
        マスタプレビュー表示用: 完了項目はフル、現在項目は実行済みシナリオ分のみ、未到達項目は主値ソースなし。
        連携・結合は到達範囲のパイプラインに含まれる列へ反映され、未到達列へのフル結果の透けを防ぐ。
        """
        if (
            self._mode == 1
            and getattr(self, "_master_step_pass_complete", False)
            and self._scenario_for_dry_run
            and self._debug_scan_paths
        ):
            final_rows = self._mpv_build_final_table_rows()
            self._mpv_display_mi_idx = None
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress use=final_table_rows rows=%s",
                    len(final_rows),
                )
            except Exception:
                pass
            return final_rows
        key = self._mpv_progress_cache_key()
        cache = self._mpv_progress_rows_cache
        if cache is not None and cache[0] == key:
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress cache=hit mi_idx=%s step_idx=%s rows=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    len(cache[1]),
                )
            except Exception:
                pass
            self._mpv_maybe_enqueue_progress_prefetch()
            return cache[1]
        if self._active_slot_indices:
            join_disp = self._mpv_try_join_step0_display_rows(key)
            if join_disp is not None:
                return join_disp
        if self._active_slot_indices:
            n_pick_req = self._mpv_progress_n_pick()
            sk = self._mpv_progress_step_cache_key(n_pick_req)
            cached_step = self._mpv_progress_rows_step_cache.get(sk)
            if cached_step is not None and self._mpv_step_cached_rows_acceptable(
                cached_step,
                mi_idx=int(self._mi_idx),
                n_pick=int(n_pick_req),
            ):
                rows = [list(r) for r in cached_step]
                self._mpv_display_mi_idx = int(self._mi_idx)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress cache=hit_step mi_idx=%s step_idx=%s n_pick=%s rows=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        n_pick_req,
                        len(rows),
                    )
                except Exception:
                    pass
                self._mpv_progress_rows_by_mi[int(self._mi_idx)] = (
                    int(self._master_step_idx),
                    list(rows),
                )
                n_act_hit = len(self._active_slot_indices)
                if rows and int(self._master_step_idx) >= n_act_hit:
                    self._last_master_completed_mi_idx = int(self._mi_idx)
                self._mpv_progress_rows_cache = (key, rows)
                self._mpv_maybe_enqueue_progress_prefetch()
                return rows
        if self._master_step_idx >= len(self._active_slot_indices) and self._active_slot_indices:
            prev = self._mpv_progress_rows_by_mi.get(int(self._mi_idx))
            n_act = len(self._active_slot_indices)
            if prev is not None and prev[1]:
                po = int(prev[0])
                # 単一スロット項目で po==0 のキャッシュ（スロット実行前の一覧）を、
                # step_idx>=n_act の「項目完了直後」に再利用すると出荷番号列が空のままになるため、
                # 全スロットが計算に反映されたときの行（prev_step >= n_act）に限定する。
                if po >= n_act:
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_progress reuse=%s mi_idx=%s step_idx=%s prev_step=%s rows=%s",
                            "end_of_item_full",
                            self._mi_idx,
                            self._master_step_idx,
                            prev[0],
                            len(prev[1]),
                        )
                    except Exception:
                        pass
                    rows = list(prev[1])
                    self._mpv_display_mi_idx = int(self._mi_idx)
                    self._mpv_progress_rows_cache = (key, rows)
                    return rows
        # 実行可能スロットが無い行（シナリオ未登録など）では compute_batch を呼ばない。
        # 将来列の透けを防ぐため fb_mi>cur_mi は使わず、cur_mi 以下でキャッシュがある最大 mi の行だけ再利用する。
        if not self._active_slot_indices:
            fb_mi = getattr(self, "_last_master_completed_mi_idx", None)
            cur_mi = int(self._mi_idx)
            if fb_mi is not None and fb_mi <= cur_mi:
                ent = self._mpv_progress_rows_by_mi.get(int(fb_mi))
                if ent is not None and ent[1]:
                    rows = list(ent[1])
                    self._mpv_display_mi_idx = int(fb_mi)
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_progress reuse=last_completed_mi "
                            "mi_idx=%s step_idx=%s fb_mi=%s rows=%s",
                            self._mi_idx,
                            self._master_step_idx,
                            fb_mi,
                            len(rows),
                        )
                    except Exception:
                        pass
                    self._mpv_progress_rows_cache = (key, rows)
                    return rows
            by_mi = getattr(self, "_mpv_progress_rows_by_mi", None) or {}
            pick_mi: int | None = None
            for m in sorted(by_mi.keys(), reverse=True):
                if m <= cur_mi:
                    e = by_mi.get(m)
                    if e and e[1]:
                        pick_mi = int(m)
                        break
            if pick_mi is not None:
                ent = by_mi[pick_mi]
                rows = list(ent[1])
                self._mpv_display_mi_idx = int(pick_mi)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress reuse=best_cached_mi_le_cur "
                        "mi_idx=%s step_idx=%s pick_mi=%s rows=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        pick_mi,
                        len(rows),
                    )
                except Exception:
                    pass
                self._mpv_progress_rows_cache = (key, rows)
                return rows
            rows = []
            self._mpv_display_mi_idx = int(self._mi_idx)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress skip=no_active_slots mi_idx=%s step_idx=%s",
                    self._mi_idx,
                    self._master_step_idx,
                )
            except Exception:
                pass
            self._mpv_progress_rows_cache = (key, rows)
            return rows
        n_pick_req = self._mpv_progress_n_pick()
        single_slot_early = self._mpv_try_single_slot_step0_rows(key)
        if single_slot_early is not None:
            return single_slot_early
        if self._mpv_is_single_slot_active() and int(n_pick_req) == 1:
            cached_one = self._mpv_rows_from_step_cache_n_pick(1)
            if cached_one:
                self._mpv_display_mi_idx = int(self._mi_idx)
                self._mpv_progress_rows_cache = (key, cached_one)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress cache=hit_step mi_idx=%s "
                        "step_idx=%s n_pick=1 rows=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        len(cached_one),
                    )
                except Exception:
                    pass
                self._mpv_maybe_enqueue_progress_prefetch()
                return cached_one
            from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                master_preview_single_slot_progress_batch_wait_ms,
            )

            wait_ms = master_preview_single_slot_progress_batch_wait_ms(
                prefetch_pending=self._mpv_is_single_slot_prefetch_pending()
            )
            waited = self._mpv_wait_single_slot_n_pick1_cache(max_wait_ms=int(wait_ms))
            if waited:
                self._mpv_display_mi_idx = int(self._mi_idx)
                self._mpv_progress_rows_cache = (key, waited)
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress cache=hit_step_after_wait "
                        "mi_idx=%s step_idx=%s n_pick=1 rows=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        len(waited),
                    )
                except Exception:
                    pass
                self._mpv_maybe_enqueue_progress_prefetch()
                return waited
            if (
                int(self._master_step_idx) == 0
                and self._mpv_is_single_slot_prefetch_pending()
            ):
                self._mpv_display_mi_idx = int(self._mi_idx)
                self._mpv_progress_rows_cache = (key, [])
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress skip=defer_prefetch_step0 "
                        "mi_idx=%s step_idx=0",
                        self._mi_idx,
                    )
                except Exception:
                    pass
                return []
        composed_rows = self._mpv_try_compose_progress_rows_from_cache(n_pick=n_pick_req)
        if composed_rows is not None:
            self._mpv_progress_rows_cache = (key, composed_rows)
            sk_comp = self._mpv_progress_step_cache_key(n_pick_req)
            self._mpv_store_step_cache(
                sk_comp,
                composed_rows,
                mi_idx=int(self._mi_idx),
                master_step_idx=int(self._master_step_idx),
            )
            return composed_rows
        t0 = time.perf_counter()
        self._mpv_display_mi_idx = int(self._mi_idx)
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_progress start mi_idx=%s step_idx=%s scan_paths=%s active_slots=%s",
                self._mi_idx,
                self._master_step_idx,
                len(self._debug_scan_paths or []),
                len(self._active_slot_indices or []),
            )
        except Exception:
            pass
        sk_need = self._mpv_progress_step_cache_key(n_pick_req)
        from_prefetch_wait = False
        rows: list[list[Any]] = []
        self._master_mpv_compute_lock_acquire_ui()
        try:
            hit_after_wait = self._mpv_progress_rows_step_cache.get(sk_need)
            if hit_after_wait is not None:
                candidate = [list(r) for r in hit_after_wait]
                if candidate and self._mpv_step_cached_rows_acceptable(
                    candidate,
                    mi_idx=int(self._mi_idx),
                    n_pick=int(n_pick_req),
                ):
                    rows = candidate
                    from_prefetch_wait = True
            if not rows:
                n_act_compute = len(self._active_slot_indices or [])
                use_max = bool(
                    self._mpv_one_shot_eligible()
                    and n_pick_req == n_act_compute
                    and n_act_compute > 0
                )
                frozen_cap: dict[str, Any] | None = None
                if (
                    self._mpv_frozen_columns_enabled()
                    and not self._mpv_current_item_has_join_defs()
                    and n_act_compute > 0
                    and n_pick_req >= n_act_compute
                ):
                    frozen_cap = {}
                _off_hook = self._mpv_master_dbg_progress_hook_or_none()
                try:
                    rows = self._mpv_compute_progress_table_rows(
                        mi_idx=int(self._mi_idx),
                        master_step_idx=int(
                            self._mpv_effective_master_step_for_preview()
                        ),
                        active_slot_indices=list(self._active_slot_indices),
                        scenario_base=self._scenario_for_dry_run or {},
                        scan_paths=list(self._debug_scan_paths),
                        n_pick=int(n_pick_req),
                        use_max_sources=use_max,
                        progress_hook=_off_hook,
                        probe_caller="mpv_progress",
                        frozen_capture_out=frozen_cap,
                    )
                except Exception:
                    _logger.exception("mpv progress batch rows failed")
                    rows = []
        finally:
            self._master_mpv_compute_lock_release_ui()
        if (
            not rows
            and self._active_slot_indices
            and int(self._master_step_idx) >= len(self._active_slot_indices)
        ):
            n_act_fb = len(self._active_slot_indices)
            prev_ent = self._mpv_progress_rows_by_mi.get(int(self._mi_idx))
            if (
                prev_ent is not None
                and prev_ent[1]
                and int(prev_ent[0]) >= n_act_fb
                and not self._mpv_current_item_has_join_defs()
            ):
                rows = [list(r) for r in prev_ent[1]]
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress reuse=nonempty_fallback "
                        "mi_idx=%s step_idx=%s prev_step=%s rows=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        prev_ent[0],
                        len(rows),
                    )
                except Exception:
                    pass
            elif not rows and self._mpv_current_item_has_join_defs():
                fb_rows = self._mpv_try_join_step_cache_fallback_rows(
                    n_pick=int(n_pick_req)
                )
                if fb_rows:
                    rows = fb_rows
        if from_prefetch_wait:
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress cache=hit_step_after_wait "
                    "mi_idx=%s step_idx=%s n_pick=%s rows=%s elapsed_ms=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    n_pick_req,
                    len(rows),
                    int((time.perf_counter() - t0) * 1000),
                )
            except Exception:
                pass
        else:
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress cache=miss mi_idx=%s step_idx=%s rows=%s elapsed_ms=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    len(rows),
                    int((time.perf_counter() - t0) * 1000),
                )
            except Exception:
                pass
        sk_store = sk_need
        if not from_prefetch_wait:
            self._mpv_store_step_cache(
                sk_store,
                rows,
                mi_idx=int(self._mi_idx),
                master_step_idx=int(self._master_step_idx),
            )
        n_act_done = len(self._active_slot_indices or [])
        use_max_done = bool(
            self._mpv_one_shot_eligible()
            and n_pick_req == n_act_done
            and n_act_done > 0
        )
        if use_max_done and not from_prefetch_wait:
            self._mpv_schedule_step_cache_backfill(
                cancel_gen=self._mpv_prefetch_cancel_gen,
                mi_idx=int(self._mi_idx),
                n_act=n_act_done,
                scenario_base=copy.deepcopy(self._scenario_for_dry_run or {}),
                active_copy=list(self._active_slot_indices or []),
                paths_copy=list(self._debug_scan_paths),
            )
        self._mpv_progress_rows_cache = (key, rows)
        self._mpv_maybe_enqueue_progress_prefetch()
        return rows

    def _mpv_try_compose_progress_rows_from_cache(
        self, *, n_pick: int
    ) -> list[list[Any]] | None:
        if not (self._scenario_for_dry_run and self._debug_scan_paths):
            return None
        act = list(self._active_slot_indices or [])
        if len(act) != 1:
            return None
        cur_mi = int(self._mi_idx)
        if cur_mi <= 0:
            return None
        prev_ent = self._mpv_progress_rows_by_mi.get(cur_mi - 1)
        if prev_ent is None or not prev_ent[1]:
            return None
        scen_items = list((self._scenario_for_dry_run or {}).get("items") or [])
        ncols = len(scen_items)
        if ncols <= 0 or cur_mi >= ncols:
            return None
        base_rows: list[list[Any]] = []
        for r in prev_ent[1]:
            rr = list(r)
            if len(rr) < ncols:
                rr.extend([None] * (ncols - len(rr)))
            elif len(rr) > ncols:
                rr = rr[:ncols]
            base_rows.append(rr)
        if not base_rows:
            return None
        if self._mpv_current_item_has_join_defs(cur_mi):
            return None
        # step0（n_pick=0）は現在列が空の仕様。前項目 rows をそのまま再利用する。
        if int(n_pick) <= 0:
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_progress reuse=prev_item_rows_step0 mi_idx=%s step_idx=%s rows=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    len(base_rows),
                )
            except Exception:
                pass
            return base_rows
        # n_pick>0 の段階は「前項目 rows + 現在列を行インデックスで差し込み」だと
        # join/item の行順差でズレる可能性があるため、従来どおり compute_batch を使う。
        return None

    def _mpv_poll_extract_cancel(
        self, cancel_check: Callable[..., None] | None
    ) -> None:
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        if cancel_check is not None:
            cancel_check(force=True)

    def _mpv_extract_colvals_blocking(
        self,
        *,
        mi_idx: int,
        si: int,
        src: dict[str, Any],
        paths_list: list[str],
        one: dict[str, Any],
        item_id: str,
        max_rows: int,
        n_paths_before: int,
        n_paths_scenario: int,
        n_paths_after: int,
        extract_hook: Callable[..., None] | None,
        cancel_check: Callable[..., None] | None,
    ) -> list[str]:
        """extract ループ本体（ワーカー可。Qt は progress_hook ブリッジ経由のみ）。"""
        from contextlib import nullcontext

        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        try:
            from svc.svc_data_agg_extract import (
                extract_item_bundle,
                xlsx_progress_cache_mark,
                xlsx_workbook_scope,
                xlsx_workbook_scope_active,
            )
        except Exception:
            extract_item_bundle = None  # type: ignore[misc, assignment]
            xlsx_progress_cache_mark = None  # type: ignore[misc, assignment]
            xlsx_workbook_scope = None  # type: ignore[misc, assignment]
            xlsx_workbook_scope_active = None  # type: ignore[misc, assignment]
        if extract_item_bundle is None or xlsx_workbook_scope is None:
            return ["（svc_data_agg_extract を読み込めませんでした）"]
        t0 = time.perf_counter()
        col_vals: list[str] = []
        _csv_prog = _master_debug_csv_precache_progress_hook(
            extract_hook, cancel_check=cancel_check
        )
        hook_paths_prev = getattr(self, "_mpv_progress_hook_paths", None)
        self._mpv_progress_hook_paths = list(paths_list)
        try:
            with self._mpv_item_wb_bind(int(mi_idx)):
                for i, fp in enumerate(paths_list, start=1):
                    try:
                        self._mpv_poll_extract_cancel(cancel_check)
                    except DataAggCancelled:
                        raise
                    if len(col_vals) >= max_rows:
                        break
                    fname = Path(str(fp)).name
                    row_prog = min(len(col_vals) + 1, max_rows)
                    mark = ""
                    if xlsx_progress_cache_mark is not None:
                        try:
                            mark = str(xlsx_progress_cache_mark(fp) or "")
                        except Exception:
                            mark = "[F] "
                    if extract_hook is not None:
                        try:
                            extract_hook(
                                4,
                                "%s行 %s/%s: %s 読込中"
                                % (mark, row_prog, max_rows, fname),
                                i,
                                len(paths_list),
                            )
                        except DataAggCancelled:
                            raise
                        except Exception:
                            pass
                    # 項目フレーム bind 済みならファイル単位 scope は張らない（閉じで共有を壊さない）
                    _file_cm = (
                        nullcontext()
                        if (
                            xlsx_workbook_scope_active is not None
                            and xlsx_workbook_scope_active()
                        )
                        else xlsx_workbook_scope()
                    )
                    with _file_cm:  # type: ignore[misc]
                        try:
                            _precache_csv_for_master_debug_extract(
                                fp,
                                progress_hook=_csv_prog,
                            )
                            jp_hdr = str(one.get("name") or one.get("id") or "").strip()
                            b = extract_item_bundle(
                                fp,
                                one,
                                item_id=item_id or None,
                                cell_positions={},
                                join_path_header=jp_hdr or None,
                                cancel_check=cancel_check,
                            )
                        except DataAggCancelled:
                            raise
                        except Exception:
                            b = {"primary_values": []}
                    _append_extract_primaries_to_col(
                        col_vals,
                        b.get("primary_values"),
                        max_rows=max_rows,
                    )
                    if extract_hook is not None:
                        try:
                            extract_hook(
                                4,
                                "%s行 %s/%s: %s（完了）"
                                % (
                                    mark,
                                    min(len(col_vals), max_rows),
                                    max_rows,
                                    fname,
                                ),
                                i,
                                len(paths_list),
                            )
                        except DataAggCancelled:
                            raise
                        except Exception:
                            pass
        finally:
            self._mpv_progress_hook_paths = hook_paths_prev
        if not col_vals:
            col_vals = ["（該当する主値がありません）"]
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_extract mi_idx=%s si=%s src_type=%s "
                "paths_before=%s paths_scenario=%s paths_source=%s col_count=%s col_head=%s elapsed_ms=%s",
                mi_idx,
                si,
                str(src.get("type") or "cell").strip().lower(),
                n_paths_before,
                n_paths_scenario,
                n_paths_after,
                len(col_vals),
                col_vals[:5],
                int((time.perf_counter() - t0) * 1000),
            )
        except Exception:
            pass
        return col_vals

    def _mpv_extract_colvals(self, mi_idx: int, si: int) -> list[str]:
        """
        マスタプレビュー用: 本番と同じ extract_item_bundle 経路で主値列を得る。
        build_master_items_live(..., preload_values=True) と同系統（行順はファイル走査順）。
        """
        key = (mi_idx, si)
        if key in self._mpv_extract_cache:
            return list(self._mpv_extract_cache[key])
        scen = self._scenario_for_dry_run or {}
        items = list(scen.get("items") or [])
        if mi_idx < 0 or mi_idx >= len(items):
            out = ["（項目がありません）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        item = items[mi_idx]
        if not isinstance(item, dict):
            out = ["（項目定義が不正です）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        sources = item.get("sources") or []
        if si < 0 or si >= len(sources):
            out = ["（ソースがありません）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        src = sources[si]
        if not isinstance(src, dict):
            out = ["（ソース定義が不正です）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        try:
            from svc.svc_data_agg import filter_file_paths_for_master_preview
            from svc.svc_data_agg_extract import file_paths_for_source_extract
        except Exception:
            filter_file_paths_for_master_preview = None  # type: ignore[misc, assignment]
            file_paths_for_source_extract = None  # type: ignore[misc, assignment]
        paths = [str(p).strip() for p in self._debug_scan_paths if str(p).strip()]
        n_paths_before = len(paths)
        if filter_file_paths_for_master_preview is not None:
            paths = list(filter_file_paths_for_master_preview(paths, items))
        n_paths_scenario = len(paths)
        if not paths:
            out = ["（検出ファイルがありません）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        one = {**item, "sources": [copy.deepcopy(src)]}
        item_id = str(item.get("id") or item.get("name") or "").strip()
        if file_paths_for_source_extract is not None:
            paths_list = file_paths_for_source_extract(paths, src)
        else:
            paths_list = list(paths)
        join_item = self._mpv_current_item_has_join_defs(int(mi_idx))
        cap_files = (
            self._master_debug_join_max_files()
            if join_item
            else self._master_debug_max_files()
        )
        if cap_files > 0 and len(paths_list) > cap_files:
            paths_list = list(paths_list[: int(cap_files)])
        n_paths_after = len(paths_list)
        if not paths_list:
            out = ["（該当する主値がありません）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        max_rows = self._master_preview_display_rows()
        cancel_chk = self._master_run_cancel_check()
        real_hook = self._mpv_master_dbg_progress_hook_or_none()
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        def _run_extract_blocking(eff_hook: Callable[..., None] | None) -> list[str]:
            return self._mpv_extract_colvals_blocking(
                mi_idx=mi_idx,
                si=si,
                src=src,
                paths_list=list(paths_list),
                one=one,
                item_id=item_id,
                max_rows=max_rows,
                n_paths_before=n_paths_before,
                n_paths_scenario=n_paths_scenario,
                n_paths_after=n_paths_after,
                extract_hook=eff_hook,
                cancel_check=cancel_chk,
            )

        try:
            if self._master_on_ui_thread():
                hook_q: queue.SimpleQueue = queue.SimpleQueue()
                bridged = (
                    self._master_bridge_progress_hook(real_hook, hook_q)
                    if real_hook is not None
                    else None
                )

                def _worker_extract() -> list[str]:
                    return _run_extract_blocking(bridged)

                col_vals = self._master_run_blocking_with_ui_pump(
                    _worker_extract,
                    progress_hook=real_hook,
                    hook_q=hook_q,
                )
            else:
                col_vals = _run_extract_blocking(real_hook)
        except DataAggCancelled:
            self._master_note_cancel_requested()
            raise
        self._mpv_extract_cache[key] = list(col_vals)
        return list(col_vals)

    def _merge_mpv_column(self, mi_idx: int, colvals: list[str]) -> None:
        """mpv: 主値のみ項目の列を書込みモードに従ってマージ（結合項目では呼ばない）。"""
        scen = self._scenario_for_dry_run or {}
        items = list(scen.get("items") or [])
        if mi_idx < 0 or mi_idx >= len(items):
            return
        from svc.svc_data_agg_scenario import infer_item_lineage, normalize_item_write_mode
        from svc.svc_data_agg_write import merge_cell_for_write_mode

        it = items[mi_idx]
        lin = infer_item_lineage(it.get("sources") or [])
        if lin == "__mixed__":
            lin = None
        wm = normalize_item_write_mode(it.get("write_mode"), lineage=lin)
        ncols = len(items)
        display_cap = self._master_preview_display_rows()
        nrows = min(
            display_cap,
            max(
                len(colvals),
                len(self._mpv_grid or []),
            ),
        )
        if nrows < 1:
            nrows = 1
        if self._mpv_grid is None:
            self._mpv_grid = [[None] * ncols for _ in range(nrows)]
        else:
            for row in self._mpv_grid:
                while len(row) < ncols:
                    row.append(None)
            while len(self._mpv_grid) < nrows:
                self._mpv_grid.append([None] * ncols)
        for r in range(len(self._mpv_grid)):
            new_val = colvals[r] if r < len(colvals) else None
            old = self._mpv_grid[r][mi_idx]
            self._mpv_grid[r][mi_idx] = merge_cell_for_write_mode(old, new_val, wm)

    def _render_mpv_grid(self) -> None:
        """マスタプレビュー: 結果一覧は run_preview_compute の table_rows（進捗行）のみを表示する。
        _mpv_grid は描画直前に prog_rows と同内容へ同期し、extract マージで batch 結果を上書き表示しない。"""
        scen = self._scenario_for_dry_run or {}
        headers = self._mpv_preview_headers()
        headers = self._decorate_debug_grid_headers(headers)
        ncols = len(headers)
        if ncols == 0:
            self._reset_value_grid()
            self._paint_result_highlights()
            return
        prog_rows: list[list[Any]] = self._mpv_progress_batch_rows()
        pr = len(prog_rows)
        if pr > 0:
            self._mpv_last_valid_table_rows = [list(r) for r in prog_rows]
        keep_last_valid_during_run = bool(
            self._mode == 1
            and pr <= 0
            and (
                bool(getattr(self, "_continuous_busy", False))
                or bool(getattr(self, "_master_step_loop_busy", False))
                or int(getattr(self, "_mpv_join_compute_busy", 0) or 0) > 0
            )
            and bool(self._mpv_last_valid_table_rows)
        )
        if keep_last_valid_during_run:
            prog_rows = [list(r) for r in self._mpv_last_valid_table_rows]
            pr = len(prog_rows)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] value_grid_keep reason=last_valid_table_rows "
                    "mi_idx=%s step_idx=%s buffer_rows=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    pr,
                )
            except Exception:
                pass
        display_floor = self._master_preview_display_rows()
        # table_rows を正とする。データ行数より display_floor で枠だけ膨らませない。
        row_basis = pr if pr > 0 else display_floor
        max_r = max(1, row_basis)
        from svc.svc_data_agg_scenario import KEY_RESULT_COLUMNS, result_column_header_names  # noqa: WPS433

        _extra_ncols = len(result_column_header_names(scen.get(KEY_RESULT_COLUMNS)))
        _n_master_cols = max(0, ncols - _extra_ncols)
        disp_mi = self._mpv_display_mi_idx
        if disp_mi is None or disp_mi < 0 or disp_mi >= _n_master_cols:
            disp_mi = self._mi_idx
        mi = _extra_ncols + max(0, min(int(disp_mi), max(0, _n_master_cols - 1)))
        show_merged_current = bool(getattr(self, "_mpv_show_merged_current", False))
        sync_grid: list[list[Any]] = []
        for r in range(max_r):
            src = prog_rows[r] if r < len(prog_rows) else []
            sync_grid.append(
                [src[c] if c < len(src) else None for c in range(ncols)]
            )
        self._mpv_grid = sync_grid
        self._mpv_join_table_active = True
        self._mpv_join_table_ncols = ncols
        self._value_col_spans = [(0, max(0, ncols - 1))] * max(1, len(self._summary_rows))
        self._value_grid_note_structure(headers)
        self.value_grid.setColumnCount(ncols)
        self.value_grid.setRowCount(max_r)
        self.value_grid.setHorizontalHeaderLabels([str(h) for h in headers])
        for r in range(max_r):
            self.value_grid.setVerticalHeaderItem(r, QTableWidgetItem(str(r + 1)))
        _pe_n = 0

        def _at(rows: list[list[Any]], rr: int, cc: int) -> Any:
            if rr < len(rows):
                rowb = rows[rr]
                if cc < len(rowb):
                    return rowb[cc]
            return None

        def _cell_raw(r: int, c: int) -> Any:
            return _at(prog_rows, r, c)

        for r in range(max_r):
            for c in range(ncols):
                raw = _cell_raw(r, c)
                tx = "" if raw is None else str(raw)
                cell = QTableWidgetItem(tx)
                cell.setToolTip(_normalize_tooltip_text(tx))
                self.value_grid.setItem(r, c, cell)
                _pe_n += 1
                if _pe_n % 48 == 0:
                    self._process_events_light()
        try:
            mismatch = 0
            check_n = min(5, max_r)
            sg = self._mpv_grid or []
            for rr in range(check_n):
                pv = _at(prog_rows, rr, mi)
                gv = _at(sg, rr, mi)
                ps = "" if pv is None else str(pv)
                gs = "" if gv is None else str(gv)
                if ps != gs:
                    mismatch += 1
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_grid_render mi_idx=%s display_mi=%s "
                "show_merged_current_legacy=%s table_rows_only=%s "
                "rows=%s prog_rows=%s sync_grid_rows=%s mismatch_first5=%s",
                self._mi_idx,
                mi,
                show_merged_current,
                True,
                max_r,
                pr,
                len(sg),
                mismatch,
            )
        except Exception:
            pass
        hdr_v = self.value_grid.horizontalHeader()
        hdr_v.setVisible(True)
        hdr_v.setMinimumHeight(32)
        vmins = self._window_int_list("VALUE_GRID_COL_MIN_WIDTHS")
        hdr_v.setMinimumSectionSize(
            self._header_global_floor(vmins, self._table_hdr_min_section())
        )
        self.value_grid.setMinimumWidth(0)
        for c in range(ncols):
            hdr_v.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
        try:
            _diag_logger.info(
                "[DATA_AGG_DEBUG] value_grid_rebuild mode=master preview=mpv_progress "
                "mi_idx=%s ncols=%s nrows=%s prog_rows=%s summary_rows=%s",
                self._mi_idx,
                ncols,
                max_r,
                pr,
                len(self._summary_rows),
            )
        except Exception:
            pass
        if self._mpv_should_defer_column_fit():
            self._mpv_column_fit_pending = True
        else:
            self._fit_value_grid_columns()
        self._paint_result_highlights()
        self._update_values_title_master()

    def _reset_value_grid(self) -> None:
        self.value_grid.clear()
        self.value_grid.setRowCount(0)
        self.value_grid.setColumnCount(0)
        self._mpv_join_table_active = False
        self._mpv_join_table_ncols = 0
        self._value_grid_user_resized = False
        self._value_grid_saved_widths = None
        self._value_grid_structure_key = None

    def _rebuild_value_grid(self) -> None:
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        try:
            self._rebuild_value_grid_impl()
        except DataAggCancelled:
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] value_grid_rebuild_skip reason=cancel "
                    "mi_idx=%s step_idx=%s",
                    self._mi_idx,
                    self._master_step_idx,
                )
            except Exception:
                pass
        except Exception:
            _logger.exception("value grid rebuild failed")
            self._log_append(
                "【内部】結果一覧グリッドの再構築に失敗しました。コンソールに詳細を出力しました。"
            )

    def _refresh_master_value_grid(self, *, finalize: bool) -> None:
        is_cycle_end = (
            self._mode == 1
            and self._mi_idx >= max(0, len(self._master_table_items()) - 1)
            and self._master_step_idx >= len(self._active_slot_indices)
        )
        join_item = self._mpv_current_item_has_join_defs()
        keep_stable_during_continuous = self._mode == 1 and (
            bool(getattr(self, "_continuous_busy", False))
            or bool(getattr(self, "_master_step_loop_busy", False))
            or int(getattr(self, "_mpv_join_compute_busy", 0) or 0) > 0
        )
        if not finalize and keep_stable_during_continuous:
            self._mpv_deferred_value_grid_mi = int(self._mi_idx)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] value_grid_defer mi_idx=%s step_idx=%s "
                    "continuous_busy=%s step_loop_busy=%s join_item=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    bool(getattr(self, "_continuous_busy", False)),
                    bool(getattr(self, "_master_step_loop_busy", False)),
                    bool(join_item),
                )
            except Exception:
                pass
            return
        if finalize and keep_stable_during_continuous and not is_cycle_end:
            self._mpv_deferred_value_grid_mi = int(self._mi_idx)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] value_grid_refresh_finalize mi_idx=%s step_idx=%s",
                    self._mi_idx,
                    self._master_step_idx,
                )
            except Exception:
                pass
            self._rebuild_value_grid()
            return
        if finalize:
            self._flush_deferred_master_value_grid_if_mi(int(self._mi_idx))
            if is_cycle_end and self._scenario_for_dry_run and self._debug_scan_paths:
                if keep_stable_during_continuous:
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_finalize_defer_end_of_continuous "
                            "mi_idx=%s step_idx=%s",
                            self._mi_idx,
                            self._master_step_idx,
                        )
                    except Exception:
                        pass
                    return
                from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                    master_preview_finalize_should_force_recompute,
                )

                cache_ok = self._mpv_finalize_step_cache_acceptable()
                self._mpv_apply_final_result_grid(
                    force_recompute=master_preview_finalize_should_force_recompute(
                        step_cache_hit=cache_ok
                    ),
                )
                return
            # 連続実行中は途中崩れ防止のため、項目完了時でもマージ列表示を有効化しない。
            # 全ステップ終了後に _finish_continuous_run から最終反映を 1 回だけ実行する。
            # 単発実行でも「項目完了ごと」ではなく「全体完了時」のみ最終反映する。
            self._mpv_show_merged_current = (not keep_stable_during_continuous) and is_cycle_end
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_finalize_request mi_idx=%s step_idx=%s "
                    "continuous_busy=%s step_loop_busy=%s cycle_end=%s apply_merged_now=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    bool(getattr(self, "_continuous_busy", False)),
                    bool(getattr(self, "_master_step_loop_busy", False)),
                    is_cycle_end,
                    self._mpv_show_merged_current,
                )
            except Exception:
                pass
            self._mpv_join_table_active = False
            self._mpv_join_table_ncols = 0
            self._rebuild_value_grid()
            return
        self._mpv_display_mi_idx = None
        self._mpv_show_merged_current = False
        self._mpv_join_table_active = False
        self._mpv_join_table_ncols = 0
        self._rebuild_value_grid()

    def _rebuild_value_grid_impl(self) -> None:
        if self._mpv_should_defer_join_value_grid_rebuild():
            self._mpv_deferred_value_grid_mi = int(self._mi_idx)
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] value_grid_defer reason=join_step_pending "
                    "mi_idx=%s step_idx=%s join_busy=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    int(getattr(self, "_mpv_join_compute_busy", 0) or 0),
                )
            except Exception:
                pass
            return
        if self._mode == 1 and self._scenario_for_dry_run and self._debug_scan_paths:
            if not self._summary_rows:
                headers = self._mpv_preview_headers()
                headers = self._decorate_debug_grid_headers(headers)
                if headers:
                    self._mpv_join_table_active = True
                    self._mpv_join_table_ncols = len(headers)
                    self._value_col_spans = [(0, max(0, len(headers) - 1))]
                    self._value_grid_note_structure(headers)
                    self.value_grid.setColumnCount(len(headers))
                    self.value_grid.setRowCount(0)
                    self.value_grid.setHorizontalHeaderLabels([str(h) for h in headers])
                    hdr_v = self.value_grid.horizontalHeader()
                    hdr_v.setVisible(True)
                    hdr_v.setMinimumHeight(32)
                    for c in range(len(headers)):
                        hdr_v.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
                    try:
                        _diag_logger.info(
                            "[DATA_AGG_DEBUG] value_grid_rebuild mode=master preview=header_only "
                            "mi_idx=%s headers=%s summary_rows=%s",
                            self._mi_idx,
                            headers,
                            len(self._summary_rows),
                        )
                    except Exception:
                        pass
                    self._fit_value_grid_columns()
                    self._paint_result_highlights()
                    return
            self._render_mpv_grid()
            return
        self._mpv_join_table_active = False
        self._mpv_join_table_ncols = 0
        expanded: list[tuple[str, list[str], list[str | None] | None]] = []
        self._value_col_spans = []
        for i, colvals in enumerate(self._value_cols):
            base_header = (
                str(self._summary_phase_labels[i])
                if i < len(self._summary_phase_labels)
                else str(i + 1)
            )
            if self._mode == 0:
                base_header = self._value_header_label(base_header)
            row_tips: list[str | None] | None = None
            if i < len(self._value_col_tooltips):
                rt = list(self._value_col_tooltips[i])
                if len(rt) == len(colvals):
                    row_tips = rt
            groups_exp: list[tuple[str, list[str]]] | None = None
            if not (
                self._mode == 0 and self._scenario_source_kind() == "name_extract"
            ):
                groups_exp = expand_hash_bracket_value_groups(
                    [str(x) for x in colvals]
                )
            if groups_exp:
                start = len(expanded)
                for tgt, vals in groups_exp:
                    expanded.append((tgt, vals, None))
                self._value_col_spans.append((start, len(expanded) - 1))
            else:
                pos = len(expanded)
                expanded.append((base_header, list(colvals), row_tips))
                self._value_col_spans.append((pos, pos))
        try:
            if self._mode == 0:
                _diag_logger.info(
                    "[DATA_AGG_DEBUG] value_grid_rebuild sc_idx=%s source_kind=%s "
                    "phase_labels=%s value_col_lens=%s expanded_headers=%s expanded_row_max=%s",
                    self._sc_idx,
                    self._scenario_source_kind(),
                    list(self._summary_phase_labels),
                    [len(c) for c in self._value_cols],
                    [h for h, _, _ in expanded],
                    max((len(col) for _, col, _ in expanded), default=0),
                )
            else:
                _diag_logger.info(
                    "[DATA_AGG_DEBUG] value_grid_rebuild mode=master expanded "
                    "mi_idx=%s phase_labels=%s value_col_lens=%s expanded_headers=%s expanded_row_max=%s",
                    self._mi_idx,
                    list(self._summary_phase_labels),
                    [len(c) for c in self._value_cols],
                    [h for h, _, _ in expanded],
                    max((len(col) for _, col, _ in expanded), default=0),
                )
        except Exception:
            pass

        if self._mode == 0:
            n_before = len(expanded)
            expanded = self._append_scenario_join_columns_if_needed(expanded)
            n_after = len(expanded)
            if n_after > n_before:
                self._value_col_spans.append((n_before, n_after - 1))

        ncols = len(expanded)
        if ncols == 0:
            self._reset_value_grid()
            self._paint_result_highlights()
            return
        max_r = max(len(col) for _, col, _ in expanded) if expanded else 0
        headers = self._decorate_debug_grid_headers([h for h, _, _ in expanded])
        self._value_grid_note_structure(headers)
        self.value_grid.setColumnCount(ncols)
        self.value_grid.setRowCount(max_r)
        self.value_grid.setHorizontalHeaderLabels(headers)
        for r in range(max_r):
            self.value_grid.setVerticalHeaderItem(r, QTableWidgetItem(str(r + 1)))
        _pe_n = 0
        for ci, (_, colvals, coltips) in enumerate(expanded):
            for r, val in enumerate(colvals):
                it = QTableWidgetItem(val)
                tip_txt = str(val)
                if coltips is not None and r < len(coltips):
                    tip = coltips[r]
                    if tip:
                        tip_txt = str(tip)
                it.setToolTip(_normalize_tooltip_text(tip_txt))
                self.value_grid.setItem(r, ci, it)
                _pe_n += 1
                if _pe_n % 48 == 0:
                    self._process_events_light()
        try:
            if self._mode == 1:
                first_row_vals: list[str] = []
                if max_r > 0:
                    for c in range(min(ncols, 6)):
                        it0 = self.value_grid.item(0, c)
                        first_row_vals.append("" if it0 is None else str(it0.text()))
                _diag_logger.info(
                    "[DATA_AGG_DEBUG] value_grid_rendered mode=master mi_idx=%s "
                    "ncols=%s nrows=%s headers_preview=%s row1_preview=%s",
                    self._mi_idx,
                    ncols,
                    max_r,
                    headers[: min(len(headers), 6)],
                    first_row_vals,
                )
        except Exception:
            pass
        hdr_v = self.value_grid.horizontalHeader()
        hdr_v.setVisible(True)
        hdr_v.setMinimumHeight(32)
        vmins = self._window_int_list("VALUE_GRID_COL_MIN_WIDTHS")
        hdr_v.setMinimumSectionSize(
            self._header_global_floor(vmins, self._table_hdr_min_section())
        )
        self.value_grid.setMinimumWidth(0)
        for c in range(ncols):
            hdr_v.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
        self._fit_value_grid_columns()
        self._paint_result_highlights()

    def _apply_scenario_step_result(
        self,
        plab: str,
        vals: list[str],
        colvals: list[str],
        col_tooltips: list[str | None] | None = None,
    ) -> None:
        """シナリオモード: 周回時は同一サマリ／取得列スロットを上書きする。"""
        idx = self._phase_idx
        ct: list[str | None] = list(col_tooltips or [])
        if len(ct) != len(colvals):
            ct = _none_tips(len(colvals))
        row = [plab] + self._summary_vals_for_display(list(vals))
        try:
            if idx < len(self._summary_rows):
                self._summary_rows[idx] = list(vals)
                self._summary_phase_labels[idx] = plab
                while self.summary_table.rowCount() <= idx:
                    self.summary_table.insertRow(self.summary_table.rowCount())
                nc = self.summary_table.columnCount()
                for c, t in enumerate(row):
                    if c >= nc:
                        break
                    disp = str(t) if c == 0 else _summary_metric_cell_display(str(t))
                    it = self.summary_table.item(idx, c)
                    if it is None:
                        self.summary_table.setItem(idx, c, QTableWidgetItem(disp))
                    else:
                        it.setText(disp)
                if idx < len(self._value_cols):
                    self._value_cols[idx] = list(colvals)
                    while len(self._value_col_tooltips) <= idx:
                        self._value_col_tooltips.append([])
                    self._value_col_tooltips[idx] = ct
                else:
                    while len(self._value_cols) < idx:
                        self._value_cols.append([])
                        self._value_col_tooltips.append([])
                    self._value_cols.append(list(colvals))
                    self._value_col_tooltips.append(ct)
            else:
                self._append_summary_row(plab, vals)
                self._value_cols.append(list(colvals))
                self._value_col_tooltips.append(ct)
        except Exception:
            _logger.exception("scenario step summary/value update failed")
            self._log_append(
                "【内部】フェーズ結果の反映に失敗しました。コンソールに詳細を出力しました。"
            )
        else:
            self._fit_summary_table_columns()

    def _append_summary_row(self, phase_label: str, vals: list[str]) -> None:
        try:
            self._summary_rows.append(vals)
            self._summary_phase_labels.append(phase_label)
            r = self.summary_table.rowCount()
            self.summary_table.insertRow(r)
            row = [phase_label] + self._summary_vals_for_display(vals)
            for c, t in enumerate(row):
                disp = t if c == 0 else _summary_metric_cell_display(str(t))
                self.summary_table.setItem(r, c, QTableWidgetItem(disp))
            self._fit_summary_table_columns()
        except Exception:
            _logger.exception("append summary row failed")
            self._log_append(
                "【内部】サマリ行の追加に失敗しました。コンソールに詳細を出力しました。"
            )

    def _on_summary_fold_toggled(self, on: bool) -> None:
        w = getattr(self, "_summary_table_wrap", None)
        if w is not None:
            w.setVisible(on)
        btn = getattr(self, "btn_summary_fold", None)
        if btn is not None:
            btn.setText(
                self._d("BTN_SUMMARY_EXPANDED", "▼ 結果サマリ")
                if on
                else self._d("BTN_SUMMARY_COLLAPSED", "▶ 結果サマリ")
            )
            self._refresh_fold_button_tooltip()

    def _master_active_count_for_item(self, m: dict[str, Any]) -> int:
        idx: list[int] = []
        for si, sc in enumerate(m.get("scenarios") or []):
            slot = sc.get("slot")
            if slot is not None and bool(slot.get("defined", True)):
                idx.append(si)
        cap = self._max_phase_slots()
        return min(len(idx), cap)

    def _master_format_elapsed_sec(self, sec: float | None) -> str:
        if sec is None:
            return ""
        try:
            v = float(sec)
        except (TypeError, ValueError):
            return ""
        if v < 0:
            return ""
        return "%.1f" % v

    def _master_clear_elapsed_timings(self) -> None:
        self._master_step_elapsed_sec.clear()
        self._master_item_elapsed_sec.clear()
        self._master_step_timing_t0 = None
        self._master_continuous_run_t0 = None

    def _master_begin_step_timing(self) -> None:
        self._master_step_timing_t0 = time.perf_counter()

    def _master_finish_step_timing(self, mi_idx: int, step_idx: int) -> None:
        t0 = getattr(self, "_master_step_timing_t0", None)
        self._master_step_timing_t0 = None
        if t0 is None:
            return
        elapsed = max(0.0, time.perf_counter() - float(t0))
        mi = int(mi_idx)
        st = int(step_idx)
        self._master_step_elapsed_sec[(mi, st)] = elapsed
        mit = self._master_table_items()
        n_steps = (
            self._master_active_count_for_item(mit[mi])
            if 0 <= mi < len(mit)
            else 0
        )
        total = sum(
            float(self._master_step_elapsed_sec.get((mi, s), 0.0))
            for s in range(n_steps)
        )
        self._master_item_elapsed_sec[mi] = total
        self._master_refresh_elapsed_ui()

    def _master_begin_continuous_run_timing(self) -> None:
        self._master_continuous_run_t0 = time.perf_counter()

    def _master_continuous_run_elapsed_sec(self) -> float:
        t0 = getattr(self, "_master_continuous_run_t0", None)
        if t0 is None:
            return 0.0
        return max(0.0, time.perf_counter() - float(t0))

    def _master_refresh_elapsed_ui(self) -> None:
        if self._mode != 1:
            return
        try:
            mit = self._master_table_items()
            for ri in range(self.left_table.rowCount()):
                reg = (
                    ri < len(mit)
                    and self._master_active_count_for_item(mit[ri]) > 0
                )
                it = self.left_table.item(ri, 2)
                if it is None:
                    continue
                txt = (
                    self._master_format_elapsed_sec(
                        self._master_item_elapsed_sec.get(int(ri))
                    )
                    if reg
                    else ""
                )
                it.setText(txt)
                it.setTextAlignment(
                    Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
            mi = int(self._mi_idx)
            for r in range(self.left_steps.rowCount()):
                it = self.left_steps.item(r, 2)
                if it is None:
                    continue
                it.setText(
                    self._master_format_elapsed_sec(
                        self._master_step_elapsed_sec.get((mi, int(r)))
                    )
                )
                it.setTextAlignment(
                    Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                )
            self._apply_master_left_registered_row_style()
        except Exception:
            pass

    def _master_cycle_total_steps(self) -> int:
        return sum(
            self._master_active_count_for_item(m) for m in self._master_table_items()
        )

    def _master_continuous_total_ticks(self) -> int:
        """連続実行 1 ティック＝ _execute_single_run_step 1 回に揃える（シナリオなし項目は 1 ティックでスキップ）。"""
        return sum(
            max(self._master_active_count_for_item(m), 1)
            for m in self._master_table_items()
        )

    def _begin_master_run_from_first_item(self) -> None:
        """全項目連続の開始位置を先頭項目・先頭ステップに戻す（周回完了ログは出さない）。"""
        self._master_step_pass_complete = False
        self._clear_master_item_snapshots()
        self._bump_mpv_prefetch_cancel()
        self._mi_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._master_global_row_idx = 0
        self._master_exec_armed = False
        self._mpv_grid = None
        self._mpv_extract_cache.clear()
        self._mpv_colvals_cache.clear()
        self._mpv_progress_rows_cache = None
        self._mpv_last_valid_table_rows = []
        self._mpv_last_stats_files_read = 0
        self._mpv_last_stats_read_rows = 0
        self._mpv_last_stats_scan_cap_hit = False
        self._mpv_join_compute_busy = 0
        self._mpv_progress_rows_step_cache.clear()
        self._mpv_progress_rows_by_mi.clear()
        self._mpv_frozen_snapshots.clear()
        self._mpv_progress_row_peak_by_mi.clear()
        self._mpv_join_search_pool_seed = None
        self._mpv_join_search_pool_seed_paths_count = -1
        self._mpv_join_pool_by_mi.clear()
        self._mpv_row_file_paths_by_mi.clear()
        self._mpv_final_table_rows = None
        self._last_master_completed_mi_idx = None
        self._mpv_display_mi_idx = None
        self._mpv_deferred_value_grid_mi = None
        if self.left_table.rowCount() > 0:
            self.left_table.blockSignals(True)
            try:
                self.left_table.selectRow(0)
            finally:
                self.left_table.blockSignals(False)
        self._update_left_detail()
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._rebuild_value_grid()
        self._paint_left_steps_executed()
        self._paint_result_highlights()

    def _master_global_row_base_for_mi(self, mi: int) -> int:
        items = self._master_table_items()
        s = 0
        for i in range(min(max(0, mi), len(items))):
            s += self._master_active_count_for_item(items[i])
        return s

    def _upsert_summary_row_at(self, idx: int, phase_label: str, vals: list[str]) -> None:
        self._ensure_summary_table_columns()
        while len(self._summary_rows) <= idx:
            self._summary_rows.append(_dash_row(self._n_metrics()))
        while len(self._summary_phase_labels) <= idx:
            self._summary_phase_labels.append("")
        self._summary_rows[idx] = list(vals)
        self._summary_phase_labels[idx] = phase_label
        while self.summary_table.rowCount() <= idx:
            self.summary_table.insertRow(self.summary_table.rowCount())
        row = [phase_label] + self._summary_vals_for_display(vals)
        nc = self.summary_table.columnCount()
        for c, t in enumerate(row):
            if c >= nc:
                break
            disp = str(t) if c == 0 else _summary_metric_cell_display(str(t))
            it = self.summary_table.item(idx, c)
            if it is None:
                self.summary_table.setItem(idx, c, QTableWidgetItem(disp))
            else:
                it.setText(disp)
        self._fit_summary_table_columns()

    def _upsert_value_cols_at(
        self,
        idx: int,
        colvals: list[str],
        tips: list[str | None] | None,
    ) -> None:
        ct: list[str | None] = list(tips or [])
        if len(ct) != len(colvals):
            ct = _none_tips(len(colvals))
        while len(self._value_cols) <= idx:
            self._value_cols.append([])
            self._value_col_tooltips.append([])
        self._value_cols[idx] = list(colvals)
        self._value_col_tooltips[idx] = ct

    def _bump_master_global_row(self) -> None:
        tot = self._master_cycle_total_steps()
        if tot <= 0:
            return
        self._master_global_row_idx = (self._master_global_row_idx + 1) % tot

    def _refresh_all(self) -> None:
        self._reload_left_table()
        if self._mode == 0:
            self._load_scenario_state(self._sc_idx)
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._rebuild_value_grid()
        self._paint_result_highlights()
        self._update_run_buttons_state()
        self._update_clear_buttons()

    def _log_prepend_plain(self, text: str) -> None:
        """タイムスタンプなしでログ先頭に1行挿入（最新が上の並びを維持）。"""
        cur = QTextCursor(self.log.document())
        cur.movePosition(QTextCursor.MoveOperation.Start)
        cur.insertText(str(text).rstrip("\n") + "\n")
        self._log_trim_to_max_lines_if_needed()

    def _log_trim_to_max_lines_if_needed(self) -> None:
        cap = DEBUG_LOG_MAX_LINES
        if cap <= 0:
            return
        text = self.log.toPlainText()
        if not text:
            return
        lines = text.split("\n")
        if len(lines) <= cap:
            return
        self.log.setPlainText("\n".join(lines[:cap]))

    def _log_append(self, msg: str, *, indent_levels: int = 0) -> None:
        pad = " " * (_LOG_INDENT_COLS_PER_LEVEL * max(0, indent_levels))
        line = "[%s] %s%s" % (_ts(), pad, msg)
        cur = QTextCursor(self.log.document())
        cur.movePosition(QTextCursor.MoveOperation.Start)
        cur.insertText(line + "\n")
        self._log_trim_to_max_lines_if_needed()
        if self._mode == 0:
            self._persist_scenario_state()

    def _log_separator_after_results_cleared(self) -> None:
        """結果をクリア後: 先頭に空行1行（次の実行ブロックとの区切り。最新が上なので先頭＝視覚的に上段）。"""
        cur = QTextCursor(self.log.document())
        cur.movePosition(QTextCursor.MoveOperation.Start)
        cur.insertBlock()

    def _log_append_master_item_row(
        self, item_name: str, comment: str, *, item_number: int | None = None
    ) -> None:
        nm = (item_name or "").strip()
        if item_number is not None and int(item_number) >= 1:
            nm = "%d. %s" % (int(item_number), nm)
        body = "%s：%s" % (nm, comment)
        self._log_append(body, indent_levels=1)

    def _log_append_master_scenario_row(
        self, step_n: int, scen_title: str, summary: str
    ) -> None:
        st = (scen_title or "").strip()
        body = "ステップ%d：%s：%s" % (int(step_n), st, summary)
        self._log_append(body, indent_levels=2)

    def _schedule_focus_results_tab(self) -> None:
        """進捗後に結果タブへ切り替え。回数は抑えクリック取りこぼしを防ぐ。"""
        QTimer.singleShot(0, self._focus_results_tab)
        QTimer.singleShot(200, self._focus_results_tab)

    def _master_log_item_and_scenario_counts(self) -> tuple[int, int]:
        items = self._master_table_items()
        n_items = len(items)
        n_scen = sum(len(m.get("scenarios") or []) for m in items)
        return n_items, n_scen

    def _log_master_exec_unit_open(self, run_kind: str) -> None:
        """マスタ実行単位の開始: （種別）マスター項目数・総シナリオ数（先頭空行は結果クリア時のみ）。"""
        if self._mode != 1:
            return
        ni, ns = self._master_log_item_and_scenario_counts()
        self._log_append("（%s）マスター項目数%d、総シナリオ数%d" % (run_kind, ni, ns))

    def _log_master_exec_unit_close(self, run_kind: str, reason: str) -> None:
        """マスタ実行単位の終了。"""
        if self._mode != 1:
            return
        self._log_append("（%s）実行単位終了（%s）" % (run_kind, reason))

    def _show_continuous_run_done_dialog(
        self,
        *,
        mode: int,
        was_full_master: bool,
        steps: int,
        elapsed_sec: float = 0.0,
    ) -> None:
        """連続実行の正常完了時に JSON 文言で終了メッセージを表示する。"""
        title = self._d("DIALOG_RUN_ALL_DONE_TITLE", "データ集約 デバッグ")
        if mode == 0:
            tpl = self._d(
                "MSG_RUN_ALL_SCENARIO_DONE",
                "シナリオの連続実行が完了しました。\n"
                "実行ステップ数: {steps}\n所要時間: {elapsed_sec} 秒",
            )
        elif was_full_master:
            tpl = self._d(
                "MSG_RUN_ALL_MASTER_ITEMS_DONE",
                "全項目の連続実行が完了しました。\n"
                "実行ステップ数: {steps}\n所要時間: {elapsed_sec} 秒",
            )
        else:
            tpl = self._d(
                "MSG_RUN_ALL_MASTER_DONE",
                "項目の連続実行が完了しました。\n"
                "実行ステップ数: {steps}\n所要時間: {elapsed_sec} 秒",
            )
        elapsed_s = self._master_format_elapsed_sec(elapsed_sec) or "0.0"
        try:
            body = tpl.format(steps=int(steps), elapsed_sec=elapsed_s)
        except Exception:
            try:
                body = tpl.format(steps=int(steps))
            except Exception:
                body = tpl
        from ui_qt.ui_common import show_done_notice

        show_done_notice(self, title, _normalize_message_newlines(body))

    def _run_progress_dialog_blocking(self) -> bool:
        pd = getattr(self, "_run_progress_dlg", None)
        if pd is None:
            return False
        try:
            return bool(pd.isVisible())
        except Exception:
            return True

    def _update_run_buttons_state(self) -> None:
        n = len(self._active_slot_indices)
        busy = getattr(self, "_continuous_busy", False)
        prog = self._run_progress_dialog_blocking()
        block = busy or prog
        snap_view = self._mode == 1 and (
            self._master_showing_row_snapshot()
            or getattr(self, "_master_snapshot_browse_after_cancel", False)
        )
        try:
            self.btn_cancel.setEnabled(True)
        except Exception:
            pass
        if self._mode == 0:
            self.btn_run.setEnabled(n > 0 and not block)
            self.btn_run_all.setEnabled(n > 0 and not block)
            self.btn_run_all_master.setEnabled(False)
        else:
            has_m = len(self._master_table_items()) > 0
            rem = max(0, n - self._master_step_idx)
            ticks = self._master_continuous_total_ticks()
            full_ok = bool(getattr(self, "_master_full_continuous_allowed", True))
            pass_done = bool(getattr(self, "_master_step_pass_complete", False))
            run_block = block or snap_view
            self.btn_run.setEnabled(
                has_m and not run_block and not pass_done
            )
            self.btn_run_all.setEnabled(
                has_m and rem > 0 and not run_block and not pass_done
            )
            self.btn_run_all_master.setEnabled(
                has_m
                and ticks > 0
                and not block
                and full_ok
                and not snap_view
                and not pass_done
            )
        if self._mode == 1:
            self._refresh_master_nav_lock_state()

    def _update_clear_buttons(self) -> None:
        has_res = (
            self.summary_table.rowCount() > 0
            or len(self._summary_rows) > 0
            or len(self._value_cols) > 0
        )
        if self._mode == 1:
            # 現在行がシナリオ未登録でサマリ／一覧が空でも、スナップショットや mpv マージバッファが残っていればクリア可能にする
            has_res = has_res or (
                self._mpv_grid is not None
                or bool(self._master_item_snapshots)
                or bool(self._master_item_snapshot_done)
                or bool(self._master_step_snapshots)
            )
        has_log = bool(self.log.toPlainText().strip())
        busy = getattr(self, "_continuous_busy", False)
        prog = self._run_progress_dialog_blocking()
        block = busy or prog
        self.btn_clear_res.setEnabled(has_res and not block)
        self.btn_clear_log.setEnabled(has_log and not block)

    def _select_left_step_row(self) -> None:
        if self.left_steps.rowCount() <= 0:
            return
        idx = self._phase_idx if self._mode == 0 else self._master_step_idx
        if idx >= self.left_steps.rowCount():
            idx = self.left_steps.rowCount() - 1
        self.left_steps.selectRow(idx)

    def _phase_label(self, no: int, name: str) -> str:
        return "%d:%s" % (no, name)

    def _summary_first_col_label(self, no: int, name: str) -> str:
        """結果サマリ先頭列: 左ステップ表に番号があるため名前のみ。"""
        _ = no
        return str(name or "").strip()

    def _selected_item_name(self) -> str:
        item = self._live_item_for_scenario_index(self._sc_idx)
        nm = str(item.get("name") or item.get("id") or "項目").strip()
        return nm or "項目"

    def _value_header_label(self, raw: str) -> str:
        """結果一覧の列見出しを要件に合わせて整形する。"""
        label = str(raw or "").strip()
        if ":" in label:
            label = label.split(":")[-1].strip()
        if self._scenario_source_kind() == "name_extract":
            return label
        if label in (
            "主キー",
            "文字列抽出",
            "抽出",
            "抽出設定",
            "結合パス",
            "検索対象・条件",
        ):
            return self._selected_item_name()
        return label

    def _cancel_scenario_link_prefetch(self, *, join: bool = False) -> None:
        self._scenario_link_prefetch_gen += 1
        try:
            self._scenario_link_prefetch_cancel.set()
        except Exception:
            pass
        th = getattr(self, "_scenario_link_prefetch_thread", None)
        if (
            join
            and th is not None
            and th.is_alive()
            and th is not threading.current_thread()
        ):
            th.join(timeout=180.0)
        if join or th is None or not th.is_alive():
            self._scenario_link_prefetch_thread = None

    def _on_scenario_link_prefetch_finished(self, gen: int, sc_idx: int) -> None:
        if gen != int(getattr(self, "_scenario_link_prefetch_gen", 0)):
            return
        if sc_idx != int(self._sc_idx):
            return
        try:
            _diag_logger.info(
                "[DATA_AGG_DEBUG] scenario_link_prefetch_done sc_idx=%s gen=%s",
                sc_idx,
                gen,
            )
        except Exception:
            pass

    def _schedule_scenario_link_prefetch(self) -> None:
        """主キー表示後、裏で連携・結合を読む。連続実行では呼ばない。"""
        if self._mode != 0 or getattr(self, "_continuous_busy", False):
            return
        item_live = self._live_item_for_scenario_index(self._sc_idx)
        if not item_live.get("sources"):
            return
        self._cancel_scenario_link_prefetch(join=True)
        self._scenario_link_prefetch_cancel = threading.Event()
        gen = int(self._scenario_link_prefetch_gen)
        sc_idx = int(self._sc_idx)
        item = copy.deepcopy(item_live)
        paths = list(self._debug_scan_paths or [])
        cache = self._scenario_bundle_caches.setdefault(sc_idx, {})
        item_id = str(item.get("id") or "item")
        cancel_ev = self._scenario_link_prefetch_cancel
        bridge = self._scenario_link_prefetch_bridge

        def _cancel_check(*, force: bool = False) -> None:
            from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

            if cancel_ev.is_set():
                raise DataAggCancelled()

        def _work() -> None:
            from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

            try:
                from svc.svc_data_agg_debug_run import (  # noqa: WPS433
                    fill_scenario_link_join_after_primary,
                )

                fill_scenario_link_join_after_primary(
                    item,
                    paths,
                    cache,
                    item_id,
                    cancel_check=_cancel_check,
                )
            except DataAggCancelled:
                pass
            except Exception:
                _logger.exception("scenario link prefetch failed")
            try:
                bridge.finished.emit(gen, sc_idx)
            except Exception:
                pass

        th = threading.Thread(
            target=_work,
            daemon=True,
            name="scenario_link_prefetch",
        )
        self._scenario_link_prefetch_thread = th
        th.start()

    def _clear_current_scenario_results_only(self) -> None:
        self._cancel_scenario_link_prefetch(join=False)
        self._scenario_bundle_caches.pop(self._sc_idx, None)
        self._phase_idx = 0
        self._summary_rows.clear()
        self._value_cols.clear()
        self._value_col_tooltips.clear()
        self._value_col_spans.clear()
        self._summary_phase_labels.clear()
        self.summary_table.setRowCount(0)
        self._reset_value_grid()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._paint_result_highlights()
        self._update_run_buttons_state()
        self._update_clear_buttons()
        self._persist_scenario_state()

    def _wrap_scenario_cycle_without_clear(self) -> None:
        """最終フェーズ後は先頭へ戻すのみ（結果は保持し次周で上書き）。"""
        self._log_append(
            "（全ステップ完了。先頭フェーズに戻ります。"
            " サマリ・結果一覧は次周で同一行を上書きします。ログは継続）"
        )
        self._phase_idx = 0

    def _focus_results_tab(self) -> None:
        """ステップ実行・スキップ後、結果タブが手前に無ければ結果へ切り替える。"""
        tabs = getattr(self, "tabs", None)
        res = getattr(self, "tab_res", None)
        if tabs is None or res is None:
            return
        idx = tabs.indexOf(res)
        if idx < 0:
            return
        tabs.setCurrentIndex(idx)
        try:
            QApplication.processEvents()
        except Exception:
            pass

    def _master_finish_step_pass_idle(self) -> None:
        """ステップ実行で全項目を終えたとき: 先頭へ周回せず最終項目で待機する（重い結果一覧の再構築は行わない）。"""
        if getattr(self, "_master_step_pass_complete", False):
            return
        self._master_step_pass_complete = True
        nmit = len(self._master_table_items())
        self._log_append(
            self._d(
                "MSG_MASTER_STEP_PASS_DONE",
                "（全マスタ項目のステップ実行が完了しました。スナップショット閲覧の準備ができた場合は左上コーナーが薄青になります。"
                " 再実行するには結果をクリアするか、全項目の連続実行を完了してください。）",
            )
        )
        if nmit <= 0:
            return
        self._mi_idx = nmit - 1
        self._master_step_idx = 0
        self._master_exec_armed = False
        self.left_table.blockSignals(True)
        try:
            self.left_table.selectRow(self._mi_idx)
        finally:
            self.left_table.blockSignals(False)
        self._update_left_detail()
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._paint_result_highlights()
        self._refresh_master_snapshot_chrome()
        if not getattr(self, "_continuous_busy", False):
            self._mpv_apply_final_result_grid()

    def _execute_single_run_step(self) -> tuple[bool, bool]:
        """ステップ実行 1 回分の本処理。

        戻り値: (実行したか, 結果タブへ切り替えるか)。
        マスタモードでは mi_idx が 0→1 に進んだタイミング（先頭項目を終えて次項目へ）のみ後者が True。
        """
        if self._mode == 0:
            if self._phase_idx >= len(self._active_slot_indices):
                return False, False
            gi = self._active_slot_indices[self._phase_idx]
            slot = self._scenario_slots()[gi]
            assert slot is not None
            plab = self._cond_keys()[gi]
            self._show_scenario_step_progress_start(
                gi,
                plab,
                detail="ステップ %s を実行中" % str(plab or "").strip(),
            )
            item_live = self._live_item_for_scenario_index(self._sc_idx)
            if item_live.get("sources"):
                from svc.svc_data_agg_debug_run import (
                    format_synthetic_events_for_log,
                    scenario_debug_phase_result,
                )

                cache = self._scenario_bundle_caches.setdefault(self._sc_idx, {})
                n_scan = len(self._debug_scan_paths or [])
                prog_hook: Callable[[int, int], None] | None = None
                if self._scenario_wants_file_progress(gi, n_scan):
                    prog_hook = self._scenario_make_file_progress_hook(
                        self._scenario_file_progress_phase_message(gi)
                    )
                continuous = bool(getattr(self, "_continuous_busy", False))
                if gi >= 3:
                    self._cancel_scenario_link_prefetch(join=True)
                try:
                    vals, colvals, events, col_tips = scenario_debug_phase_result(
                        item_live,
                        self._debug_scan_paths,
                        gi,
                        self._max_value_rows(),
                        cache,
                        scan_root=self._scan_root,
                        name_extract_debug_labels=self._cfg.get("NAME_EXTRACT_DEBUG"),
                        progress_hook=prog_hook,
                        phase2_primary_only=not continuous,
                    )
                finally:
                    self._close_run_progress()
                if gi == 2 and not continuous:
                    self._schedule_scenario_link_prefetch()
                colvals, col_tips = self._icap_with_tips(colvals, col_tips)
                sid = str(item_live.get("id") or "")
                for ln in format_synthetic_events_for_log(events, sid):
                    self._log_prepend_plain(ln)
            else:
                vals = list(slot["summary_vals"])
                colvals = self._icap(list(slot["values_column"]))
                col_tips = _none_tips(len(colvals))
                self._show_run_progress(
                    "結果を反映中",
                    1,
                    1,
                    window_title=self._scenario_progress_window_title(),
                    detail="ステップ %s の表示を更新中" % str(plab or "").strip(),
                )
                self._close_run_progress()
            self._apply_scenario_step_result(plab, vals, colvals, col_tips)
            try:
                if self._mode == 0:
                    _diag_logger.info(
                        "[DATA_AGG_DEBUG] scenario_step sc_idx=%s phase_idx=%s gi=%s "
                        "plab=%s source_kind=%s summary=%s col_count=%s col_preview=%s",
                        self._sc_idx,
                        self._phase_idx,
                        gi,
                        plab,
                        self._scenario_source_kind(),
                        vals,
                        len(colvals),
                        colvals[:5],
                    )
            except Exception:
                pass
            self._log_append("ステップ %s　%s" % (plab, self._slot_summary_row(vals)))
            self._phase_idx += 1
            if self._phase_idx >= len(self._active_slot_indices):
                self._wrap_scenario_cycle_without_clear()
            self._rebuild_left_steps()
            self._rebuild_value_grid()
            self._paint_left_steps_executed()
            self._paint_result_highlights()
            if self.left_steps.rowCount() > 0:
                self._select_left_step_row()
            self._persist_scenario_state()
            return True, True

        if self._master_step_pass_complete and not getattr(
            self, "_continuous_busy", False
        ):
            self._log_append(
                self._d(
                    "MSG_MASTER_STEP_PASS_BLOCK",
                    "（ステップ実行は全項目まで完了済みです。再開するには結果をクリアするか、全項目の連続実行を完了してください。）",
                )
            )
            return False, False

        self._ensure_master_run_cancel()
        self._ensure_master_cancel_pump_timer()
        focus_results_tab = False
        while True:
            nmit = len(self._master_table_items())
            if nmit <= 0:
                return False, False
            if self._mi_idx >= nmit:
                if not self._master_step_pass_complete:
                    self._master_finish_step_pass_idle()
                else:
                    self._mi_idx = max(0, nmit - 1)
                nmit = len(self._master_table_items())
                if nmit <= 0:
                    return False, False
                if self._master_step_pass_complete:
                    return False, False
            if self._active_slot_indices:
                n_act = len(self._active_slot_indices)
                lr = self.left_steps.currentRow()
                if (
                    lr >= 0
                    and lr < n_act
                    and self._master_step_idx < n_act
                ):
                    self._master_step_idx = lr
            if not self._active_slot_indices:
                _mt_skip = str(self._current_master().get("title") or "").strip() or "項目"
                self._log_append_master_item_row(
                    _mt_skip,
                    "シナリオなしのためスキップ",
                    item_number=self._mi_idx + 1,
                )
                self._bump_mpv_prefetch_cancel()
                done_mi = self._mi_idx
                self._capture_master_leave_item(done_mi, empty=True)
                self._mi_idx += 1
                self._master_step_idx = 0
                self._master_exec_armed = False
                if self._mi_idx < len(self._master_table_items()):
                    self.left_table.blockSignals(True)
                    try:
                        self.left_table.selectRow(self._mi_idx)
                    finally:
                        self.left_table.blockSignals(False)
                    self._update_left_detail()
                    self._reload_conditions()
                    self._rebuild_active_slots()
                    self._rebuild_left_steps()
                    self._paint_left_steps_executed()
                    self._paint_result_highlights()
                    self._apply_master_left_registered_row_style()
                else:
                    self._master_finish_step_pass_idle()
                # 実行可能シナリオが無い項目の「次」が実行ありのとき、入場直後の再構築は
                # その項目の全ステップ完了時（_flush_deferred_master_value_grid_if_mi）へ遅延する。
                if self._active_slot_indices:
                    self._mpv_deferred_value_grid_mi = int(self._mi_idx)
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] value_grid_defer_post_empty_skip "
                            "landing_mi=%s join_item=%s",
                            self._mi_idx,
                            self._mpv_current_item_has_join_defs(int(self._mi_idx)),
                        )
                    except Exception:
                        pass
                else:
                    self._mpv_deferred_value_grid_mi = None
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] value_grid_keep reason=no_active_slots mi_idx=%s",
                            self._mi_idx,
                        )
                    except Exception:
                        pass
                return True, done_mi == 0
            if self._master_step_idx >= len(self._active_slot_indices):
                self._master_session_start_step += self._last_master_active_count
                self._bump_mpv_prefetch_cancel()
                done_mi = self._mi_idx
                self._flush_deferred_master_value_grid_if_mi(done_mi)
                self._capture_master_leave_item(done_mi, empty=False)
                self._mi_idx += 1
                self._master_step_idx = 0
                self._master_exec_armed = False
                if self._mi_idx >= len(self._master_table_items()):
                    self._master_finish_step_pass_idle()
                    return True, done_mi == 0
                self.left_table.blockSignals(True)
                try:
                    self.left_table.selectRow(self._mi_idx)
                finally:
                    self.left_table.blockSignals(False)
                self._update_left_detail()
                self._reload_conditions()
                self._rebuild_active_slots()
                self._rebuild_left_steps()
                self._paint_left_steps_executed()
                self._paint_result_highlights()
                self._apply_master_left_registered_row_style()
                if done_mi == 0:
                    focus_results_tab = True
                continue
            break
        self._master_step_exec_depth = int(getattr(self, "_master_step_exec_depth", 0) or 0) + 1
        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        m = self._current_master()
        _mt_item = str(m.get("title") or m.get("name") or m.get("id") or "").strip() or "項目"
        if self._master_step_idx == 0:
            if self._mpv_current_item_has_join_defs():
                self._bump_mpv_prefetch_cancel()
                sk0 = self._mpv_progress_step_cache_key(1)
                cached0 = self._mpv_progress_rows_step_cache.get(sk0)
                if cached0 is not None and self._mpv_step_cached_rows_acceptable(
                    cached0,
                    mi_idx=int(self._mi_idx),
                    n_pick=1,
                ):
                    self._mpv_progress_rows_cache = None
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_join_item_enter mi_idx=%s "
                            "progress_cache_cleared=0 step_cache_rows=%s",
                            self._mi_idx,
                            len(cached0),
                        )
                    except Exception:
                        pass
                else:
                    self._mpv_progress_rows_cache = None
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] mpv_join_item_enter mi_idx=%s "
                            "progress_cache_cleared=1 last_valid_preserved=%s seed_preserved=%s",
                            self._mi_idx,
                            len(self._mpv_last_valid_table_rows),
                            len(self._mpv_join_pool_by_mi.get(int(self._mi_idx) - 1, [])),
                        )
                    except Exception:
                        pass
            self._log_append_master_item_row(
                _mt_item, "実行開始", item_number=self._mi_idx + 1
            )
        si = self._active_slot_indices[self._master_step_idx]
        sc = m["scenarios"][si]
        slot = sc["slot"]
        assert slot is not None
        self._master_begin_step_timing()
        sub_total = len(_MASTER_DEBUG_PROGRESS_PHASES)
        prog_wt = self._scenario_progress_window_title()
        # compute_batch フックは _show_run_progress より前に _mpv_progress_batch_rows を
        # 呼ぶため、この時点でタイトルを同期しないと直前項目（例: 品名_PSU）のタイトルのままになる。
        self._master_progress_window_title = prog_wt
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] master_step_prepare_start mi_idx=%s step_idx=%s si=%s title=%s",
                self._mi_idx,
                self._master_step_idx,
                si,
                str(sc.get("title") or ""),
            )
        except Exception:
            pass
        vals = list(slot["summary_vals"])
        gno = self._master_session_start_step + self._master_step_idx + 1
        sc_title = str(sc["title"] or m["title"])
        n_scan = len(self._debug_scan_paths or [])
        if self._master_step_idx == 0 and self._mpv_current_item_has_join_defs():
            from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                master_preview_join_step0_initial_progress,
            )

            _phase0, _done0 = master_preview_join_step0_initial_progress()
            self._show_run_progress(
                _phase0,
                _done0,
                sub_total,
                window_title=prog_wt,
                detail="シナリオ「%s」検出 %s 件 — 表示上限 %s 行"
                % (sc_title, n_scan, self._master_preview_display_rows()),
            )
        else:
            self._show_run_progress(
                _MASTER_DEBUG_PROGRESS_PHASES[0],
                1,
                sub_total,
                window_title=prog_wt,
                detail="シナリオ「%s」検出 %s 件 — 表示上限 %s 行"
                % (sc_title, n_scan, self._master_preview_display_rows()),
            )
        self._process_events_light()
        plab = self._phase_label(gno, sc_title)
        plab_summary = self._summary_first_col_label(gno, sc_title)
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        try:
            colvals = self._mpv_resolve_master_step_colvals(si)
        except DataAggCancelled:
            self._master_note_cancel_requested()
            return self._master_abort_step_after_cancel()
        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        if self._scenario_for_dry_run and self._debug_scan_paths:
            self._mpv_colvals_cache[(int(self._mi_idx), int(si))] = list(colvals)
        gr = self._master_global_row_idx
        self._show_run_progress(
            _MASTER_DEBUG_PROGRESS_PHASES[1],
            2,
            sub_total,
            window_title=prog_wt,
            detail=plab_summary,
        )
        self._process_events_light()
        self._upsert_summary_row_at(gr, plab_summary, vals)
        self._show_run_progress(
            _MASTER_DEBUG_PROGRESS_PHASES[2],
            3,
            sub_total,
            window_title=prog_wt,
            detail="列「%s」%s 件"
            % (str(m.get("title") or m.get("name") or ""), len(colvals)),
        )
        self._process_events_light()
        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        self._upsert_value_cols_at(gr, colvals, _none_tips(len(colvals)))
        if (
            self._scenario_for_dry_run
            and self._debug_scan_paths
            and colvals
            and not self._mpv_current_item_has_join_defs()
        ):
            self._merge_mpv_column(self._mi_idx, colvals)
        self._bump_master_global_row()
        try:
            _diag_logger.info(
                "[DATA_AGG_DEBUG] master_step mi_idx=%s master_step_idx=%s si=%s sc_title=%s "
                "plab=%s summary=%s col_count=%s col_preview=%s",
                self._mi_idx,
                self._master_step_idx,
                si,
                str(sc.get("title") or ""),
                plab,
                vals,
                len(colvals),
                colvals[:5],
            )
        except Exception:
            pass
        _log_step_n = int(self._master_step_idx) + 1
        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        self._capture_master_step_snapshot(
            int(self._mi_idx),
            max(0, int(_log_step_n) - 1),
            colvals=list(colvals) if colvals else None,
        )
        self._master_finish_step_timing(int(self._mi_idx), int(self._master_step_idx))
        self._master_step_idx += 1
        self._last_master_active_count = len(self._active_slot_indices)
        will_finish_master_item = self._master_step_idx >= len(self._active_slot_indices)
        should_finalize_now = will_finish_master_item
        if will_finish_master_item and self._mpv_is_single_slot_active():
            from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                master_preview_item_complete_prefetch_wait_ms,
                master_preview_item_complete_should_capture_frozen,
                master_preview_item_complete_should_ensure_n_pick1,
                master_preview_should_warmup_single_slot,
            )

            if master_preview_should_warmup_single_slot(
                has_join_defs=self._mpv_current_item_has_join_defs()
            ) and self._mpv_is_single_slot_prefetch_pending():
                self._mpv_wait_single_slot_n_pick1_cache(
                    max_wait_ms=master_preview_item_complete_prefetch_wait_ms(
                        prefetch_pending=True,
                        cache_hit=False,
                    )
                )
            _cached_done = self._mpv_rows_from_step_cache_n_pick(1)
            _cache_hit_done = bool(_cached_done)
            _need_ensure = master_preview_item_complete_should_ensure_n_pick1(
                single_slot=True,
                cache_hit=_cache_hit_done,
            )
            _need_frozen = master_preview_item_complete_should_capture_frozen(
                frozen_enabled=self._mpv_frozen_columns_enabled(),
                snapshot_exists=int(self._mi_idx) in self._mpv_frozen_snapshots,
            )
            if self._mpv_current_item_has_join_defs():
                # join_sync_compute で step キャッシュ済み。完了時の再 compute は列を壊す。
                _need_ensure = False
                _need_frozen = False
            if _need_ensure or _need_frozen:
                _fcap_done: dict[str, Any] | None = (
                    {} if self._mpv_frozen_columns_enabled() else None
                )
                _prefetch_pending = (
                    _need_ensure and self._mpv_is_single_slot_prefetch_pending()
                )
                self._mpv_ensure_single_slot_n_pick1_cached(
                    progress_hook=self._mpv_master_dbg_progress_hook_or_none(),
                    frozen_capture_out=_fcap_done,
                    wait_async_ms=master_preview_item_complete_prefetch_wait_ms(
                        prefetch_pending=_prefetch_pending,
                        cache_hit=_cache_hit_done,
                    ),
                )
        elif will_finish_master_item:
            from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                master_preview_item_complete_should_capture_frozen,
            )

            n_act_cap = len(self._active_slot_indices or [])
            cached_cap = (
                self._mpv_rows_from_step_cache_n_pick(n_act_cap)
                if n_act_cap > 0
                else None
            )
            if (
                master_preview_item_complete_should_capture_frozen(
                    frozen_enabled=self._mpv_frozen_columns_enabled(),
                    snapshot_exists=int(self._mi_idx) in self._mpv_frozen_snapshots,
                )
                and not cached_cap
                and n_act_cap > 0
            ):
                self._mpv_ensure_step_n_pick_cached(
                    n_pick=n_act_cap,
                    progress_hook=self._mpv_master_dbg_progress_hook_or_none(),
                    frozen_capture_out={},
                    wait_async_ms=0,
                    probe_caller="mpv_multislot_frozen_capture",
                )
        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        if will_finish_master_item:
            n_act_done = len(self._active_slot_indices or [])
            if self._mpv_current_item_has_join_defs() and n_act_done > 0:
                cached_join = self._mpv_rows_from_step_cache_n_pick(n_act_done)
                need_recompute = not cached_join or not self._mpv_step_cached_rows_acceptable(
                    cached_join,
                    mi_idx=int(self._mi_idx),
                    n_pick=n_act_done,
                )
                if need_recompute:
                    self._mpv_ensure_step_n_pick_cached(
                        n_pick=n_act_done,
                        progress_hook=self._mpv_master_dbg_progress_hook_or_none(),
                        probe_caller="mpv_join_item_complete",
                    )
            n_pick_sync = 1 if self._mpv_is_single_slot_active() else max(1, n_act_done)
            self._mpv_sync_progress_cache_from_step_n_pick(n_pick_sync)
            if not self._mpv_current_item_has_join_defs():
                self._mpv_join_search_pool_seed = None
                self._mpv_join_search_pool_seed_paths_count = -1
            done_mi_key = int(self._mi_idx)
            if (
                self._mpv_current_item_has_join_defs()
                and not self._mpv_is_single_slot_active()
            ):
                for sk in list(self._mpv_progress_rows_step_cache.keys()):
                    if isinstance(sk, tuple) and sk and int(sk[0]) == done_mi_key:
                        self._mpv_progress_rows_step_cache.pop(sk, None)
        # 条件ステップ表の selectRow は _rebuild_left_steps 末尾で行われる。
        # グリッド再構成（本番同等プレビュー）より先に呼ぶと、重い処理の前に次行へフォーカスが移ってしまうため、順序をグリッド→左ステップにする。
        self._master_progress_window_title = prog_wt
        if self._mpv_current_item_has_join_defs():
            n_sync = 1 if self._mpv_is_single_slot_active() else max(
                1, len(self._active_slot_indices or [])
            )
            self._mpv_sync_progress_cache_from_step_n_pick(n_sync)
        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        try:
            self._refresh_master_value_grid(finalize=should_finalize_now)
        except DataAggCancelled:
            self._master_note_cancel_requested()
            return self._master_abort_step_after_cancel()
        aborted = self._master_return_if_step_cancelled()
        if aborted is not None:
            return aborted
        self._rebuild_left_steps()
        self._show_run_progress(
            _MASTER_DEBUG_PROGRESS_PHASE_DONE,
            sub_total,
            sub_total,
            window_title=prog_wt,
        )
        self._process_events_light()
        self._paint_left_steps_executed()
        self._paint_result_highlights()
        _sc_title = str(sc.get("title") or "").strip()
        self._log_append_master_scenario_row(
            _log_step_n, _sc_title, self._slot_summary_row(vals)
        )
        if self.left_steps.rowCount() > 0:
            self._select_left_step_row()

        if self._master_step_idx >= len(self._active_slot_indices):
            self._master_session_start_step += self._last_master_active_count
            self._bump_mpv_prefetch_cancel()
            done_mi = self._mi_idx
            self._flush_deferred_master_value_grid_if_mi(done_mi)
            self._capture_master_leave_item(done_mi, empty=False)
            self._mi_idx += 1
            self._master_exec_armed = False
            if done_mi == 0:
                focus_results_tab = True
            if self._mi_idx >= len(self._master_table_items()):
                self._master_finish_step_pass_idle()
            elif self._mi_idx < len(self._master_table_items()):
                self.left_table.blockSignals(True)
                try:
                    self.left_table.selectRow(self._mi_idx)
                finally:
                    self.left_table.blockSignals(False)
                self._master_step_idx = 0
                self._master_exec_armed = False
                self._update_left_detail()
                self._reload_conditions()
                self._rebuild_active_slots()
                self._rebuild_left_steps()
                self._paint_left_steps_executed()
                self._paint_result_highlights()
                self._apply_master_left_registered_row_style()
                if self._active_slot_indices:
                    self._mpv_warmup_single_slot_progress_cache(int(self._mi_idx))
                    self._rebuild_value_grid()
                else:
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] value_grid_keep reason=no_active_slots_after_item_done mi_idx=%s",
                            self._mi_idx,
                        )
                    except Exception:
                        pass
        if self._master_cancel_pending():
            return self._master_abort_step_after_cancel()
        self._close_run_progress()
        self._master_step_exec_depth_leave()
        return True, focus_results_tab

    def _master_step_exec_depth_leave(self) -> None:
        self._master_step_exec_depth = max(
            0, int(getattr(self, "_master_step_exec_depth", 0) or 0) - 1
        )

    def _finish_continuous_run(self) -> None:
        if not self._continuous_busy:
            return
        # マスタ連続実行完了時: 左が「シナリオなし」項目にいても、
        # 直前までの結果一覧を最終反映する（マージ列表示を 1 回）。
        if self._mode == 1:
            self._continuous_busy = False
            self._master_step_loop_busy = False
            fb_mi = getattr(self, "_last_master_completed_mi_idx", None)
            disp_mi = self._mpv_finalize_target_mi()
            self._mpv_display_mi_idx = (
                disp_mi if disp_mi is not None else (fb_mi if fb_mi is not None else self._mi_idx)
            )
            try:
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_finalize_apply_end_of_continuous "
                    "mi_idx=%s step_idx=%s active_slots=%s last_completed_mi=%s",
                    self._mi_idx,
                    self._master_step_idx,
                    len(self._active_slot_indices or []),
                    getattr(self, "_last_master_completed_mi_idx", None),
                )
            except Exception:
                pass
            from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                master_preview_finalize_should_force_recompute,
            )

            cache_ok = self._mpv_finalize_step_cache_acceptable()
            self._mpv_apply_final_result_grid(
                force_recompute=master_preview_finalize_should_force_recompute(
                    step_cache_hit=cache_ok
                ),
            )
            self._paint_result_highlights()
        self._continuous_busy = False
        self._continuous_steps_left = 0
        self._stop_master_cancel_pump_timer()
        self._clear_master_run_cancel()
        if self._mode == 1:
            if getattr(self, "_continuous_was_full_master", False):
                self._master_full_continuous_allowed = True
            self._continuous_was_full_master = False
            # _master_step_pass_complete はクリアしない。クリアすると待機直後の再描画が
            # n_pick=0 ステップキャッシュに戻り結合列が空に見える（結果タブと table_rows の整合が崩れる）。
        self._update_run_buttons_state()
        self._update_clear_buttons()

    def _run_continuous_next(self) -> None:
        if not self._continuous_busy:
            return
        if getattr(self, "_master_continuous_cancel_requested", False):
            self._master_continuous_cancel_requested = False
            self._log_append(
                self._d("MSG_MASTER_RUN_CANCEL", "（連続実行を中止しました）")
            )
            self._show_master_run_cancel_notice(continuous=True)
            self._finish_continuous_run()
            return
        if self._continuous_steps_left <= 0:
            self._finish_continuous_run()
            return
        self._ensure_master_run_cancel()
        self._ensure_master_cancel_pump_timer()
        ok = False
        focus_tab = False
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        try:
            ok, focus_tab = self._execute_single_run_step()
        except DataAggCancelled:
            self._master_note_cancel_requested()
            self._master_abort_step_after_cancel()
            return
        if ok:
            if focus_tab:
                self._focus_results_tab()
                self._schedule_focus_results_tab()
            self._update_run_buttons_state()
            self._update_clear_buttons()
        if not self._continuous_busy:
            return
        if not ok:
            if self._mode == 1:
                _rk = (
                    "全項目連続実行"
                    if getattr(self, "_continuous_was_full_master", False)
                    else "連続実行"
                )
                self._log_master_exec_unit_close(_rk, "中断")
            else:
                self._log_append(
                    self._d(
                        "MSG_RUN_ALL_ABORT",
                        "連続実行を中断しました（実行できるステップがありません）。",
                    )
                )
            self._finish_continuous_run()
            return
        self._continuous_steps_left -= 1
        if self._continuous_steps_left <= 0:
            done_mode = self._mode
            done_was_full = getattr(self, "_continuous_was_full_master", False)
            done_steps = int(getattr(self, "_continuous_initial_steps", 0) or 0)
            done_elapsed = self._master_continuous_run_elapsed_sec()
            if self._mode == 1:
                _rk = "全項目連続実行" if done_was_full else "連続実行"
                self._log_master_exec_unit_close(_rk, "完了")
            else:
                self._log_append(
                    self._d("MSG_RUN_ALL_DONE", "連続実行が完了しました。")
                )
            self._finish_continuous_run()
            self._show_continuous_run_done_dialog(
                mode=done_mode,
                was_full_master=done_was_full,
                steps=done_steps,
                elapsed_sec=done_elapsed,
            )
            self._master_continuous_run_t0 = None
        else:
            QTimer.singleShot(0, self._run_continuous_next)

    def _on_run_all_continuous(self) -> None:
        if self._continuous_busy:
            return
        if self._mode == 1:
            self._continuous_was_full_master = False
            self._master_full_continuous_allowed = False
            self._master_step_pass_complete = False
            self._master_snapshot_browse_after_cancel = False
            self._reset_master_cancel_state()
        if self._mode == 0:
            n = len(self._active_slot_indices)
        else:
            n = max(0, len(self._active_slot_indices) - self._master_step_idx)
        if n <= 0:
            return
        self._continuous_busy = True
        self._continuous_steps_left = n
        self._continuous_initial_steps = n
        self._master_begin_continuous_run_timing()
        self._reset_master_cancel_state()
        self._ensure_master_run_cancel()
        self._ensure_master_cancel_pump_timer()
        self._mpv_show_merged_current = False
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] run_all_start mode=%s steps=%s mi_idx=%s step_idx=%s",
                self._mode,
                n,
                self._mi_idx,
                self._master_step_idx,
            )
        except Exception:
            pass
        if self._mode == 1:
            self._log_master_exec_unit_open("連続実行")
        else:
            self._log_append(
                self._d("MSG_RUN_ALL_START", "連続実行を開始しました（%d ステップ）。") % n
            )
        self._update_run_buttons_state()
        QTimer.singleShot(0, self._run_continuous_next)

    def _on_run_all_master_items_continuous(self) -> None:
        if self._continuous_busy:
            return
        if self._mode != 1:
            return
        total = self._master_continuous_total_ticks()
        if total <= 0:
            return
        self._continuous_was_full_master = True
        self._master_full_continuous_allowed = False
        self._begin_master_run_from_first_item()
        self._continuous_busy = True
        self._continuous_steps_left = total
        self._continuous_initial_steps = total
        self._master_begin_continuous_run_timing()
        self._reset_master_cancel_state()
        self._ensure_master_run_cancel()
        self._ensure_master_cancel_pump_timer()
        self._mpv_show_merged_current = False
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] run_all_master_items_start ticks=%s",
                total,
            )
        except Exception:
            pass
        self._log_master_exec_unit_open("全項目連続実行")
        self._update_run_buttons_state()
        QTimer.singleShot(0, self._run_continuous_next)

    def _on_run(self) -> None:
        if self._mode == 1 and getattr(self, "_master_step_pass_complete", False):
            self._log_append(
                self._d(
                    "MSG_MASTER_STEP_PASS_BLOCK",
                    "（ステップ実行は全項目まで完了済みです。再開するには結果をクリアするか、全項目の連続実行を完了してください。）",
                )
            )
            return
        if self._mode == 1:
            self._log_master_exec_unit_open("ステップ実行")
            self._master_snapshot_browse_after_cancel = False
            self._reset_master_cancel_state()
            self._ensure_master_run_cancel()
            self._ensure_master_cancel_pump_timer()
        ok = False
        focus_tab = False
        use_wait_cursor = self._mode != 1
        if use_wait_cursor:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433

        try:
            try:
                ok, focus_tab = self._execute_single_run_step()
                if ok:
                    if self._mode == 1:
                        self._master_full_continuous_allowed = False
                    if focus_tab:
                        self._focus_results_tab()
                        self._schedule_focus_results_tab()
                    self._update_run_buttons_state()
                    self._update_clear_buttons()
            except DataAggCancelled:
                self._master_note_cancel_requested()
                self._master_abort_step_after_cancel()
                ok = False
        finally:
            if self._mode == 1:
                self._log_master_exec_unit_close(
                    "ステップ実行", "完了" if ok else "中断"
                )
            if self._mode == 1:
                QTimer.singleShot(350, self._close_run_progress)
                if not getattr(self, "_continuous_busy", False):
                    self._stop_master_cancel_pump_timer()
            if use_wait_cursor:
                QApplication.restoreOverrideCursor()

    def _clear_master_debug_results(self) -> None:
        self._master_full_continuous_allowed = True
        self._master_step_pass_complete = False
        self._master_clear_elapsed_timings()
        self._clear_master_item_snapshots()
        self._bump_mpv_prefetch_cancel()
        self._master_sparse_notice_shown = False
        self._phase_idx = 0
        self._mi_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._master_exec_armed = False
        self._master_global_row_idx = 0
        self._mpv_grid = None
        self._mpv_extract_cache.clear()
        self._mpv_colvals_cache.clear()
        self._mpv_progress_rows_cache = None
        self._mpv_last_valid_table_rows = []
        self._mpv_last_stats_files_read = 0
        self._mpv_last_stats_read_rows = 0
        self._mpv_last_stats_scan_cap_hit = False
        self._mpv_join_compute_busy = 0
        self._mpv_progress_rows_step_cache.clear()
        self._mpv_progress_rows_by_mi.clear()
        self._mpv_frozen_snapshots.clear()
        self._mpv_progress_row_peak_by_mi.clear()
        self._mpv_join_search_pool_seed = None
        self._mpv_join_search_pool_seed_paths_count = -1
        self._mpv_join_pool_by_mi.clear()
        self._mpv_row_file_paths_by_mi.clear()
        self._mpv_final_table_rows = None
        self._mpv_column_fit_pending = False
        self._mpv_final_grid_applied = False
        self._last_master_completed_mi_idx = None
        self._mpv_display_mi_idx = None
        self._summary_rows.clear()
        self._value_cols.clear()
        self._value_col_tooltips.clear()
        self._value_col_spans.clear()
        self._summary_phase_labels.clear()
        self.summary_table.setRowCount(0)
        self.left_table.blockSignals(True)
        try:
            if self.left_table.rowCount() > 0:
                self.left_table.selectRow(0)
        finally:
            self.left_table.blockSignals(False)
        self._reload_left_table()
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_value_grid()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._paint_result_highlights()

    def _on_clear_results(self) -> None:
        if self._mode == 0:
            self._clear_current_scenario_results_only()
            self._log_separator_after_results_cleared()
            self._persist_scenario_state()
            return
        self._clear_master_debug_results()
        self._log_separator_after_results_cleared()
        self._update_run_buttons_state()
        self._update_clear_buttons()

    def _on_clear_log_only(self) -> None:
        self.log.clear()
        if self._mode == 0:
            st = self._ensure_scenario_state(self._sc_idx)
            st["log"] = ""
        self._update_clear_buttons()

    def _on_close_request(self) -> None:
        self._continuous_busy = False
        self._continuous_steps_left = 0
        if self._mode == 0:
            self._persist_scenario_state()
        self.reject()


def create_data_agg_debug_dialog(
    parent: QWidget | None = None,
    debug_cfg: dict[str, Any] | None = None,
    live_items: list[dict[str, Any]] | None = None,
    scan_paths: list[str] | None = None,
    fixed_mode: int | None = None,
    scenario_for_dry_run: dict[str, Any] | None = None,
    scan_root: str | None = None,
) -> DataAggDebugDialog:
    """SCREENS.DEBUG を渡してデバッグダイアログを生成する。live_items 指定時は左一覧を実項目に合わせる。"""
    return DataAggDebugDialog(
        parent,
        debug_cfg,
        live_items,
        scan_paths,
        fixed_mode,
        scenario_for_dry_run,
        scan_root=scan_root,
    )
