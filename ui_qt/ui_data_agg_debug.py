# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_data_agg_debug.py
Purpose: データ集約デバッグウィンドウ（要求定義 §3.1.3）。文言・列見出し・ツールチップは config/ui_data_agg.json の SCREENS.DEBUG（TIP_*）。
History: 連続実行（シナリオ／マスタ一括）正常完了時に QMessageBox（SCREENS.DEBUG の MSG_RUN_ALL_*_DONE）。
  本番コードに内蔵デモデータは含めない。live_items なし時は空状態プレースホルダ。抽出は svc_data_agg_extract。マスタプレビュー（mpv）は進捗行ベースの結合表示＋列マージバッファ（svc.data_agg_master_preview / run_preview_compute）。結合探索なし・複数シナリオ時は項目内一括 compute＋段階キャッシュ先読み／バックフィル（DATA_AGG_MASTER_ONE_SHOT=0 で無効）。シナリオモード: 連携／結合フェーズかつ検出ファイルが多いとき、svc_data_agg_debug_run の progress_hook で非モーダル進捗を表示。build_master_items_live / _mpv_extract_colvals はファイル単位で xlsx_workbook_scope を張り .xlsx の load_workbook を再利用。
  2026-04-14: デバッグ—シナリオ/マスタでウィンドウタイトル（TITLE_SCENARIO/TITLE_MASTER）と連続実行ボタン（BTN_RUN_ALL_*、TIP_RUN_ALL_*）をモード連動。
  2026-04-14: 結果一覧: 列幅プログラム変更直後の遅延 sectionResized で user_resized が誤立ちしないよう、programmatic 解除を QTimer.singleShot(0) に遅延（世代で連続フィットに対応）。bump も同じセッション内で保護。
  2026-04-14: 診断: 結果一覧列幅—_fit_value_grid_columns で復元／内容フィットの分岐・viewport・先頭列幅・代表列 lo/hi/raw/fin を DATA_AGG_DIAG に出力。
  2026-04-14: 名前から取得—結果一覧で #n[項目] 列展開を抑止。COND_KEYS 等の「主キー」を「抜取り文字」に統一（JSON／既定）。
  2026-04-14: 条件タブ—マスタ親行は要約1行のみ（全文はツールチップ）。editor_lines あり時は details 子行を付けず二重解消。初期は collapseAll。
  2026-04-14: 条件タブ—要約（全文）Section 廃止（ツリーのみ）。ツリー開閉マーク上寄せ。結果サマリ／一覧ヘッダ省略抑止＋見出し幅で列拡張。シナリオ要約列はステップ単位・セル上寄せ。
  2026-04-13: SCREENS.DEBUG.SCENARIO_PROGRESS_MIN_FILES でシナリオ連携/結合フェーズのファイル進捗閾値を JSON 化。
  2026-04-13: SCREENS.DEBUG の TIP_* を全主要ウィジェットに反映。フォールバック文言を JSON 実体に整合。
  2026-04-13: cond_tree / master_cond_tree は QTreeWidget のため列見出しツールチップは header()（horizontalHeader は QTableWidget 専用で AttributeError）。
  結果一覧 value_grid: 全列 Interactive（最終列 Stretch なし）＋横スクロールで長文可読化。省略 … は delegate / TextElideMode で抑止。ユーザーが変えた列幅は同一ヘッダ構成のプレビュー更新で維持。縦横スクロールは AsNeeded＋ScrollPerPixel。WINDOW に VALUE_GRID_COL_* 等。
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
from typing import Any, Callable

from PySide6.QtCore import QEventLoop, Qt, QTimer
from PySide6.QtGui import (
    QBrush,
    QCloseEvent,
    QColor,
    QFont,
    QFontMetrics,
    QPainter,
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
from ui_qt.ui_common import _deep_merge, _normalize_message_newlines
from ui_qt.ui_common import create_progress_dialog
_data_agg_probe_log = get_data_agg_diag_logger()

from svc.data_agg_master_preview import (
    FROZEN_SNAPSHOT_VERSION,
    frozen_snapshot_invalid_reason,
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
        lines.append(
            "#%d セル=%s 行=%s 列=%s 項目=%s"
            % (
                i + 1,
                jd.get("cell", ""),
                jd.get("row", ""),
                jd.get("col", ""),
                jd.get("item", ""),
            )
        )
    return lines or ["（結合キー定義なし）"]


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
            join_defs = p.get("join_defs") if isinstance(p.get("join_defs"), list) else []
            link_defs = p.get("link_defs") if isinstance(p.get("link_defs"), list) else []
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
        from svc.svc_data_agg_extract import extract_item_bundle, xlsx_workbook_scope
    except Exception:
        extract_item_bundle = None  # type: ignore[misc, assignment]
        xlsx_workbook_scope = None  # type: ignore[misc, assignment]

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
                slots: list[Any] = rows[0].get("slots") if rows else []
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
                        try:
                            from svc.svc_data_agg_extract import name_extract_hit_files_ordered
                        except Exception:
                            name_extract_hit_files_ordered = None  # type: ignore[misc, assignment]
                        typ_src = str(src.get("type") or "cell").strip().lower()
                        paths_iter = paths
                        if typ_src in ("metadata", "meta", "filename", "name_extract") and name_extract_hit_files_ordered:
                            paths_iter = name_extract_hit_files_ordered(paths, src)
                        for fp in paths_iter:
                            if len(col_vals) >= max_rows:
                                break
                            with xlsx_workbook_scope():  # type: ignore[misc]
                                try:
                                    jp_hdr = str(one.get("name") or one.get("id") or "").strip()
                                    b = extract_item_bundle(
                                        fp,
                                        one,
                                        item_id=item_id,
                                        cell_positions={},
                                        join_path_header=jp_hdr or None,
                                    )
                                except Exception:
                                    b = {"primary_values": [None]}
                            prim = b.get("primary_values") or [None]
                            for v in prim:
                                if len(col_vals) >= max_rows:
                                    break
                                col_vals.append("" if v is None else str(v))
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


# マスタ項目ステップ実行の進捗（8 段）。phase_i は 1 始まりで本配列の index+1 に対応。
_MASTER_DEBUG_PROGRESS_PHASES: tuple[str, ...] = (
    "準備しています",
    "サマリー表を更新中",
    "サマリーに項目の値を表示中",
    "結果一覧用に取り出し中",
    "取得データを行にまとめ中",
    "項目間で照合中",
    "結果一覧表を組み立て表示中",
    "完了しました",
)

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

    def initStyleOption(self, option: QStyleOptionViewItem, index) -> None:  # type: ignore[override]
        super().initStyleOption(option, index)
        option.textElideMode = Qt.TextElideMode.ElideNone


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
        # mpv: 現在項目列の差分マージ用 2D バッファ。描画は進捗行と合成する。
        self._mpv_grid: list[list[Any]] | None = None
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
        # mpv 描画: 項目ごとの progress 行キャッシュ（step_idx, rows）
        self._mpv_progress_rows_by_mi: dict[int, tuple[int, list[list[Any]]]] = {}
        # 直近まで compute 済みの項目（実行可能シナリオなし項目へ移ったときの prog フォールバック用）
        self._last_master_completed_mi_idx: int | None = None
        # mpv: 完了項目列の凍結（行キー __norm_path + __iter_index）。次項目 compute の再走査を抑える。
        self._mpv_frozen_snapshots: dict[int, dict[str, Any]] = {}
        # 描画時に「現在列」として扱う項目 index（フォールバック表示整合用）
        self._mpv_display_mi_idx: int | None = None
        # シナリオなし項目の直後に「実行あり」項目へ入ったとき、入場直後の value グリッド再構築を
        # その項目の全ステップ完了時（離脱直前）まで遅延する。対象 mi（到着先の index）。
        self._mpv_deferred_value_grid_mi: int | None = None
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
        self._run_progress_dlg: Any | None = None
        self._run_progress_path: Path | None = None
        self._run_progress_seq: int = 0
        self._debug_progress_locked: bool = False
        self._master_item_snapshots: dict[int, dict[str, Any]] = {}
        self._master_item_snapshot_done: set[int] = set()

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
            apply_window_config(
                self, {"WINDOW": self._cfg.get("WINDOW") or {}}, ph, "DEBUG"
            )
        except Exception:
            pass
        self._apply_mode()
        self._refresh_all()

    def showEvent(self, event: Any) -> None:
        """親（メイン／シナリオ編集）の中央付近に重ねて表示する。"""
        super().showEvent(event)
        pw = self.parentWidget()
        if pw is not None:
            pr = pw.frameGeometry()
            gr = self.frameGeometry()
            x = pr.x() + (pr.width() - gr.width()) // 2
            y = pr.y() + (pr.height() - gr.height()) // 2
            self.move(x, y)
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

    def _d(self, key: str, default: str) -> str:
        s = _normalize_message_newlines(str(self._cfg.get(key) or default).strip())
        if key.endswith("_HTML"):
            return s.replace("\n", "<br/>")
        return s

    def _tip(self, key: str, default: str = "") -> str:
        """DEBUG 用ツールチップ文言（プレーン）。JSON の TIP_* を想定。"""
        return _normalize_message_newlines(str(self._cfg.get(key) or default).strip())

    def _set_tip(self, w: QWidget | None, key: str, default: str = "") -> None:
        if w is None:
            return
        t = self._tip(key, default).strip()
        if not t and default:
            t = _normalize_message_newlines(str(default).strip())
        if t:
            w.setToolTip(t)

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
        restore = (
            self._value_grid_user_resized
            and self._value_grid_saved_widths
            and len(self._value_grid_saved_widths) == n
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
                len(self._value_grid_saved_widths or []),
                cur_key == self._value_grid_structure_key,
            )
        except Exception:
            pass
        if restore:
            hdr = self.value_grid.horizontalHeader()
            floor = hdr.minimumSectionSize()
            self._value_grid_programmatic_gen += 1
            gen = self._value_grid_programmatic_gen
            self._value_grid_header_programmatic = True
            try:
                hdr.blockSignals(True)
                for c, w in enumerate(self._value_grid_saved_widths):
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

    def _max_value_rows(self) -> int:
        try:
            return int(self._cfg.get("MAX_VALUE_ROWS") or MAX_VALUE_ROWS_DEFAULT)
        except (TypeError, ValueError):
            return MAX_VALUE_ROWS_DEFAULT

    def _master_preview_read_rows(self) -> int:
        try:
            return max(
                self._max_value_rows(),
                int(self._cfg.get("MASTER_PREVIEW_READ_ROWS") or 70),
            )
        except (TypeError, ValueError):
            return max(self._max_value_rows(), 70)

    def _master_preview_display_rows(self) -> int:
        try:
            return max(1, int(self._cfg.get("MASTER_PREVIEW_DISPLAY_ROWS") or self._max_value_rows()))
        except (TypeError, ValueError):
            return max(1, self._max_value_rows())

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
        try:
            QApplication.processEvents(
                QEventLoop.ProcessEventsFlag.ExcludeUserInputEvents
            )
        except Exception:
            try:
                QApplication.processEvents()
            except Exception:
                pass

    def _set_debug_progress_locked(self, locked: bool) -> None:
        if self._mode != 1:
            return
        self._debug_progress_locked = locked
        self._refresh_master_nav_lock_state()

    def _show_run_progress(
        self,
        phase: str,
        done: int,
        total: int,
        *,
        window_title: str = "",
        pct_override: int | None = None,
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
            self._master_progress_pct_floor = pct
            ph = self._debug_parent_hwnd()
            if self._run_progress_dlg is None or self._run_progress_path is None:
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
                    "non_modal_progress": True,
                    "done_delay_ms": 220,
                    "center_on_parent_widget": True,
                }
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
                self._run_progress_seq = 0
                self._set_debug_progress_locked(True)
            self._run_progress_seq += 1
            wt = str(window_title or "").strip()
            ipc_file.write_pickle(
                self._run_progress_path,
                {
                    "status": "RUN",
                    "seq": int(self._run_progress_seq),
                    "phase_total": tot,
                    "phase_i": max(1, dn) if dn > 0 else 1,
                    "phase": str(phase or "実行中"),
                    "done": dn,
                    "total": tot,
                    "pct": pct,
                    "current_file": "",
                    "window_title": wt,
                    "msg": "",
                },
            )
            self._process_events_light()
        except Exception:
            pass

    def _master_dbg_batch_progress_hook(self, sub_phase: int, detail: str, *rest: Any) -> None:
        """compute_batch_table_rows からのコールバック（phase 4〜7）。rest は file_index, n_files（任意）。"""
        dlg = getattr(self, "_run_progress_dlg", None)
        try:
            dlg_open = dlg is not None and dlg.isVisible()
        except Exception:
            dlg_open = False
        if not getattr(self, "_master_run_progress_active", False) and not dlg_open:
            return
        if not (4 <= sub_phase <= 7):
            return
        msg = _MASTER_DEBUG_PROGRESS_PHASES[sub_phase - 1] + str(detail or "")
        wt = getattr(self, "_master_progress_window_title", "") or ""
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
        nf = int(getattr(self, "_master_batch_hook_last_nf", 1) or 1)
        nf = max(1, nf)
        denom = max(1, 4 * nf)
        mi_m = re.search(r"項目\s*(\d+)\s*/\s*(\d+)", detail_s)
        row_m = re.search(r"行\s*(\d+)\s*/\s*(\d+)", detail_s)
        pct_ov: int
        if sub_phase == 4 and mi_m:
            inn = max(1, int(mi_m.group(2)))
            ii = min(inn, max(0, int(mi_m.group(1))))
            step_start = (fi - 1) * 4 + 1
            step_end = (fi - 1) * 4 + 2
            p0 = 37.5 + 50.0 * float(step_start) / float(denom)
            p1 = 37.5 + 50.0 * float(step_end) / float(denom)
            t = min(1.0, max(0.0, float(ii) / float(inn)))
            pct_ov = int(round(p0 + (p1 - p0) * t))
        elif sub_phase == 7 and row_m:
            rnn = max(1, int(row_m.group(2)))
            rr = min(rnn, max(0, int(row_m.group(1))))
            step_start = (fi - 1) * 4 + 3
            step_end = (fi - 1) * 4 + 4
            p0 = 37.5 + 50.0 * float(step_start) / float(denom)
            p1 = 37.5 + 50.0 * float(step_end) / float(denom)
            t = min(1.0, max(0.0, float(rr) / float(rnn)))
            pct_ov = int(round(p0 + (p1 - p0) * t))
        else:
            step_k = (fi - 1) * 4 + (sub_phase - 4) + 1
            pct_ov = int(round(37.5 + 50.0 * float(step_k) / float(denom)))
        self._show_run_progress(
            msg,
            sub_phase,
            len(_MASTER_DEBUG_PROGRESS_PHASES),
            window_title=wt,
            pct_override=pct_ov,
        )
        self._process_events_light()

    def _close_run_progress(self) -> None:
        dlg = self._run_progress_dlg
        try:
            if self._run_progress_path is not None:
                self._run_progress_seq += 1
                ipc_file.write_pickle(
                    self._run_progress_path,
                    {
                        "status": "DONE",
                        "seq": int(self._run_progress_seq),
                        "phase_i": 8,
                        "phase": "完了しました",
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
            self._run_progress_dlg = None
            self._run_progress_path = None
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
        tt = list(tips or [])
        if len(tt) != len(colvals):
            tt = [None] * len(colvals)
        if any(re.match(r"^#\d+\[[^\]]*\]", str(x)) for x in colvals):
            return list(colvals), tt
        cap = self._max_value_rows()
        if len(colvals) == cap + 1 and colvals and "省略" in str(colvals[-1]):
            return list(colvals), tt
        if len(colvals) <= cap:
            return list(colvals), tt
        out = _cap_list_capped(colvals, cap)
        out_t = list(tt[:cap]) + [None]
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
                [None] * len(c) for c in self._value_cols
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
        self.value_grid.setItemDelegate(_ValueGridNoElideDelegate(self.value_grid))
        _vgh = self.value_grid.horizontalHeader()
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
        return int(self._mi_idx) in self._master_item_snapshots

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

    def _apply_master_left_registered_row_style(self) -> None:
        """マスタ一覧: 実行可能シナリオ登録ありの行は薄ベージュ＋項目名列を青太字。スナップショット示唆は結果一覧のみ。"""
        if self._mode != 1:
            return
        try:
            mit = self._master_table_items()
            nr = self.left_table.rowCount()
            nc = self.left_table.columnCount()
            clear_bg = QBrush()
            default_fg = QBrush()
            beige = QBrush(_DEBUG_MASTER_REGISTERED_ROW_BG)
            blue = QBrush(_DEBUG_MASTER_REGISTERED_NAME_COLOR)
            for ri in range(nr):
                reg = (
                    ri < len(mit)
                    and self._master_active_count_for_item(mit[ri]) > 0
                )
                for ci in range(nc):
                    it = self.left_table.item(ri, ci)
                    if it is None:
                        continue
                    if reg:
                        it.setBackground(beige)
                        if ci == 1:
                            f = it.font()
                            f.setBold(True)
                            it.setFont(f)
                            it.setForeground(blue)
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
            mpr = self._master_preview_read_rows()
            mpd = self._master_preview_display_rows()
            ma_hint = self._d("HINT_MASTER_HTML", "") or (
                "<b>マスタ項目ステップ</b>：最終項目まで実行すると周回せず待機します（スナップショット閲覧が可能なときは左上コーナーが薄青）。"
                " 本番経路の一覧を再計算し、読込上限<b>%d行</b>・表示上限<b>%d行</b>で表示。"
                % (mpr, mpd)
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
            self.res_hint.setText(
                self._d("RES_HINT_MASTER_HTML", "")
                or (
                    "<b>結果サマリ</b>：全ステップを積み重ね。"
                    " <b>結果一覧</b>：結合後テーブルの最大表示行でプレビューします。"
                )
            )
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
        self.values_title.setText(fmt % m["title"])
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
        self._bump_mpv_prefetch_cancel()
        self._scenario_bundle_caches.clear()
        self._master_sparse_notice_shown = False
        self._mpv_extract_cache.clear()
        self._mpv_colvals_cache.clear()
        self._mpv_progress_rows_cache = None
        self._mpv_progress_rows_step_cache.clear()
        self._mpv_progress_rows_by_mi.clear()
        self._mpv_frozen_snapshots.clear()
        self._last_master_completed_mi_idx = None
        self._mpv_display_mi_idx = None
        self._mpv_deferred_value_grid_mi = None
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
        """連携・結合フェーズかつ検出ファイル数が多いときのみファイル単位進捗を出す。"""
        if int(phase_gi) not in (3, 4):
            return False
        floor = int(getattr(self, "_scenario_progress_min_files", 0) or 0)
        return int(n_paths or 0) >= floor

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
            self._show_run_progress(phase_msg, done, total, window_title=wt)

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
                            scenario_source_tooltip_plain(src, dn, detail_cell_cfg=dcell)
                        )
                    self.left_table.setItem(i, 0, it)
                self.left_table.selectRow(self._sc_idx)
                _lh0 = self.left_table.horizontalHeader()
                _lh0.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            else:
                self.left_table.setColumnCount(2)
                self.left_table.setHorizontalHeaderLabels(
                    [
                        self._d("LEFT_TABLE_COL_INDEX", "#"),
                        self._d("LEFT_TABLE_HEADER_MASTER", "項目名"),
                    ]
                )
                mit = self._master_table_items()
                self.left_table.setRowCount(len(mit))
                for i, m in enumerate(mit):
                    self.left_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                    self.left_table.setItem(
                        i, 1, QTableWidgetItem(str(m.get("title") or ""))
                    )
                self.left_table.selectRow(self._mi_idx)
                _lh = self.left_table.horizontalHeader()
                _lh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
                _lh.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
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
                self._persist_scenario_state()
                self._sc_idx = r
                self._load_scenario_state(self._sc_idx)
        else:
            if r != self._mi_idx:
                self._bump_mpv_prefetch_cancel()
                self._mpv_deferred_value_grid_mi = None
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
            if self._master_snapshot_priority_active() and r in self._master_item_snapshots:
                self._apply_master_item_snapshot(r)
            else:
                self._sync_summary_table_from_lists()
                self._rebuild_value_grid()
        if self._mode == 1:
            self._update_left_detail()
        self._paint_result_highlights()
        self._update_run_buttons_state()
        self._update_clear_buttons()

    def _clear_master_item_snapshots(self) -> None:
        self._master_item_snapshots.clear()
        self._master_item_snapshot_done.clear()
        self._refresh_master_snapshot_chrome()

    def _master_snapshot_priority_active(self) -> bool:
        if self._mode != 1:
            return False
        items = self._master_table_items()
        if not items:
            return False
        return len(self._master_item_snapshot_done) >= len(items)

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
        if empty:
            self._master_item_snapshots[completed_mi] = {"empty": True}
        else:
            gh: list[str] = []
            nc = self.value_grid.columnCount()
            for c in range(nc):
                hi = self.value_grid.horizontalHeaderItem(c)
                gh.append("" if hi is None else str(hi.text()))
            gr: list[list[str]] = []
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
            }
        self._master_item_snapshot_done.add(completed_mi)
        self._refresh_master_snapshot_chrome()

    def _apply_value_grid_from_snapshot(self, headers: list[str], rows: list[list[str]]) -> None:
        self._mpv_join_table_active = False
        self._mpv_join_table_ncols = 0
        nc = len(headers)
        nr = len(rows)
        self._value_grid_note_structure([str(h) for h in headers])
        self.value_grid.clear()
        self.value_grid.setColumnCount(max(0, nc))
        self.value_grid.setRowCount(max(0, nr))
        if nc > 0:
            self.value_grid.setHorizontalHeaderLabels([str(h) for h in headers])
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
                cell.setToolTip(tx)
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
        jdefs = p.get("join_defs") if isinstance(p.get("join_defs"), list) else []
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
                    it0.setToolTip(row_tip)
                    it1.setToolTip(row_tip)
                    self.left_steps.setItem(r, 0, it0)
                    self.left_steps.setItem(r, 1, it1)
            else:
                m = self._current_master()
                for _li, si in enumerate(self._active_slot_indices):
                    sc = m["scenarios"][si]
                    slot = sc["slot"]
                    assert slot is not None
                    r = self.left_steps.rowCount()
                    self.left_steps.insertRow(r)
                    step_txt = str(sc["title"] or "シナリオ")
                    it0 = QTableWidgetItem(str(self._display_step_no(r)))
                    it1 = QTableWidgetItem(step_txt)
                    for it in (it0, it1):
                        it.setTextAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
                    src1 = sc.get("source")
                    if isinstance(src1, dict):
                        row_tip = scenario_source_tooltip_plain(src1, dn, detail_cell_cfg=dcell)
                    else:
                        row_tip = _format_condition_step_tooltip(step_txt, slot)
                    for it in (it0, it1):
                        it.setToolTip(row_tip)
                    self.left_steps.setItem(r, 0, it0)
                    self.left_steps.setItem(r, 1, it1)
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
                    top.setToolTip(col, step_tip)
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
                    top.setToolTip(col, row_tip)
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
        items = list((self._scenario_for_dry_run or {}).get("items") or [])
        return [
            it.get("name") or it.get("id") or ("項目_%s" % i)
            for i, it in enumerate(items)
        ]

    def _mpv_preview_compute_paths(self) -> list[str]:
        """compute_batch 内部 filter と同じ絞り込み後パス（凍結検証用）。"""
        return preview_compute_file_paths(
            self._scenario_for_dry_run or {},
            list(self._debug_scan_paths or []),
        )

    def _mpv_frozen_context_for_mi(self, mi_idx: int) -> tuple[dict[str, Any] | None, int | None]:
        """次項目 compute 用: (frozen_prior, frozen_through_mi)。不適格時は (None, None)。"""
        if not self._mpv_frozen_columns_enabled():
            return None, None
        if int(mi_idx) <= 0:
            return None, None
        expected_through = int(mi_idx) - 1
        snap = self._mpv_frozen_snapshots.get(expected_through)
        headers = self._mpv_preview_headers()
        paths = self._mpv_preview_compute_paths()
        reason = frozen_snapshot_invalid_reason(
            snap,
            headers=headers,
            file_paths=paths,
            expected_through_mi=expected_through,
        )
        if reason is not None:
            try:
                scan_n = len(self._debug_scan_paths or [])
                _data_agg_probe_log.info(
                    "[DATA_AGG_DIAG] mpv_frozen skip mi_idx=%s reason=%s "
                    "expected_through=%s snap_through=%s paths_filter=%s scan_paths=%s",
                    mi_idx,
                    reason,
                    expected_through,
                    (snap or {}).get("through_mi") if isinstance(snap, dict) else None,
                    len(paths),
                    scan_n,
                )
            except Exception:
                pass
            return None, None
        return snap, expected_through

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
        frozen_prior, frozen_through = self._mpv_frozen_context_for_mi(int(mi_idx))
        cap_acc: list[dict[str, Any]] | None = (
            [] if frozen_capture_out is not None else None
        )
        scen = scenario_for_stepped_preview(
            scenario_base,
            mi_idx=int(mi_idx),
            master_step_idx=int(master_step_idx),
            active_slot_indices=list(active_slot_indices),
            use_max_sources_for_current_item=bool(use_max_sources),
            frozen_through_mi=frozen_through,
            frozen_prior=frozen_prior,
            frozen_capture_out=frozen_capture_out,
            frozen_capture_acc=cap_acc,
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
        _h, table_rows, _ev, _jt = run_preview_compute(
            scen,
            scan_paths,
            max_primary_rows=self._master_preview_read_rows(),
            max_table_rows=self._master_preview_display_rows(),
            progress_hook=progress_hook,
            probe_caller=probe_caller,
        )
        rows_out = [list(r) for r in table_rows]
        if frozen_capture_out is not None:
            self._mpv_store_frozen_snapshot(frozen_capture_out)
        return rows_out

    def _mpv_store_step_cache(
        self,
        sk: tuple[Any, ...],
        rows: list[list[Any]],
        *,
        mi_idx: int,
        master_step_idx: int,
    ) -> None:
        self._mpv_progress_rows_step_cache[sk] = [list(r) for r in rows]
        self._mpv_progress_rows_by_mi[int(mi_idx)] = (
            int(master_step_idx),
            list(rows),
        )
        n_act = len(self._active_slot_indices or [])
        if rows and (n_act <= 0 or int(master_step_idx) >= n_act):
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
                    self._mpv_prog_compute_lock.release()
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
        """結合プレビュー行から当該列を切り出す（extract 回避）。"""
        if not self._scenario_for_dry_run or not self._debug_scan_paths:
            return False
        if int(self._master_step_idx) <= 1:
            return True
        if not self._mpv_one_shot_eligible():
            return False
        n_act = len(self._active_slot_indices or [])
        if n_act <= 0:
            return False
        sk_full = self._mpv_progress_step_cache_key(n_act)
        return sk_full in self._mpv_progress_rows_step_cache

    def _bump_mpv_prefetch_cancel(self) -> None:
        self._mpv_prefetch_cancel_gen += 1

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
            if sk_next in self._mpv_progress_rows_step_cache:
                continue
            if not self._mpv_prog_compute_lock.acquire(blocking=False):
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_progress prefetch=skip_busy n_pick=%s",
                        sk_next[1] if len(sk_next) > 1 else "?",
                    )
                except Exception:
                    pass
                continue
            out_rows: list[list[Any]] = []
            apply_after = False
            try:
                if cancel_gen != self._mpv_prefetch_cancel_gen:
                    pass
                elif sk_next in self._mpv_progress_rows_step_cache:
                    pass
                else:
                    dlg = self
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
            except Exception:
                _logger.exception("mpv progress prefetch failed")
            finally:
                self._mpv_prog_compute_lock.release()
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
        if sk in self._mpv_progress_rows_step_cache:
            return
        self._mpv_progress_rows_step_cache[sk] = [list(r) for r in rows]
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_progress prefetch=done n_pick=%s rows=%s one_shot=%s",
                sk[1] if len(sk) > 1 else "?",
                len(rows),
                schedule_backfill,
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
        try:
            done_np = int(sk[1])
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
        self, *, next_master_step_override: int | None = None
    ) -> None:
        from core import core_env

        if not core_env.data_agg_master_progress_prefetch_enabled():
            return
        if self._mode != 1:
            return
        if not self._scenario_for_dry_run or not self._debug_scan_paths:
            return
        act = self._active_slot_indices or []
        n_act = len(act)
        if n_act <= 0:
            return
        one_shot = self._mpv_one_shot_eligible()
        if one_shot:
            next_master_step = n_act
            use_max_sources = True
            schedule_backfill = True
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
        max_pr = self._master_preview_read_rows()
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

    def _mpv_progress_batch_rows(self) -> list[list[Any]]:
        try:
            return self._mpv_progress_batch_rows_impl()
        finally:
            self._mpv_request_progress_prefetch_debounced()

    def _mpv_progress_batch_rows_impl(self) -> list[list[Any]]:
        """
        マスタプレビュー表示用: 完了項目はフル、現在項目は実行済みシナリオ分のみ、未到達項目は主値ソースなし。
        連携・結合は到達範囲のパイプラインに含まれる列へ反映され、未到達列へのフル結果の透けを防ぐ。
        """
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
            n_pick_req = self._mpv_progress_n_pick()
            sk = self._mpv_progress_step_cache_key(n_pick_req)
            cached_step = self._mpv_progress_rows_step_cache.get(sk)
            if cached_step is not None:
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
        with self._mpv_prog_compute_lock:
            hit_after_wait = self._mpv_progress_rows_step_cache.get(sk_need)
            if hit_after_wait is not None:
                rows = [list(r) for r in hit_after_wait]
                from_prefetch_wait = True
            else:
                n_act_compute = len(self._active_slot_indices or [])
                use_max = bool(
                    self._mpv_one_shot_eligible()
                    and n_pick_req == n_act_compute
                    and n_act_compute > 0
                )
                frozen_cap: dict[str, Any] | None = None
                if (
                    self._mpv_frozen_columns_enabled()
                    and n_act_compute > 0
                    and n_pick_req >= n_act_compute
                ):
                    frozen_cap = {}
                _off_hook = None
                try:
                    _pd = getattr(self, "_run_progress_dlg", None)
                    if _pd is not None and _pd.isVisible():
                        _off_hook = self._master_dbg_batch_progress_hook
                except Exception:
                    _off_hook = None
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

    def _mpv_extract_colvals(self, mi_idx: int, si: int) -> list[str]:
        """
        マスタプレビュー用: 本番と同じ extract_item_bundle 経路で主値列を得る。
        build_master_items_live(..., preload_values=True) と同系統（行順はファイル走査順）。
        """
        key = (mi_idx, si)
        if key in self._mpv_extract_cache:
            return list(self._mpv_extract_cache[key])
        t0 = time.perf_counter()
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
            from svc.svc_data_agg_extract import extract_item_bundle, xlsx_workbook_scope
        except Exception:
            extract_item_bundle = None  # type: ignore[misc, assignment]
            filter_file_paths_for_master_preview = None  # type: ignore[misc, assignment]
            xlsx_workbook_scope = None  # type: ignore[misc, assignment]
        paths = [str(p).strip() for p in self._debug_scan_paths if str(p).strip()]
        n_paths_before = len(paths)
        if extract_item_bundle is None:
            out = ["（svc_data_agg_extract を読み込めませんでした）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        if filter_file_paths_for_master_preview is not None:
            paths = list(filter_file_paths_for_master_preview(paths, items))
        n_paths_after = len(paths)
        if not paths:
            out = ["（検出ファイルがありません）"]
            self._mpv_extract_cache[key] = list(out)
            return out
        one = {**item, "sources": [copy.deepcopy(src)]}
        item_id = str(item.get("id") or item.get("name") or "").strip()
        try:
            from svc.svc_data_agg_extract import name_extract_hit_files_ordered
        except Exception:
            name_extract_hit_files_ordered = None  # type: ignore[misc, assignment]
        typ_src = str(src.get("type") or "cell").strip().lower()
        paths_iter: Any = paths
        if typ_src in ("metadata", "meta", "filename", "name_extract") and name_extract_hit_files_ordered:
            paths_iter = name_extract_hit_files_ordered(paths, src)
        max_rows = self._max_value_rows()
        col_vals: list[str] = []
        _pe_n = 0
        for fp in paths_iter:
            if len(col_vals) >= max_rows:
                break
            with xlsx_workbook_scope():  # type: ignore[misc]
                try:
                    jp_hdr = str(one.get("name") or one.get("id") or "").strip()
                    b = extract_item_bundle(
                        fp,
                        one,
                        item_id=item_id or None,
                        cell_positions={},
                        join_path_header=jp_hdr or None,
                    )
                except Exception:
                    b = {"primary_values": [None]}
            prim = b.get("primary_values") or [None]
            for v in prim:
                if len(col_vals) >= max_rows:
                    break
                col_vals.append("" if v is None else str(v))
            _pe_n += 1
            if _pe_n % 8 == 0:
                self._process_events_light()
        if not col_vals:
            col_vals = ["（該当する主値がありません）"]
        try:
            _data_agg_probe_log.info(
                "[DATA_AGG_DIAG] mpv_extract mi_idx=%s si=%s src_type=%s "
                "paths_before=%s paths_after=%s col_count=%s col_head=%s elapsed_ms=%s",
                mi_idx,
                si,
                typ_src,
                n_paths_before,
                n_paths_after,
                len(col_vals),
                col_vals[:5],
                int((time.perf_counter() - t0) * 1000),
            )
        except Exception:
            pass
        self._mpv_extract_cache[key] = list(col_vals)
        return list(col_vals)

    def _merge_mpv_column(self, mi_idx: int, colvals: list[str]) -> None:
        """mpv: 現在のマスタ項目列だけを書込みモードに従ってマージする。"""
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
        read_cap = self._master_preview_read_rows()
        nrows = min(
            read_cap,
            max(
                self._master_preview_display_rows(),
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
        items = list(scen.get("items") or [])
        headers = [
            str(it.get("name") or it.get("id") or ("項目_%s" % i))
            for i, it in enumerate(items)
        ]
        ncols = len(headers)
        if ncols == 0:
            self._reset_value_grid()
            self._paint_result_highlights()
            return
        prog_rows: list[list[Any]] = self._mpv_progress_batch_rows()
        pr = len(prog_rows)
        read_cap = self._master_preview_read_rows()
        display_floor = self._master_preview_display_rows()
        # table_rows を正とする。旧マージバッファの行数で表示行を伸ばさない。
        row_basis = max(pr, display_floor)
        max_r = max(1, min(read_cap, row_basis))
        disp_mi = self._mpv_display_mi_idx
        if disp_mi is None or disp_mi < 0 or disp_mi >= ncols:
            disp_mi = self._mi_idx
        mi = max(0, min(int(disp_mi), ncols - 1))
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
                cell.setToolTip(tx)
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
        self._fit_value_grid_columns()
        self._paint_result_highlights()

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
        try:
            self._rebuild_value_grid_impl()
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
        keep_stable_during_continuous = self._mode == 1 and (
            bool(getattr(self, "_continuous_busy", False))
            or bool(getattr(self, "_master_step_loop_busy", False))
        )
        if finalize:
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
        if self._mode == 1 and self._scenario_for_dry_run and self._debug_scan_paths:
            if not self._summary_rows:
                items = list((self._scenario_for_dry_run or {}).get("items") or [])
                headers = [
                    str(it.get("name") or it.get("id") or ("項目_%s" % i))
                    for i, it in enumerate(items)
                ]
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
            groups: dict[str, tuple[str, list[str]]] = {}
            if not (
                self._mode == 0 and self._scenario_source_kind() == "name_extract"
            ):
                for v in colvals:
                    m = re.match(r"^#(\d+)\[([^\]]*)\]\s*(.*)$", str(v))
                    if not m:
                        groups = {}
                        break
                    key = m.group(1)
                    tgt = (m.group(2) or "").strip() or "未指定"
                    val = m.group(3)
                    if key not in groups:
                        groups[key] = (tgt, [])
                    groups[key][1].append(val)
            if groups:
                start = len(expanded)
                for k in sorted(groups.keys(), key=lambda x: int(x)):
                    tgt, vals = groups[k]
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
            expanded = self._append_scenario_join_columns_if_needed(expanded)

        ncols = len(expanded)
        if ncols == 0:
            self._reset_value_grid()
            self._paint_result_highlights()
            return
        max_r = max(len(col) for _, col, _ in expanded) if expanded else 0
        headers = [h for h, _, _ in expanded]
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
                it.setToolTip(tip_txt)
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
        ct = list(col_tooltips or [])
        if len(ct) != len(colvals):
            ct = [None] * len(colvals)
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
        self._mpv_progress_rows_step_cache.clear()
        self._mpv_progress_rows_by_mi.clear()
        self._mpv_frozen_snapshots.clear()
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
        ct = list(tips or [])
        if len(ct) != len(colvals):
            ct = [None] * len(colvals)
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
        self, *, mode: int, was_full_master: bool, steps: int
    ) -> None:
        """連続実行の正常完了時に JSON 文言で終了メッセージを表示する。"""
        title = self._d("DIALOG_RUN_ALL_DONE_TITLE", "データ集約 デバッグ")
        if mode == 0:
            tpl = self._d(
                "MSG_RUN_ALL_SCENARIO_DONE",
                "シナリオの連続実行が完了しました。\n実行ステップ数: {steps}",
            )
        elif was_full_master:
            tpl = self._d(
                "MSG_RUN_ALL_MASTER_ITEMS_DONE",
                "全項目の連続実行が完了しました。\n実行ステップ数: {steps}",
            )
        else:
            tpl = self._d(
                "MSG_RUN_ALL_MASTER_DONE",
                "項目の連続実行が完了しました。\n実行ステップ数: {steps}",
            )
        try:
            body = tpl.format(steps=int(steps))
        except Exception:
            body = tpl
        QMessageBox.information(
            self, title, _normalize_message_newlines(body)
        )

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
        snap_view = self._mode == 1 and self._master_showing_row_snapshot()
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

    def _clear_current_scenario_results_only(self) -> None:
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
                    )
                finally:
                    if prog_hook is not None:
                        self._close_run_progress()
                colvals, col_tips = self._icap_with_tips(colvals, col_tips)
                sid = str(item_live.get("id") or "")
                for ln in format_synthetic_events_for_log(events, sid):
                    self._log_prepend_plain(ln)
            else:
                vals = list(slot["summary_vals"])
                colvals = self._icap(list(slot["values_column"]))
                col_tips = [None] * len(colvals)
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
                else:
                    self._master_finish_step_pass_idle()
                # 実行可能シナリオが無い項目の「次」が実行ありのとき、入場直後の再構築は
                # その項目の全ステップ完了時（_flush_deferred_master_value_grid_if_mi）へ遅延する。
                if self._active_slot_indices:
                    self._mpv_deferred_value_grid_mi = int(self._mi_idx)
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] value_grid_defer_post_empty_skip landing_mi=%s",
                            self._mi_idx,
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
                if done_mi == 0:
                    focus_results_tab = True
                continue
            break
        m = self._current_master()
        _mt_item = str(m.get("title") or m.get("name") or m.get("id") or "").strip() or "項目"
        if self._master_step_idx == 0:
            self._log_append_master_item_row(
                _mt_item, "実行開始", item_number=self._mi_idx + 1
            )
        si = self._active_slot_indices[self._master_step_idx]
        sc = m["scenarios"][si]
        slot = sc["slot"]
        assert slot is not None
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
        self._show_run_progress(
            _MASTER_DEBUG_PROGRESS_PHASES[0], 1, sub_total, window_title=prog_wt
        )
        self._process_events_light()
        vals = list(slot["summary_vals"])
        gno = self._master_session_start_step + self._master_step_idx + 1
        sc_title = str(sc["title"] or m["title"])
        plab = self._phase_label(gno, sc_title)
        plab_summary = self._summary_first_col_label(gno, sc_title)
        colvals = self._icap(list(slot.get("values_prod", slot["values_column"])))
        if self._scenario_for_dry_run and self._debug_scan_paths:
            # 速度優先: まず progress 側の本番同等行から現在列を取り、取れない場合のみ extract にフォールバック。
            use_progress_colvals = self._mpv_can_colvals_from_progress()
            n_pick_now = int(self._mpv_progress_n_pick())
            if use_progress_colvals and n_pick_now <= 0:
                use_progress_colvals = False
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_colvals_from_progress skip=n_pick_zero "
                        "mi_idx=%s step_idx=%s si=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        si,
                    )
                except Exception:
                    pass
            col_from_prog: list[str] = []
            if use_progress_colvals:
                try:
                    t_prog_col = time.perf_counter()
                    prog_rows_now = self._mpv_progress_batch_rows()
                    for rr in prog_rows_now[: self._max_value_rows()]:
                        v = rr[self._mi_idx] if self._mi_idx < len(rr) else None
                        col_from_prog.append("" if v is None else str(v))
                    while col_from_prog and (not str(col_from_prog[-1]).strip()):
                        col_from_prog.pop()
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_colvals_from_progress mi_idx=%s step_idx=%s si=%s "
                        "row_count=%s col_count=%s elapsed_ms=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        si,
                        len(prog_rows_now),
                        len(col_from_prog),
                        int((time.perf_counter() - t_prog_col) * 1000),
                    )
                except Exception:
                    col_from_prog = []
            else:
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_colvals_from_progress skip=no_progress_cache "
                        "mi_idx=%s step_idx=%s si=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        si,
                    )
                except Exception:
                    pass
            if col_from_prog:
                colvals = self._icap(col_from_prog)
            else:
                t_extract = time.perf_counter()
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_extract_start mi_idx=%s step_idx=%s si=%s title=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        si,
                        str(sc.get("title") or ""),
                    )
                except Exception:
                    pass
                colvals = self._icap(self._mpv_extract_colvals(self._mi_idx, si))
                try:
                    _data_agg_probe_log.info(
                        "[DATA_AGG_DIAG] mpv_extract_end mi_idx=%s step_idx=%s si=%s col_count=%s elapsed_ms=%s",
                        self._mi_idx,
                        self._master_step_idx,
                        si,
                        len(colvals),
                        int((time.perf_counter() - t_extract) * 1000),
                    )
                except Exception:
                    pass
            self._mpv_colvals_cache[(int(self._mi_idx), int(si))] = list(colvals)
        gr = self._master_global_row_idx
        self._show_run_progress(
            _MASTER_DEBUG_PROGRESS_PHASES[1], 2, sub_total, window_title=prog_wt
        )
        self._process_events_light()
        self._upsert_summary_row_at(gr, plab_summary, vals)
        self._show_run_progress(
            _MASTER_DEBUG_PROGRESS_PHASES[2], 3, sub_total, window_title=prog_wt
        )
        self._process_events_light()
        self._upsert_value_cols_at(gr, colvals, [None] * len(colvals))
        if self._scenario_for_dry_run and self._debug_scan_paths and colvals:
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
        self._master_step_idx += 1
        self._last_master_active_count = len(self._active_slot_indices)
        will_finish_master_item = self._master_step_idx >= len(self._active_slot_indices)
        should_finalize_now = will_finish_master_item
        # 条件ステップ表の selectRow は _rebuild_left_steps 末尾で行われる。
        # グリッド再構成（本番同等プレビュー）より先に呼ぶと、重い処理の前に次行へフォーカスが移ってしまうため、順序をグリッド→左ステップにする。
        self._master_progress_window_title = prog_wt
        self._refresh_master_value_grid(finalize=should_finalize_now)
        self._rebuild_left_steps()
        self._show_run_progress(
            _MASTER_DEBUG_PROGRESS_PHASES[7], 8, sub_total, window_title=prog_wt
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
                if self._active_slot_indices:
                    self._rebuild_value_grid()
                else:
                    try:
                        _data_agg_probe_log.info(
                            "[DATA_AGG_DIAG] value_grid_keep reason=no_active_slots_after_item_done mi_idx=%s",
                            self._mi_idx,
                        )
                    except Exception:
                        pass
        self._close_run_progress()
        return True, focus_results_tab

    def _finish_continuous_run(self) -> None:
        if not self._continuous_busy:
            return
        # マスタ連続実行完了時: 左が「シナリオなし」項目にいても、
        # 直前までの結果一覧を最終反映する（マージ列表示を 1 回）。
        if self._mode == 1:
            fb_mi = getattr(self, "_last_master_completed_mi_idx", None)
            self._mpv_display_mi_idx = fb_mi if fb_mi is not None else self._mi_idx
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
            # 最終表示は progress 行（本番同等パイプライン結果）を優先して差分見え崩れを防ぐ。
            self._mpv_show_merged_current = False
            self._mpv_join_table_active = False
            self._mpv_join_table_ncols = 0
            self._rebuild_value_grid()
            self._paint_result_highlights()
        self._continuous_busy = False
        self._continuous_steps_left = 0
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
        if self._continuous_steps_left <= 0:
            self._finish_continuous_run()
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        ok = False
        try:
            ok, focus_tab = self._execute_single_run_step()
            if ok:
                if focus_tab:
                    self._focus_results_tab()
                    self._schedule_focus_results_tab()
                self._update_run_buttons_state()
                self._update_clear_buttons()
        finally:
            QApplication.restoreOverrideCursor()
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
            )
        else:
            QTimer.singleShot(0, self._run_continuous_next)

    def _on_run_all_continuous(self) -> None:
        if self._continuous_busy:
            return
        if self._mode == 1:
            self._continuous_was_full_master = False
            self._master_full_continuous_allowed = False
            self._master_step_pass_complete = False
        if self._mode == 0:
            n = len(self._active_slot_indices)
        else:
            n = max(0, len(self._active_slot_indices) - self._master_step_idx)
        if n <= 0:
            return
        self._continuous_busy = True
        self._continuous_steps_left = n
        self._continuous_initial_steps = n
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
        ok = False
        focus_tab = False
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
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
        finally:
            if self._mode == 1:
                self._log_master_exec_unit_close(
                    "ステップ実行", "完了" if ok else "中断"
                )
            if self._mode == 1:
                QTimer.singleShot(350, self._close_run_progress)
            QApplication.restoreOverrideCursor()

    def _clear_master_debug_results(self) -> None:
        self._master_full_continuous_allowed = True
        self._master_step_pass_complete = False
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
        self._mpv_progress_rows_step_cache.clear()
        self._mpv_progress_rows_by_mi.clear()
        self._mpv_frozen_snapshots.clear()
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
