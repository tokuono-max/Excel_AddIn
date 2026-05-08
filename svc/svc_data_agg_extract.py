# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_data_agg_extract.py
Created: 2026-03-18
Updated: 2026-03-18
Version: 0.1.0
Purpose:
  データ集約用の抽出エンジン。座標（絶対セル）・メタデータ（パス・フォルダ名・ファイル名）・
  ファイル名からの文字列抽出（範囲・デリミタ・正規表現）を提供する。OpenPyXL / csv で Excel/CSV を直接読む。
  svc_data_agg から呼び出され、サブモジュールとして分離する。
History (latest 3):
  - 0.1.6 (2026-04-07) extract_item_values: sources 空はファイル名ではなく主値 [""]（未設定列を一括出力で空にする）。
  - 0.1.5 (2026-04-04) read_only シートを iter_rows で一度具体化し ws[ref] の都度走査を避ける（スコープ内・繰り返し抽出）。
  - 0.1.4 (2026-04-04) DATA_AGG_PER_FILE_TIMING=1 時 load_workbook 所要をスレッドローカルに集計し consume で取得。
  - 0.1.7 (2026-04-14) extract_item_bundle(link/join): deepcopy を浅いコピーに変更（主値・連携は共有、再計算分のみ新 dict）。
  - 0.1.3 (2026-04-01) xlsx_workbook_scope: compute_batch 等で同一 .xlsx の load_workbook をファイル単位で再利用。
  - 0.1.2 (2026-03-28) 主キー・連携の UI 正規表現を廃止。連携は checks→value_shape_script。link/join 抽出後 postprocess_link_rule_value。
  - 0.1.1 (2026-03-28) extract_item_values: 主値に正規表現・チェック・value_shape_script（core.core_value_shape）を適用。
  - 0.1.0 (2026-03-18) 新規作成。extract_metadata / extract_from_filename / extract_cell / extract_item_values。
"""
from __future__ import annotations

import csv
import importlib
import os
import re
import sys
import threading
import time
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Optional

_path_svc = Path(__file__).resolve().parent
_root = _path_svc.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_log import get_logger  # noqa: E402
from svc.data_agg_source_ui import source_ui_block  # noqa: E402
from svc.data_agg_value_post import (  # noqa: E402
    postprocess_cell_primary,
    postprocess_link_rule_value,
    postprocess_metadata_like_primary,
    postprocess_name_extract_primary,
)

logger = get_logger(__name__)
__version__ = "0.1.5"

# compute_batch 等で同一 .xlsx を複数項目・複数セル参照するときの重複 load_workbook を避ける。
_tls_wb_scope = threading.local()
# DATA_AGG_PER_FILE_TIMING=1: load_workbook の秒をパスキー別に累積（consume で取り出し）
_tls_wb_open_sec = threading.local()


def _per_file_workbook_timing_enabled() -> bool:
    from core import core_env

    return core_env.data_agg_file_timing_enabled()


def _add_workbook_open_seconds(path_key: str, seconds: float) -> None:
    if seconds <= 0 or not _per_file_workbook_timing_enabled():
        return
    d = getattr(_tls_wb_open_sec, "by_path", None)
    if d is None:
        d = {}
        _tls_wb_open_sec.by_path = d
    d[path_key] = float(d.get(path_key, 0.0)) + float(seconds)


def consume_workbook_open_ms_for_path(file_path: str) -> int:
    """
    当該ファイルパスについてスコープ内で計測した openpyxl.load_workbook 合計を ms で返し、内部を消す。
    .xlsx 以外や未ロード時は 0。
    """
    if not _per_file_workbook_timing_enabled():
        return 0
    d = getattr(_tls_wb_open_sec, "by_path", None)
    if not d:
        return 0
    try:
        key = str(Path(file_path).resolve())
    except Exception:
        key = str(file_path)
    sec = float(d.pop(key, 0.0))
    return int(sec * 1000.0 + 0.5)


@contextmanager
def xlsx_workbook_scope() -> Iterator[None]:
    """
    入れ子可。スレッドローカル。同一スコープ内では resolve 済みパス文字列キーで read_only Workbook を1つだけ保持し、
    終了時にまとめて close する。スコープ外では従来どおり extract_cell ごとに開閉する。
    各フレームは wbs（パス→Workbook）と sheet_mats（(パスキー, シート名)→行行列）を持つ。
    """
    stack: list[dict[str, Any]] = getattr(_tls_wb_scope, "stack", None)
    if stack is None:
        stack = []
        _tls_wb_scope.stack = stack
    frame: dict[str, Any] = {"wbs": {}, "sheet_mats": {}}
    stack.append(frame)
    try:
        yield
    finally:
        stack.pop()
        for wb in frame["wbs"].values():
            try:
                wb.close()
            except Exception:
                pass


def _xlsx_workbook_cache_top() -> Optional[dict[str, Any]]:
    stack = getattr(_tls_wb_scope, "stack", None)
    if not stack:
        return None
    return stack[-1]


def _xlsx_workbook_from_cache(path: Path) -> Optional[Any]:
    """スコープ内ならキャッシュから取得または load して登録。スコープ外は None。"""
    frame = _xlsx_workbook_cache_top()
    if frame is None:
        return None
    if path.suffix.lower() != ".xlsx":
        return None
    wbs: dict[str, Any] = frame.setdefault("wbs", {})
    key = str(path.resolve())
    wb = wbs.get(key)
    if wb is not None:
        return wb
    try:
        import openpyxl  # noqa: E402
    except ImportError:
        return None
    try:
        if _per_file_workbook_timing_enabled():
            t_ld = time.perf_counter()
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            _add_workbook_open_seconds(key, time.perf_counter() - t_ld)
        else:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        logger.debug("[DATA_AGG_EXTRACT] Excel 読込エラー %s: %s", path, e)
        return None
    wbs[key] = wb
    return wb

_POLARS_MODULE: Any | None = None
_POLARS_CHECKED = False


def _get_polars() -> Any | None:
    global _POLARS_MODULE, _POLARS_CHECKED
    if _POLARS_CHECKED:
        return _POLARS_MODULE
    try:
        _POLARS_MODULE = importlib.import_module("polars")
    except Exception:
        _POLARS_MODULE = None
    _POLARS_CHECKED = True
    return _POLARS_MODULE


def _has_polars() -> bool:
    return _get_polars() is not None


def source_passes_file_name_filter(file_path: str | Path, src: dict[str, Any]) -> bool:
    """
    セル系ソースの UI 保存ブロック（`ui_scenario_source_v1`、レガシーキーは source_ui_block で吸収）内の file_pattern / file_name_rule を評価する。
    file_pattern が空ならフィルタなし（True）。大小文字は区別しない（§10.3.3）。
    """
    stype = (src.get("type") or "cell").strip().lower()
    if stype != "cell":
        return True
    block = source_ui_block(src)
    if not isinstance(block, dict):
        return True
    p = Path(file_path)
    stem_l = p.stem.lower()
    ext_l = p.suffix.lower()
    ext_checked = block.get("ext_checked")
    if isinstance(ext_checked, list) and ext_checked:
        norm_exts = {
            (str(e).strip().lower() if str(e).strip().startswith(".") else "." + str(e).strip().lower())
            for e in ext_checked
            if str(e).strip()
        }
        if norm_exts and ext_l not in norm_exts:
            return False
    pattern = str(block.get("file_pattern") or "").strip()
    if not pattern:
        return True
    rule = str(block.get("file_name_rule") or "含む").strip()
    pat_l = pattern.lower()
    if "完全一致" in rule or rule.lower() in ("exact", "equals"):
        return stem_l == pat_l
    if "含まない" in rule or rule.lower() in ("exclude", "not_contains"):
        return pat_l not in stem_l
    return pat_l in stem_l


# メタデータのソース種別（後方互換用）
SOURCE_FULL_PATH = "full_path"
SOURCE_DIR_NAME = "dir_name"
SOURCE_FILE_NAME = "file_name"

# 名前から抽出の開始モード
START_MODE_HEAD = "head"
START_MODE_DELIMITER = "delimiter"
START_MODE_POSITION = "position"

# 名前から抽出の長さモード
LENGTH_MODE_CHAR = "char"
LENGTH_MODE_COUNT = "count"
LENGTH_MODE_END = "end"


def extract_metadata(
    file_path: str | Path,
    source_type: str = SOURCE_FILE_NAME,
) -> str:
    """
    ファイルパスからメタデータ（フルパス・フォルダ名・ファイル名）を抽出する。

    【概要】
      指定した source_type に応じて、フルパス・ディレクトリ名・ファイル名のいずれかを返す。

    【引数】
      file_path: 対象ファイルのパス。
      source_type: "full_path" | "dir_name" | "file_name"。

    【戻り値】
      抽出した文字列。取得できない場合は空文字。
    """
    from svc.data_agg_path_norm import normalize_source_path  # noqa: E402

    p = Path(file_path).resolve()
    if source_type == SOURCE_FULL_PATH:
        return normalize_source_path(p)
    if source_type == SOURCE_DIR_NAME:
        return p.parent.name if p.parent else ""
    if source_type == SOURCE_FILE_NAME:
        return p.name
    return ""


def _extract_from_string(
    text: str,
    start: int = 0,
    length: Optional[int] = None,
    delimiter: Optional[str] = None,
    part_index: Optional[int] = None,
    pattern: Optional[str] = None,
) -> str:
    """
    文字列に範囲・デリミタ・正規表現の抽出ルールを適用する。
    extract_from_filename およびメタデータの抽出ルールで共通利用。
    """
    if pattern:
        m = re.search(pattern, text)
        if m:
            return m.group(1) if m.lastindex and m.lastindex >= 1 else m.group(0)
        return ""
    if delimiter is not None and part_index is not None:
        parts = text.split(delimiter)
        if 0 <= part_index < len(parts):
            return parts[part_index]
        return ""
    if length is not None:
        return text[start : start + length]
    return text[start:] if start < len(text) else ""


def extract_from_filename(
    file_path: str | Path,
    start: int = 0,
    length: Optional[int] = None,
    delimiter: Optional[str] = None,
    part_index: Optional[int] = None,
    pattern: Optional[str] = None,
) -> str:
    """
    ファイル名（またはパス）から、範囲指定・デリミタ分割・正規表現で文字列を抽出する。

    【概要】
      length 指定時は name[start:start+length]。delimiter と part_index 指定時は分割して第 n 要素。
      pattern 指定時は正規表現の最初のグループを返す（グループが無い場合はマッチ全体）。

    【引数】
      file_path: 対象ファイルのパス。
      start: 開始文字位置（0 始まり）。
      length: 取得文字数。None のときは末尾まで。
      delimiter: 区切り文字。指定時は name を分割する。
      part_index: 分割後のインデックス（0 始まり）。delimiter と併用。
      pattern: 正規表現パターン。マッチの第1グループまたはマッチ全体を返す。

    【戻り値】
      抽出した文字列。
    """
    p = Path(file_path).resolve()
    return _extract_from_string(
        p.name, start, length, delimiter, part_index, pattern
    )


def extract_from_name(
    file_path: str | Path,
    source_type: str = "file_name",
    search_condition: Optional[str] = None,
    search_text: Optional[str] = None,
    search_keywords: Optional[list[str]] = None,  # deprecated
    keyword_logic: Optional[str] = None,  # deprecated
    start_mode: str = "head",
    start_value: Any = 0,
    length_mode: str = "end",
    length_value: Any = None,
    delimiter: Optional[str] = None,
    part_index: Optional[int] = None,
    pattern: Optional[str] = None,
    replacement: Optional[str] = None,
) -> str:
    """
    ファイル名またはフォルダ名から、開始位置・抽出長さ・正規表現加工で文字列を抽出する。

    【概要】
      source_type で file_name / dir_name を選択。
      search_condition と search_text で抽出前のフィルタ（含む/含まない）を適用。
      start_mode: head（先頭）→ length_mode で抽出長さを指定。
      start_mode: delimiter → 区切記号＋part_index で分割し、該当部分を取得（抽出長さは不要）。
      start_mode: position → start_value を開始位置として length_mode で抽出。
      抽出後に pattern と replacement で re.sub による加工を適用。

    【引数】
      file_path: 対象ファイルのパス。
      source_type: "file_name" | "dir_name"
      search_condition: "include"（含む）| "exclude"（含まない）| None（フィルタなし）
      search_text: 検索文字列。空の場合はフィルタなし。
      search_keywords: 互換用（未使用）。
      keyword_logic: 互換用（未使用）。
      start_mode: "head" | "delimiter" | "position"
      start_value: position 時の開始位置（UI は 1 始まり＝先頭文字を 1。内部で 0 始まりに変換）。
      length_mode: "char" | "count" | "end"
        - char: length_value を終端文字として、その文字を含めて抽出。
        - count: length_value 文字数だけ抽出。
        - end: 末尾まで。
      length_value: char/count 時の値。
      delimiter: 区切記号（delimiter モード時）。
      part_index: 分割後のインデックス（UI は 1 始まり。内部で 0 始まりに変換）。
      pattern: 抽出後の正規表現パターン（re.sub）。
      replacement: 置換文字列。

    【戻り値】
      抽出・加工した文字列。
    """
    p = Path(file_path).resolve()
    st = str(source_type or "file_name").strip().lower()
    text = _name_extract_text_for_match_and_extract(p, st)

    # 検索条件フィルタ（単一文字列）: 条件を満たさない場合は空文字を返す（name_extract_search_matches と整合）
    kw = (search_text or "").strip()
    if kw and search_condition:
        t_l = text.lower()
        k_l = kw.lower()
        sc = str(search_condition).strip().lower()
        if sc in ("exclude", "含まない"):
            if k_l in t_l:
                return ""
        elif sc in ("exact", "equals", "完全一致"):
            if k_l != t_l:
                return ""
        else:
            if k_l not in t_l:
                return ""

    def _apply_regex(t: str) -> str:
        if pattern:
            try:
                # 置換欄はUI非表示。replacement 未指定時は空文字置換で整形する。
                return re.sub(pattern, replacement if replacement is not None else "", t)
            except re.error:
                pass
        return t

    if start_mode == START_MODE_DELIMITER:
        if delimiter is not None and str(delimiter) != "" and part_index is not None:
            parts = text.split(str(delimiter))
            try:
                idx0 = int(part_index)
            except (TypeError, ValueError):
                idx0 = 1
            idx0 = max(0, idx0 - 1)
            extracted = parts[idx0] if 0 <= idx0 < len(parts) else ""
        else:
            extracted = text
        return _apply_regex(extracted)

    if start_mode == START_MODE_HEAD:
        start = 0
    elif start_mode == START_MODE_POSITION:
        try:
            sv = int(start_value) if start_value is not None else 1
        except (TypeError, ValueError):
            sv = 1
        start = max(0, sv - 1)
    else:
        start = 0

    if start >= len(text):
        return _apply_regex("")

    if length_mode == LENGTH_MODE_CHAR:
        end_char = str(length_value) if length_value is not None else ""
        if end_char:
            pos = text.find(end_char, start)
            if pos >= 0:
                hi = min(len(text), pos + len(end_char))
                extracted = text[start:hi]
            else:
                extracted = text[start:]
        else:
            extracted = text[start:]
    elif length_mode == LENGTH_MODE_COUNT:
        cnt = int(length_value) if length_value is not None else None
        extracted = (
            text[start : start + cnt]
            if cnt is not None and cnt > 0
            else text[start:]
        )
    else:
        extracted = text[start:]

    return _apply_regex(extracted)


def _name_extract_text_for_match_and_extract(path: Path, source_type: str) -> str:
    """
    名前取得の検索フィルタ・抽出で共有する対象文字列。
    file_name: 拡張子なし（stem）。dir_name: 親フォルダ名。
    """
    st = str(source_type or "file_name").strip().lower()
    if st == "dir_name":
        return (path.parent.name if path.parent else "") or ""
    return path.stem or ""


def name_extract_search_matches(file_path: str | Path, src: dict[str, Any]) -> bool:
    """
    extract_from_name と同じ対象文字列（dir: 親フォルダ名 / file: stem）で
    search_text・search_condition を評価する。キーワード空または条件未設定はフィルタ無し。
    """
    kw = (src.get("search_text") or "").strip()
    if not kw or not src.get("search_condition"):
        return True
    p = Path(file_path).resolve()
    st = str(src.get("source_type") or "file_name").strip().lower()
    text = _name_extract_text_for_match_and_extract(p, st).lower()
    kw_l = kw.lower()
    hit_sub = kw_l in text
    hit_eq = kw_l == text
    sc = str(src.get("search_condition") or "include").strip().lower()
    if sc in ("exclude", "含まない"):
        return not hit_sub
    if sc in ("exact", "equals", "完全一致"):
        return hit_eq
    return hit_sub


def name_extract_join_comparison_key(file_path: str | Path, src: dict[str, Any]) -> str:
    """
    名前取得の結合パス照合キー（正規化済み）。file_name: ファイルの正規化フルパス。
    dir_name: 親ディレクトリの正規化パス。svc_data_agg._apply_name_extract_path_assignment と一致させる。
    """
    from svc.data_agg_path_norm import normalize_source_path

    p = Path(file_path).resolve()
    st = str(src.get("source_type") or "file_name").strip().lower()
    if st == "dir_name":
        return normalize_source_path(p.parent)
    return normalize_source_path(p)


def name_extract_unique_join_keys_ordered(allowed_paths: list[str], src: dict[str, Any]) -> list[str]:
    """検索一致ファイルについて、照合キーを allowed の出現順で重複なく列挙。"""
    seen: set[str] = set()
    out: list[str] = []
    for fp in allowed_paths:
        if not name_extract_search_matches(fp, src):
            continue
        key = name_extract_join_comparison_key(fp, src)
        if key not in seen:
            seen.add(key)
            out.append(key)
    return out


def name_extract_hit_files_ordered(allowed_paths: list[str], src: dict[str, Any]) -> list[str]:
    """検索一致したファイルパスを allowed の順で列挙（同一フォルダ内の複数ファイルは別行のまま）。"""
    return [fp for fp in allowed_paths if name_extract_search_matches(fp, src)]


def extract_cells_repeat(
    file_path: str | Path,
    sheet_name: Optional[str] = None,
    cell_ref: str = "A1",
    repeat_direction: str = "vertical",
    repeat_until_empty: bool = True,
    repeat_max: Optional[int] = None,
) -> list[Any]:
    """
    開始セルから縦または横に繰り返しセルを取得する。

    【引数】
      repeat_direction: "vertical"（下方向）| "horizontal"（右方向）
      repeat_until_empty: True で空セルで終了
      repeat_max: 最大取得数。None で制限なし。

    【戻り値】
      取得した値のリスト。
    """
    p = Path(file_path).resolve()
    if not p.is_file():
        return []
    if p.suffix.lower() == ".csv":
        pl_mod = _get_polars()
        if pl_mod is not None:
            try:
                df = pl_mod.read_csv(str(p), has_header=False, encoding="utf8-lossy")
                col, row = _parse_cell_ref(cell_ref)
                if col is None or row is None:
                    return []
                out: list[Any] = []
                dc = 1 if repeat_direction == "horizontal" else 0
                dr = 1 if repeat_direction == "vertical" else 0
                for _ in range(repeat_max or 9999):
                    if row >= df.height or col >= df.width or row < 0 or col < 0:
                        break
                    v = df.row(row)[col]
                    if repeat_until_empty and (v is None or v == ""):
                        break
                    out.append(v)
                    col += dc
                    row += dr
                    if repeat_max and len(out) >= repeat_max:
                        break
                return out
            except Exception:
                pass
    col, row = _parse_cell_ref(cell_ref)
    if col is None or row is None:
        return []
    results: list[Any] = []
    delta_col = 1 if repeat_direction == "horizontal" else 0
    delta_row = 1 if repeat_direction == "vertical" else 0
    for _ in range(repeat_max or 9999):
        cr = _col_row_to_cell_ref(col, row)
        v = extract_cell(p, sheet_name, cr)
        if repeat_until_empty and (v is None or v == ""):
            break
        results.append(v)
        col += delta_col
        row += delta_row
        if repeat_max and len(results) >= repeat_max:
            break
    return results


def extract_cell(
    file_path: str | Path,
    sheet_name: Optional[str] = None,
    cell_ref: str = "A1",
) -> Any:
    """
    指定ファイルのセル値を取得する。Excel (.xlsx) は OpenPyXL、CSV は csv モジュールで読む。

    【概要】
      絶対位置のセル参照（例: "B5"）で値を返す。CSV の場合は sheet_name を無視し、先頭行を 1 として行・列で解釈する
      （cell_ref は "A1" 形式のみ対応。列 A=0, 行 1=0 のインデックスでアクセス）。

    【引数】
      file_path: 対象ファイルのパス。
      sheet_name: シート名。None のときは先頭シート（Excel）または唯一のシート（CSV）。
      cell_ref: セル番地（例: "A1", "B5"）。CSV では A=列0, 1=行0。

    【戻り値】
      セルの値。取得できない場合は None。
    """
    p = Path(file_path).resolve()
    if not p.is_file():
        return None
    suf = p.suffix.lower()
    if suf == ".csv":
        return _get_csv_cell(p, cell_ref)
    if suf in (".xlsx", ".xls"):
        return _get_excel_cell(p, sheet_name, cell_ref)
    return None


def _get_csv_cell(path: Path, cell_ref: str) -> Any:
    """CSV ファイルから A1 形式のセル参照で値を取得する。"""
    col, row = _parse_cell_ref(cell_ref)
    if col is None or row is None:
        return None
    try:
        pl_mod = _get_polars()
        if pl_mod is not None:
            df = pl_mod.read_csv(str(path), has_header=False, encoding="utf8-lossy")
            if 0 <= row < df.height and 0 <= col < df.width:
                return df.row(row)[col]
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = list(csv.reader(f))
        if row < len(reader) and col < len(reader[row]):
            return reader[row][col]
    except Exception as e:
        logger.debug("[DATA_AGG_EXTRACT] CSV 読込エラー %s: %s", path, e)
    return None


def _col_row_to_cell_ref(col: int, row: int) -> str:
    """0始まりの (列, 行) を A1 形式に変換。"""
    if col < 0 or row < 0:
        return "A1"
    col_letters = ""
    c = col + 1
    while c > 0:
        c, r = divmod(c - 1, 26)
        col_letters = chr(ord("A") + r) + col_letters
    return col_letters + str(row + 1)


def _parse_cell_ref(cell_ref: str) -> tuple[Optional[int], Optional[int]]:
    """A1 形式を (列インデックス, 行インデックス) に変換。行・列は 0 始まり。"""
    s = (cell_ref or "").strip().upper()
    if not s:
        return (None, None)
    col_str = ""
    i = 0
    while i < len(s) and s[i].isalpha():
        col_str += s[i]
        i += 1
    row_str = s[i:] if i < len(s) else ""
    if not col_str or not row_str:
        return (None, None)
    try:
        row = int(row_str) - 1  # 1-based to 0-based
    except ValueError:
        return (None, None)
    col = 0
    for c in col_str:
        col = col * 26 + (ord(c) - ord("A") + 1)
    col -= 1  # 1-based to 0-based
    return (col, row)


def _get_excel_cell(
    path: Path,
    sheet_name: Optional[str],
    cell_ref: str,
) -> Any:
    """Excel ファイル（.xlsx）からセル値を取得する。.xls は未対応の場合は None。"""
    if path.suffix.lower() == ".xls":
        logger.debug("[DATA_AGG_EXTRACT] .xls は未対応: %s", path)
        return None
    try:
        import openpyxl  # noqa: E402
    except ImportError:
        logger.warning("[DATA_AGG_EXTRACT] openpyxl が利用できません")
        return None
    wb_cached = _xlsx_workbook_from_cache(path)
    if wb_cached is not None:
        try:
            return _xlsx_cell_value_open_workbook(
                wb_cached,
                sheet_name,
                cell_ref,
                path=path,
            )
        except Exception:
            return None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                ws = wb.active
            else:
                ws = wb[sheet_name]
        else:
            ws = wb.active
        if ws is None:
            return None
        val = ws[cell_ref].value
        wb.close()
        return val
    except Exception as e:
        logger.debug("[DATA_AGG_EXTRACT] Excel 読込エラー %s: %s", path, e)
    return None


def _resolve_readonly_worksheet(wb: Any, sheet_name: Optional[str]) -> tuple[Any, str]:
    """read_only Workbook から対象シートと実効シート名を返す。"""
    names = getattr(wb, "sheetnames", None) or []
    if sheet_name and sheet_name in names:
        return wb[sheet_name], sheet_name
    ws = wb.active
    title = getattr(ws, "title", "") if ws is not None else ""
    return ws, str(title)


def _materialize_readonly_sheet_matrix(ws: Any) -> list[list[Any]]:
    """read_only シートを一度だけ走査し values の行リストにする（以降 ws[ref] は使わない）。"""
    if ws is None:
        return []
    rows: list[list[Any]] = []
    for tup in ws.iter_rows(values_only=True):
        rows.append(list(tup))
    return rows


def _matrix_cell_value(matrix: list[list[Any]], col: int, row: int) -> Any:
    if row < 0 or col < 0:
        return None
    if row >= len(matrix):
        return None
    r = matrix[row]
    if col >= len(r):
        return None
    return r[col]


def _xlsx_cell_value_open_workbook(
    wb: Any,
    sheet_name: Optional[str],
    cell_ref: str,
    *,
    path: Optional[Path] = None,
    ephemeral_sheet_cache: bool = False,
) -> Any:
    """
    既に開いた read_only Workbook から 1 セル読み。
    xlsx_workbook_scope 内、または ephemeral_sheet_cache=True（同一 wb での繰り返し）では
    シートを iter_rows で一度具体化し、メモリ参照に切り替える。
    """
    col, row = _parse_cell_ref(cell_ref)
    if col is None or row is None:
        return None

    frame = _xlsx_workbook_cache_top()
    use_mat = frame is not None or ephemeral_sheet_cache
    if not use_mat:
        try:
            ws, _ = _resolve_readonly_worksheet(wb, sheet_name)
            if ws is None:
                return None
            return ws[cell_ref].value
        except Exception:
            return None

    ws, resolved_name = _resolve_readonly_worksheet(wb, sheet_name)
    if ws is None:
        return None

    path_key = str(path.resolve()) if path is not None else f"id:{id(wb)}"
    mat: Optional[list[list[Any]]] = None
    mats_store: Optional[dict[Any, list[list[Any]]]] = None
    store_key: Any = None

    if frame is not None:
        mats_store = frame.setdefault("sheet_mats", {})
        store_key = (path_key, resolved_name)
        mat = mats_store.get(store_key)
    elif ephemeral_sheet_cache:
        mats_store = getattr(wb, "_da_sheet_mats", None)
        if mats_store is None:
            mats_store = {}
            setattr(wb, "_da_sheet_mats", mats_store)
        store_key = resolved_name
        mat = mats_store.get(store_key)

    if mat is None:
        try:
            mat = _materialize_readonly_sheet_matrix(ws)
        except Exception:
            return None
        if mats_store is not None and store_key is not None:
            mats_store[store_key] = mat

    return _matrix_cell_value(mat, col, row)


def extract_item_values(
    file_path: str | Path,
    item_config: dict[str, Any],
    item_id: Optional[str] = None,
    cell_positions: Optional[dict[str, tuple[int, int]]] = None,
    max_primary_rows: Optional[int] = None,
    *,
    cell_source_spans_out: Optional[dict[int, tuple[int, int]]] = None,
) -> list[Any]:
    """
    1 項目分の設定に基づき、指定ファイルから抽出した値のリストを返す。

    【概要】
      item_config の sources を走査し、座標・メタデータ・ファイル名の各ソースから値を集める。
      sources が空のときは [""] のみを返す（ファイル名での主値フォールバックは行わない）。
      複数ソースがある場合は順に取得し、空でない値をリストに追加する。
      単一セル・単一メタデータの場合は要素 1 のリストを返す。
      セル座標の相対指定（anchor + row_offset, col_offset）に対応。

    【引数】
      file_path: 対象ファイルのパス。
      item_config: 項目設定（sources: [ { type, sheet_name?, cell_ref?, anchor?, row_offset?, col_offset? } ] 等）。
      item_id: 項目ID。セル位置の記録に使用。
      cell_positions: 項目ID→(col, row) の辞書。相対指定の基準と、抽出したセル位置の記録に使用。

    【戻り値】
      抽出した値のリスト。1 ステップで 1 項目の「列」に書き込む値の並び。

    cell_source_spans_out:
      指定時、セル系ソースごとに (結合主値リスト上の開始インデックス, 行数) を格納する。
    """
    results: list[Any] = []
    sources = item_config.get("sources") or []
    positions = cell_positions if cell_positions is not None else {}
    if not sources:
        # ソース未設定列は主値を載せない（ファイル名フォールバックは一括で誤出力になるため）
        return [""]
    for si, src in enumerate(sources):
        if not isinstance(src, dict):
            continue
        stype = (src.get("type") or "cell").strip().lower()
        if stype == "name_extract":
            if not name_extract_search_matches(file_path, src):
                results.append(None)
                continue
            ui_blk = source_ui_block(src)
            ex_mode = str((ui_blk or {}).get("extract_mode") or "extract").strip().lower()
            if ex_mode == "fixed":
                raw_f = src.get("length_value")
                v_fix = "" if raw_f is None else str(raw_f).strip()
                if v_fix:
                    results.append(
                        postprocess_name_extract_primary(
                            v_fix, ui_blk if isinstance(ui_blk, dict) else None
                        )
                    )
                else:
                    results.append(None)
                continue
            v = extract_from_name(
                file_path,
                source_type=src.get("source_type") or "file_name",
                search_condition=src.get("search_condition"),
                search_text=src.get("search_text"),
                start_mode=src.get("start_mode") or START_MODE_HEAD,
                start_value=src.get("start_value"),
                length_mode=src.get("length_mode") or LENGTH_MODE_END,
                length_value=src.get("length_value"),
                delimiter=src.get("delimiter"),
                part_index=int(x) if (x := src.get("part_index")) is not None else None,
                pattern=src.get("pattern"),
                replacement=src.get("replacement"),
            )
            if v is not None:
                results.append(
                    postprocess_name_extract_primary(v, ui_blk if isinstance(ui_blk, dict) else None)
                )
        elif stype == "metadata" or stype == "meta":
            meta_type = src.get("source_type") or SOURCE_FILE_NAME
            raw = extract_metadata(file_path, meta_type)
            start = src.get("start")
            length = src.get("length")
            delimiter = src.get("delimiter")
            part_index = src.get("part_index")
            pattern = src.get("pattern")
            if (
                start is not None
                or length is not None
                or delimiter is not None
                or part_index is not None
                or pattern
            ):
                ln_val = int(length) if length is not None else None
                if ln_val is not None and ln_val <= 0:
                    ln_val = None
                v = _extract_from_string(
                    raw,
                    start=int(start) if start is not None else 0,
                    length=ln_val,
                    delimiter=delimiter,
                    part_index=int(part_index) if part_index is not None else None,
                    pattern=pattern,
                )
            else:
                v = raw
            if v is not None:
                ui_blk = source_ui_block(src)
                results.append(
                    postprocess_metadata_like_primary(v, ui_blk if isinstance(ui_blk, dict) else None)
                )
        elif stype == "filename":
            v = extract_from_filename(
                file_path,
                start=src.get("start", 0),
                length=src.get("length"),
                delimiter=src.get("delimiter"),
                part_index=src.get("part_index"),
                pattern=src.get("pattern"),
            )
            ui_blk = source_ui_block(src)
            results.append(
                postprocess_metadata_like_primary(v, ui_blk if isinstance(ui_blk, dict) else None)
            )
        else:
            # cell / coordinate（絶対 or 相対）
            if not source_passes_file_name_filter(file_path, src):
                continue
            cell_start = len(results)
            sheet_name = src.get("sheet_name")
            cell_ref = src.get("cell_ref") or "A1"
            anchor = src.get("anchor")
            row_off = int(src.get("row_offset") or 0)
            col_off = int(src.get("col_offset") or 0)
            if anchor and anchor in positions:
                base_col, base_row = positions[anchor]
                col = base_col + col_off
                row = base_row + row_off
                cell_ref = _col_row_to_cell_ref(col, row)
            ui_blk = source_ui_block(src)
            _blk = ui_blk if isinstance(ui_blk, dict) else None
            repeat_dir = (src.get("repeat_direction") or "").strip().lower()
            if repeat_dir in ("vertical", "horizontal"):
                repeat_until_empty = bool(src.get("repeat_until_empty", True))
                repeat_max = int(x) if (x := src.get("repeat_max")) is not None else None
                limit = repeat_max if (repeat_max is not None and repeat_max > 0) else 9999
                if max_primary_rows is not None and max_primary_rows > 0:
                    limit = min(limit, max_primary_rows)
                vals: list[Any] = []
                p_abs = Path(file_path).resolve()
                wb_ctx: Any = None
                wb_owned = False
                try:
                    if p_abs.suffix.lower() == ".xlsx":
                        wb_ctx = _xlsx_workbook_from_cache(p_abs)
                        if wb_ctx is None:
                            try:
                                import openpyxl  # noqa: E402

                                wb_ctx = openpyxl.load_workbook(
                                    p_abs, read_only=True, data_only=True
                                )
                                wb_owned = True
                            except Exception:
                                wb_ctx = None
                    # 取得座標 = 基準セル + (行/列オフセット * N)（N は 0 始まり）
                    for n in range(limit):
                        cell_ref_n = _resolve_cell_with_offset(
                            cell_ref, row_off * n, col_off * n
                        )
                        if wb_ctx is not None:
                            v = _xlsx_cell_value_open_workbook(
                                wb_ctx,
                                sheet_name,
                                cell_ref_n,
                                path=p_abs,
                                ephemeral_sheet_cache=wb_owned,
                            )
                        else:
                            v = extract_cell(
                                file_path,
                                sheet_name=sheet_name,
                                cell_ref=cell_ref_n,
                            )
                        if repeat_until_empty and (v is None or v == ""):
                            break
                        vals.append(v)
                finally:
                    if wb_owned and wb_ctx is not None:
                        try:
                            wb_ctx.close()
                        except Exception:
                            pass
                for v in vals:
                    if v is not None and (v != "" or src.get("allow_empty")):
                        results.append(postprocess_cell_primary(v, _blk))
                        if (
                            max_primary_rows is not None
                            and max_primary_rows > 0
                            and len(results) >= max_primary_rows
                        ):
                            break
                if vals and item_id:
                    n_last = len(vals) - 1
                    cell_ref_last = _resolve_cell_with_offset(
                        cell_ref,
                        row_off * n_last,
                        col_off * n_last,
                    )
                    c0, r0 = _parse_cell_ref(cell_ref_last)
                    if c0 is not None and r0 is not None:
                        positions[item_id] = (c0, r0)
            else:
                v = extract_cell(
                    file_path,
                    sheet_name=sheet_name,
                    cell_ref=cell_ref,
                )
                if v is not None and (v != "" or src.get("allow_empty")):
                    results.append(postprocess_cell_primary(v, _blk))
                    col, row = _parse_cell_ref(cell_ref)
                    if col is not None and row is not None and item_id:
                        positions[item_id] = (col, row)
            if cell_source_spans_out is not None:
                cell_source_spans_out[si] = (cell_start, len(results) - cell_start)
        if max_primary_rows is not None and max_primary_rows > 0 and len(results) >= max_primary_rows:
            break
    if max_primary_rows is not None and max_primary_rows > 0:
        results = results[:max_primary_rows]
        if cell_source_spans_out is not None:
            for k in list(cell_source_spans_out.keys()):
                g0, ln = cell_source_spans_out[k]
                if g0 + ln > len(results):
                    cell_source_spans_out[k] = (g0, max(0, len(results) - g0))
    return results if results else [None]


def _resolve_cell_with_offset(base_cell: str, row_off: Any = 0, col_off: Any = 0) -> str:
    """A1 形式セルへ行/列オフセットを適用して A1 形式を返す。"""
    c0, r0 = _parse_cell_ref(base_cell or "A1")
    if c0 is None or r0 is None:
        c0, r0 = 0, 0
    try:
        ro = int(row_off or 0)
    except (TypeError, ValueError):
        ro = 0
    try:
        co = int(col_off or 0)
    except (TypeError, ValueError):
        co = 0
    return _col_row_to_cell_ref(c0 + co, r0 + ro)


def _extract_from_cell_rule(
    file_path: str | Path,
    src: dict[str, Any],
    rule: dict[str, Any],
) -> Any:
    """link_defs / join_defs の 1 ルールから値を抽出する。"""
    mode = str(rule.get("mode") or "セル座標").strip()
    rdict = rule if isinstance(rule, dict) else {}
    if "固定" in mode or mode.lower() in ("fixed", "literal"):
        return postprocess_link_rule_value(rule.get("cell"), rdict)
    sheet_name = src.get("sheet_name")
    base_cell = str(rule.get("cell") or src.get("cell_ref") or "A1")
    # 連携キー/結合キーとも同一の座標解決規約（基準セル + 独自 row/col offset）を使う。
    cell_ref = _resolve_cell_with_offset(base_cell, rule.get("row"), rule.get("col"))
    v = extract_cell(file_path, sheet_name=sheet_name, cell_ref=cell_ref)
    return postprocess_link_rule_value(v, rdict)


def _cell_ref_with_iteration(base_cell: str, repeat_direction: str, iter_index: int) -> str:
    """主キーの反復方向に合わせて基準セルを iter_index 分だけ進める。"""
    c0, r0 = _parse_cell_ref(base_cell or "A1")
    if c0 is None or r0 is None:
        c0, r0 = 0, 0
    rd = (repeat_direction or "").strip().lower()
    if rd == "vertical":
        r0 += max(0, int(iter_index))
    elif rd == "horizontal":
        c0 += max(0, int(iter_index))
    return _col_row_to_cell_ref(c0, r0)


def _extract_from_cell_rule_with_context(
    file_path: str | Path,
    src: dict[str, Any],
    rule: dict[str, Any],
    iter_ctx: dict[str, Any],
) -> Any:
    """
    反復コンテキスト（file_path + base_cell）に基づいて link/join 値を取得する。
    rule の row/col は link/join 側の独自オフセットとして常に適用する。
    シート上の段進みには rule_iter_index（ソース内 0 始まり）を使い、未指定時は iter_index
    （結合主値リスト上の行。後方互換）にフォールバックする。
    """
    mode = str(rule.get("mode") or "セル座標").strip()
    rdict = rule if isinstance(rule, dict) else {}
    if "固定" in mode or mode.lower() in ("fixed", "literal"):
        return postprocess_link_rule_value(rule.get("cell"), rdict)
    sheet_name = src.get("sheet_name")
    try:
        rule_iter = int(iter_ctx.get("rule_iter_index", iter_ctx.get("iter_index", 0)))
    except (TypeError, ValueError):
        rule_iter = 0
    base_cell = str(rule.get("cell") or iter_ctx.get("base_cell") or src.get("cell_ref") or "A1")

    def _to_int(v: Any, default: int = 0) -> int:
        try:
            return int(v)
        except (TypeError, ValueError):
            return default

    # 移動オフセットは 0 開始。rule_iter=0 は 0、1 は offset …（ソース内の何行目か）
    row_step = _to_int(rule.get("row"), 0)
    col_step = _to_int(rule.get("col"), 0)
    row_off = row_step * max(0, rule_iter)
    col_off = col_step * max(0, rule_iter)
    # 取得座標 = 設定座標 + (オフセット * N)（N は 0 開始）
    # ここでは基準セルを二重に進めず、offset*N のみを適用する。
    cell_ref = _resolve_cell_with_offset(base_cell, row_off, col_off)
    v = extract_cell(file_path, sheet_name=sheet_name, cell_ref=cell_ref)
    return postprocess_link_rule_value(v, rdict)


def extract_item_bundle(
    file_path: str | Path,
    item_config: dict[str, Any],
    item_id: Optional[str] = None,
    cell_positions: Optional[dict[str, tuple[int, int]]] = None,
    join_path_header: Optional[str] = None,
    *,
    debug_step_scope: Optional[str] = None,
    existing_bundle: Optional[dict[str, Any]] = None,
    max_primary_rows: Optional[int] = None,
) -> dict[str, Any]:
    """
    1 項目分の抽出結果を主値・連携値・結合キー値に分けて返す。

    debug_step_scope:
      None … 本番相当（主キー＋連携＋結合を一括）。
      "primary" … 主キー（と path_item／正規化フルパス）のみ。デバッグの主キーステップ用。
      "link" … existing_bundle 必須。連携キーだけ再計算してマージ。
      "join" … existing_bundle 必須。結合キーだけ再計算してマージ。

    戻り値:
      {
        "primary_values": list[Any],
        "iteration_contexts": list[dict[str, Any]],
        "link_values": dict[item_name, list[Any]],
        "link_contexts": dict[item_name, list[dict[str, Any]]],
        "join_values": dict[item_name, list[Any]],
        "join_contexts": dict[item_name, list[dict[str, Any]]],
        "path_item_values": dict[item_name, list[Any]],  # 名前系のフルパス関連付け用
        "path_item_contexts": dict[item_name, list[dict[str, Any]]],
      }
    """
    sources = item_config.get("sources") or []

    if debug_step_scope in ("link", "join"):
        if not existing_bundle or not isinstance(existing_bundle, dict):
            return {
                "primary_values": [],
                "iteration_contexts": [],
                "link_values": {},
                "link_contexts": {},
                "join_values": {},
                "join_contexts": {},
                "path_item_values": {},
                "path_item_contexts": {},
            }
        # 浅いコピー: 主値・連携・path_item 等は共有し、再計算する link/join だけ新規 dict（デバッグ段階実行のコスト削減）。
        bundle = existing_bundle.copy()
        if debug_step_scope == "link":
            bundle["link_values"] = {}
            bundle["link_contexts"] = {}
        else:
            bundle["join_values"] = {}
            bundle["join_contexts"] = {}
        span_map = bundle.get("_cell_source_spans")
        legacy_link_iter = not isinstance(span_map, dict)
        n_all = max(1, len(bundle.get("primary_values") or []))
        for si, src in enumerate(sources):
            if not isinstance(src, dict):
                continue
            stype = (src.get("type") or "cell").strip().lower()
            ui_blk = source_ui_block(src)
            if not isinstance(ui_blk, dict):
                continue
            if stype != "cell":
                continue
            if not source_passes_file_name_filter(file_path, src):
                continue
            if legacy_link_iter:
                g_off, n_src = 0, n_all
            else:
                sp = span_map.get(si) if isinstance(span_map, dict) else None
                if not sp or not isinstance(sp, (tuple, list)) or len(sp) != 2:
                    continue
                try:
                    g_off = int(sp[0])
                    n_src = int(sp[1])
                except (TypeError, ValueError):
                    continue
                if n_src < 1:
                    continue
            src_base = str(src.get("cell_ref") or "A1")
            try:
                row_step = int(src.get("row_offset") or 0)
            except (TypeError, ValueError):
                row_step = 0
            try:
                col_step = int(src.get("col_offset") or 0)
            except (TypeError, ValueError):
                col_step = 0
            iter_contexts: list[dict[str, Any]] = []
            for local_i in range(n_src):
                base_cell_i = _resolve_cell_with_offset(
                    src_base,
                    row_step * local_i,
                    col_step * local_i,
                )
                iter_contexts.append(
                    {
                        "file_path": str(file_path),
                        "iter_index": int(g_off + local_i),
                        "rule_iter_index": int(local_i),
                        "base_cell": base_cell_i,
                    }
                )
            b_ictx = bundle.get("iteration_contexts") or []
            for local_i, ic in enumerate(iter_contexts):
                gi = g_off + local_i
                if gi < len(b_ictx) and isinstance(b_ictx[gi], dict):
                    b_ictx[gi]["base_cell"] = ic.get("base_cell")
            if debug_step_scope == "link":
                for ld in ui_blk.get("link_defs") or []:
                    if not isinstance(ld, dict):
                        continue
                    target = str(ld.get("item") or "").strip()
                    if not target:
                        continue
                    for iter_ctx in iter_contexts:
                        v = _extract_from_cell_rule_with_context(file_path, src, ld, iter_ctx)
                        bundle["link_values"].setdefault(target, []).append(v)
                        bundle["link_contexts"].setdefault(target, []).append(
                            {
                                "file_path": str(iter_ctx.get("file_path") or file_path),
                                "iter_index": int(iter_ctx.get("iter_index", 0)),
                                "base_cell": iter_ctx.get("base_cell"),
                            }
                        )
            else:
                for jd in ui_blk.get("join_defs") or []:
                    if not isinstance(jd, dict):
                        continue
                    target = str(jd.get("item") or "").strip()
                    if not target:
                        continue
                    for iter_ctx in iter_contexts:
                        v = _extract_from_cell_rule_with_context(file_path, src, jd, iter_ctx)
                        bundle["join_values"].setdefault(target, []).append(v)
                        bundle["join_contexts"].setdefault(target, []).append(
                            {
                                "file_path": str(iter_ctx.get("file_path") or file_path),
                                "iter_index": int(iter_ctx.get("iter_index", 0)),
                                "base_cell": iter_ctx.get("base_cell"),
                            }
                        )
        return bundle

    cell_spans: dict[int, tuple[int, int]] = {}
    prim_vals = extract_item_values(
        file_path,
        item_config,
        item_id=item_id,
        cell_positions=cell_positions,
        max_primary_rows=max_primary_rows,
        cell_source_spans_out=cell_spans,
    )
    bundle: dict[str, Any] = {
        "primary_values": prim_vals,
        "iteration_contexts": [
            {
                "file_path": str(file_path),
                "iter_index": int(i),
                "base_cell": None,
                "base_row": None,
                "base_col": None,
                "filter_snapshot": {
                    "file_path": str(file_path),
                    "passed": True,
                },
                "primary_value": prim_vals[i] if i < len(prim_vals) else None,
            }
            for i in range(len(prim_vals or []))
        ],
        "link_values": {},
        "link_contexts": {},
        "join_values": {},
        "join_contexts": {},
        "path_item_values": {},
        "path_item_contexts": {},
        "_cell_source_spans": dict(cell_spans),
    }
    run_link_join = debug_step_scope is None
    for si, src in enumerate(sources):
        if not isinstance(src, dict):
            continue
        stype = (src.get("type") or "cell").strip().lower()
        ui_blk = source_ui_block(src)
        if not isinstance(ui_blk, dict):
            continue
        if stype == "cell":
            if not source_passes_file_name_filter(file_path, src):
                continue
            sp = cell_spans.get(si)
            if not sp or not isinstance(sp, (tuple, list)) or len(sp) != 2:
                continue
            try:
                g_off = int(sp[0])
                n_src = int(sp[1])
            except (TypeError, ValueError):
                continue
            if n_src < 1:
                continue
            src_base = str(src.get("cell_ref") or "A1")
            try:
                row_step = int(src.get("row_offset") or 0)
            except (TypeError, ValueError):
                row_step = 0
            try:
                col_step = int(src.get("col_offset") or 0)
            except (TypeError, ValueError):
                col_step = 0
            iter_contexts = []
            for local_i in range(n_src):
                base_cell_i = _resolve_cell_with_offset(
                    src_base,
                    row_step * local_i,
                    col_step * local_i,
                )
                iter_contexts.append(
                    {
                        "file_path": str(file_path),
                        "iter_index": int(g_off + local_i),
                        "rule_iter_index": int(local_i),
                        "base_cell": base_cell_i,
                    }
                )
            b_ictx = bundle.get("iteration_contexts") or []
            for local_i, ic in enumerate(iter_contexts):
                gi = g_off + local_i
                if gi < len(b_ictx) and isinstance(b_ictx[gi], dict):
                    b_ictx[gi]["base_cell"] = ic.get("base_cell")
            if run_link_join:
                for ld in ui_blk.get("link_defs") or []:
                    if not isinstance(ld, dict):
                        continue
                    target = str(ld.get("item") or "").strip()
                    if not target:
                        continue
                    for iter_ctx in iter_contexts:
                        v = _extract_from_cell_rule_with_context(file_path, src, ld, iter_ctx)
                        bundle["link_values"].setdefault(target, []).append(v)
                        bundle["link_contexts"].setdefault(target, []).append(
                            {
                                "file_path": str(iter_ctx.get("file_path") or file_path),
                                "iter_index": int(iter_ctx.get("iter_index", 0)),
                                "base_cell": iter_ctx.get("base_cell"),
                            }
                        )
                for jd in ui_blk.get("join_defs") or []:
                    if not isinstance(jd, dict):
                        continue
                    target = str(jd.get("item") or "").strip()
                    if not target:
                        continue
                    for iter_ctx in iter_contexts:
                        v = _extract_from_cell_rule_with_context(file_path, src, jd, iter_ctx)
                        bundle["join_values"].setdefault(target, []).append(v)
                        bundle["join_contexts"].setdefault(target, []).append(
                            {
                                "file_path": str(iter_ctx.get("file_path") or file_path),
                                "iter_index": int(iter_ctx.get("iter_index", 0)),
                                "base_cell": iter_ctx.get("base_cell"),
                            }
                        )
        elif stype == "name_extract":
            if not name_extract_search_matches(file_path, src):
                continue
            target = str(ui_blk.get("path_item") or "").strip()
            if target:
                # 名前系の連携キー（フルパスで結合）は、まずフルパス列として保持する。
                pv = extract_metadata(file_path, SOURCE_FULL_PATH)
                bundle["path_item_values"].setdefault(target, []).append(pv)
                bundle["path_item_contexts"].setdefault(target, []).append(
                    {
                        "file_path": str(file_path),
                        "iter_index": int(len(bundle["path_item_values"][target]) - 1),
                    }
                )
    # セル座標系: 照合列（名前取得の path_item で決まる列、レガシーは join_path_item_id）へ正規化フルパスを付与
    from svc.svc_data_agg_scenario import LINEAGE_CELL, infer_item_lineage  # noqa: E402

    jh = (join_path_header or "").strip()
    if jh and infer_item_lineage(sources) == LINEAGE_CELL:
        from svc.data_agg_path_norm import normalize_source_path  # noqa: E402

        norm = normalize_source_path(file_path)
        prim = bundle.get("primary_values") or []
        n = max(1, len(prim))
        bundle["path_item_values"][jh] = [norm] * n
        bundle["path_item_contexts"][jh] = [
            {"file_path": str(file_path), "iter_index": int(i)} for i in range(n)
        ]
    return bundle
