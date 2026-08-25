# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_data_agg_scenario.py
Created: 2026-03-18
Updated: 2026-03-18
Version: 0.2.2
Purpose:
  データ集約用シナリオの JSON 形式定義・読込・検証・保存 API。
  項目一覧、項目ごとの取得ソース・抽出ルール・書き込みモード、照合キー、走査条件を格納する。
  svc_data_agg から呼び出され、サブモジュールとして分離する。

  JSON スキーマ（要点）:
    - version: 数値（省略可）
    - items[]: { id, name, sources[], write_mode, join_path_item_id?（読込互換のみ・保存時は UI から除去可） }
        - sources[]: { type, scenario_id?, ui_scenario_source_v1?, ... } type は cell | name_extract | metadata | filename 等
        - 同一項目内 sources は cell 系と path_name 系の混在不可
    - match_keys[]: 照合に使う item id の列（AND）
    - scan: { start_path, recursive, extensions, keyword }
    - master_path: マスターファイルパス
    - debug_flags?: { scenario_step?, item_preview? } デバッグ UI 用
History (latest 3):
  - 0.2.2 (2026-04-04) excel_options: new_sheet_custom_name・custom_sheet_name ルール・検証。
  - 0.2.1 (2026-04-04) excel_options（メイン Excel タブ）既定・normalize_excel_options・検証。
  - 0.1.0 (2026-03-18) 新規作成。load_scenario / validate_scenario / save_scenario API。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Optional

_path_svc = Path(__file__).resolve().parent
_root = _path_svc.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_log import get_logger  # noqa: E402
from core.core_value_shape import compile_shape_script  # noqa: E402
from svc.data_agg_source_ui import (  # noqa: E402
    SCENARIO_SOURCE_UI_KEY,
    SCENARIO_SOURCE_UI_KEY_LEGACY,
    source_ui_block,
)

logger = get_logger(__name__)
__version__ = "0.2.2"

# シナリオ JSON のトップレベルキー（必須でないものは省略可）
KEY_ITEMS = "items"
KEY_MATCH_KEYS = "match_keys"
KEY_SCAN = "scan"
KEY_MASTER_PATH = "master_path"
KEY_VERSION = "version"

# 走査条件のキー
KEY_START_PATH = "start_path"
KEY_RECURSIVE = "recursive"
KEY_EXTENSIONS = "extensions"
KEY_KEYWORD = "keyword"

# 項目（item）のキー
KEY_ITEM_ID = "id"
KEY_ITEM_NAME = "name"
KEY_ITEM_SOURCES = "sources"
KEY_ITEM_WRITE_MODE = "write_mode"
# 結合パス項目：名前・パス文字列系で、どのマスタ項目のパス列と行対応させるか（既定は先頭項目＝主キー）
KEY_JOIN_PATH_ITEM_ID = "join_path_item_id"

# ソース dict：シナリオ行の識別子（ログ・イベントログ用）
KEY_SCENARIO_ID = "scenario_id"

# トップレベル：デバッグ実行の意図（UI が設定）
KEY_DEBUG_FLAGS = "debug_flags"
# メイン画面 Excel タブ（出力先・ジャンプ・並べ替え）。省略時は normalize で補完。
KEY_EXCEL_OPTIONS = "excel_options"

# ソース type（extract_item_values と整合）
SOURCE_TYPE_CELL = "cell"
SOURCE_TYPE_NAME_EXTRACT = "name_extract"
SOURCE_TYPE_METADATA = "metadata"
SOURCE_TYPE_META = "meta"
SOURCE_TYPE_FILENAME = "filename"

# 系統：セル座標系（§要求定義 2.3） / 名前・パス文字列系
LINEAGE_CELL = "cell"
LINEAGE_PATH_NAME = "path_name"

# 書き込みモード（セル座標系項目）
WRITE_MODES_CELL = ("fill_in", "overwrite", "append", "duplicate_append")
# 名前・パス文字列系: 空き／強制／文頭追加／文末追加（区切りなし連結。§2 項6）
WRITE_MODES_NAME = ("fill_in", "overwrite", "prepend", "append_end")
# 後方互換・型チェック用の総集合
WRITE_MODES = WRITE_MODES_CELL


def normalize_item_write_mode(wm: Any, *, lineage: Optional[str] = None) -> str:
    """items[].write_mode を系統に応じて正規化する。lineage は LINEAGE_CELL / LINEAGE_PATH_NAME / None（セル扱い）。"""
    s = (str(wm).strip().lower() if wm is not None else "") or "fill_in"
    if lineage == LINEAGE_PATH_NAME:
        if s in WRITE_MODES_NAME:
            return s
        return "fill_in"
    if s in WRITE_MODES_CELL:
        return s
    return "fill_in"


def _normalize_scenario_payload(data: dict[str, Any]) -> None:
    """読込直後に write_mode および UI 内の旧書込みインデックスを補正する（インプレース）。"""
    items = data.get(KEY_ITEMS)
    if isinstance(items, list):
        for it in items:
            if not isinstance(it, dict):
                continue
            lin = infer_item_lineage(it.get(KEY_ITEM_SOURCES) or [])
            if lin == "__mixed__":
                lin = None
            it[KEY_ITEM_WRITE_MODE] = normalize_item_write_mode(
                it.get(KEY_ITEM_WRITE_MODE), lineage=lin
            )
            for src in it.get(KEY_ITEM_SOURCES) or []:
                if not isinstance(src, dict):
                    continue
                pb = source_ui_block(src)
                if not isinstance(pb, dict):
                    continue
                wmc = pb.get("write_mode_cell_idx")
                if isinstance(wmc, int) and wmc >= 4:
                    pb["write_mode_cell_idx"] = 0
                wmn = pb.get("write_mode_name_idx")
                if isinstance(wmn, int) and wmn >= len(WRITE_MODES_NAME):
                    pb["write_mode_name_idx"] = 0


def load_scenario(path: str | Path) -> dict[str, Any]:
    """
    シナリオファイル（JSON）を読み込み、辞書で返す。

    【概要】
      指定パスから UTF-8 で JSON を読む。ファイルが存在しない・JSON が不正な場合は例外を送出する。

    【引数】
      path: シナリオファイルのパス（.json / .scenario 等）。

    【戻り値】
      シナリオの辞書（items, match_keys, scan, master_path 等を含む）。

    【例外】
      FileNotFoundError: ファイルが存在しない。
      json.JSONDecodeError: JSON として不正。
    """
    p = Path(path).resolve()
    if not p.is_file():
        raise FileNotFoundError("シナリオファイルが存在しません: %s" % p)
    text = p.read_text(encoding="utf-8")
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("シナリオのルートはオブジェクトである必要があります")
    _normalize_scenario_payload(data)
    logger.info("[DATA_AGG_SCENARIO] 読込 完了 path=%s", p)
    return data


def _classify_source_lineage(src: dict[str, Any]) -> str:
    """
    1 ソース dict の系統を返す。cell または path_name。
    """
    stype = (src.get("type") or SOURCE_TYPE_CELL).strip().lower()
    if stype == SOURCE_TYPE_CELL:
        return LINEAGE_CELL
    if stype in (
        SOURCE_TYPE_NAME_EXTRACT,
        SOURCE_TYPE_METADATA,
        SOURCE_TYPE_META,
        SOURCE_TYPE_FILENAME,
    ):
        return LINEAGE_PATH_NAME
    # 未知は cell 扱い（後方互換）
    return LINEAGE_CELL


def _resolve_path_item_label_to_header(
    path_item: str,
    items: list[dict[str, Any]],
    headers: list[str],
) -> str:
    """
    名前取得 UI の path_item（表示ラベル）をマスタ列名へ解決する（items 順と headers 順が対応）。
    svc_data_agg._resolve_path_item_label_to_header と同趣旨（循環 import 回避のためここに保持）。
    """
    if not headers:
        return ""
    p = (path_item or "").strip()
    if not p:
        return headers[0]
    if p in headers:
        return p
    for idx, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        iname = str(it.get("name") or "").strip()
        iid = str(it.get("id") or "").strip()
        if p == iname or p == iid:
            return headers[idx] if idx < len(headers) else headers[0]
    if "主キー" in p or "先頭" in p or "項目一覧" in p:
        return headers[0]
    return headers[0]


def infer_item_lineage(sources: list[Any]) -> Optional[str]:
    """
    項目の sources から系統を推定する。空なら None。
    混在時は "__mixed__" を返す（判定用途）。
    """
    if not sources:
        return None
    kinds: set[str] = set()
    for src in sources:
        if not isinstance(src, dict):
            continue
        kinds.add(_classify_source_lineage(src))
    if not kinds:
        return None
    if len(kinds) > 1:
        return "__mixed__"
    return next(iter(kinds))


def validate_scenario(data: dict[str, Any]) -> list[str]:
    """
    シナリオ辞書の簡易検証を行い、エラーメッセージのリストを返す。

    【概要】
      items がリストであること、各項目に id / name が含まれること、write_mode が許容値であること、
      match_keys がリストであること、scan が辞書であることなどを確認する。
      詳細なソース・抽出ルールの検証は抽出エンジン側で行う想定。

    【引数】
      data: load_scenario で得た辞書。

    【戻り値】
      検証エラーのメッセージリスト。空リストのときは検証通過。
    """
    errors: list[str] = []
    if not isinstance(data, dict):
        errors.append("シナリオはオブジェクトである必要があります")
        return errors

    dbg = data.get(KEY_DEBUG_FLAGS)
    if dbg is not None and not isinstance(dbg, dict):
        errors.append("debug_flags はオブジェクトである必要があります")

    all_item_ids: set[str] = set()
    all_item_names: set[str] = set()
    items = data.get(KEY_ITEMS)
    if items is not None:
        if not isinstance(items, list):
            errors.append("items は配列である必要があります")
        else:
            for i, it in enumerate(items):
                if not isinstance(it, dict):
                    errors.append("items[%s] はオブジェクトである必要があります" % i)
                    continue
                if not (it.get(KEY_ITEM_ID) or it.get(KEY_ITEM_NAME)):
                    errors.append("items[%s] に id または name が必要です" % i)
                iid = str(it.get(KEY_ITEM_ID) or "").strip()
                if iid:
                    all_item_ids.add(iid)
                iname = str(it.get(KEY_ITEM_NAME) or "").strip()
                if iname:
                    all_item_names.add(iname)
                sources = it.get(KEY_ITEM_SOURCES) or []
                wm = it.get(KEY_ITEM_WRITE_MODE)
                if wm is not None:
                    raw_wm = str(wm).strip().lower()
                    lin_wm = infer_item_lineage(sources)
                    if lin_wm == LINEAGE_PATH_NAME:
                        if raw_wm not in WRITE_MODES_NAME:
                            errors.append(
                                "items[%s].write_mode（名前・パス系）は %s のいずれかです"
                                % (i, WRITE_MODES_NAME)
                            )
                    elif lin_wm in (LINEAGE_CELL, None):
                        if raw_wm not in WRITE_MODES_CELL:
                            errors.append(
                                "items[%s].write_mode（セル系）は %s のいずれかです"
                                % (i, WRITE_MODES_CELL)
                            )
                mix = infer_item_lineage(sources)
                if mix == "__mixed__":
                    errors.append(
                        "items[%s]: 同一項目内でセル座標系と名前・パス文字列系のソースを混在できません"
                        % i
                    )
                if isinstance(sources, list) and sources:
                    for j, src in enumerate(sources):
                        if not isinstance(src, dict):
                            errors.append("items[%s].sources[%s] はオブジェクトです" % (i, j))
                            continue
                        sid = src.get(KEY_SCENARIO_ID)
                        if sid is not None and not isinstance(sid, (str, int)):
                            errors.append(
                                "items[%s].sources[%s].scenario_id は文字列または数値です" % (i, j)
                            )
                        raw_ui = src.get(SCENARIO_SOURCE_UI_KEY)
                        if raw_ui is not None and not isinstance(raw_ui, dict):
                            errors.append(
                                "items[%s].sources[%s].ui_scenario_source_v1 はオブジェクトです" % (i, j)
                            )
                        raw_leg = src.get(SCENARIO_SOURCE_UI_KEY_LEGACY)
                        if raw_leg is not None and not isinstance(raw_leg, dict):
                            errors.append(
                                "items[%s].sources[%s].旧形式 UI ブロック（レガシーキー）はオブジェクトです" % (i, j)
                            )
                        pb = source_ui_block(src)
                        if not isinstance(pb, dict):
                            continue
                        vs_raw = pb.get("value_shape_script")
                        if vs_raw is not None and str(vs_raw).strip():
                            ok_vs, msg_vs = compile_shape_script(str(vs_raw))
                            if not ok_vs:
                                errors.append(
                                    "items[%s].sources[%s].value_shape_script: %s" % (i, j, msg_vs)
                                )
                        st = (src.get("type") or SOURCE_TYPE_CELL).strip().lower()
                        if st == SOURCE_TYPE_CELL:
                            ldefs = pb.get("link_defs") or []
                            if not isinstance(ldefs, list):
                                errors.append("items[%s].sources[%s].link_defs は配列です" % (i, j))
                            else:
                                for k, ld in enumerate(ldefs):
                                    if not isinstance(ld, dict):
                                        errors.append("items[%s].sources[%s].link_defs[%s] はオブジェクトです" % (i, j, k))
                                        continue
                                    if not str(ld.get("item") or "").strip():
                                        errors.append("items[%s].sources[%s].link_defs[%s].item は必須です" % (i, j, k))
                                    for rk in ("row", "col"):
                                        rv = ld.get(rk, 0)
                                        if not isinstance(rv, (int, float)):
                                            errors.append("items[%s].sources[%s].link_defs[%s].%s は数値です" % (i, j, k, rk))
                                    lvss = ld.get("value_shape_script")
                                    if lvss is not None and str(lvss).strip():
                                        ok_ls, msg_ls = compile_shape_script(str(lvss))
                                        if not ok_ls:
                                            errors.append(
                                                "items[%s].sources[%s].link_defs[%s].value_shape_script: %s"
                                                % (i, j, k, msg_ls)
                                            )
                            jdefs = pb.get("join_defs") or []
                            if not isinstance(jdefs, list):
                                errors.append("items[%s].sources[%s].join_defs は配列です" % (i, j))
                            else:
                                for k, jd in enumerate(jdefs):
                                    if not isinstance(jd, dict):
                                        errors.append("items[%s].sources[%s].join_defs[%s] はオブジェクトです" % (i, j, k))
                                        continue
                                    if not str(jd.get("item") or "").strip():
                                        errors.append("items[%s].sources[%s].join_defs[%s].item は必須です" % (i, j, k))
                                    for rk in ("row", "col"):
                                        rv = jd.get(rk, 0)
                                        if not isinstance(rv, (int, float)):
                                            errors.append("items[%s].sources[%s].join_defs[%s].%s は数値です" % (i, j, k, rk))
                                    jvss = jd.get("value_shape_script")
                                    if jvss is not None and str(jvss).strip():
                                        ok_js, msg_js = compile_shape_script(str(jvss))
                                        if not ok_js:
                                            errors.append(
                                                "items[%s].sources[%s].join_defs[%s].value_shape_script: %s"
                                                % (i, j, k, msg_js)
                                            )
                            ro = int(src.get("row_offset") or 0)
                            co = int(src.get("col_offset") or 0)
                            ru = bool(src.get("repeat_until_empty", True))
                            rlast = bool(src.get("repeat_until_last", False))
                            rm = src.get("repeat_max")
                            if (
                                ((ru and (rm is None or int(rm or 0) <= 0)) or rlast)
                                and ro == 0
                                and co == 0
                            ):
                                errors.append(
                                    "items[%s].sources[%s]: 行・列移動オフセットがともに 0 のときは"
                                    "終結「空白まで／終端」は指定できません。N件を指定してください。"
                                    % (i, j)
                                )
                        elif st == SOURCE_TYPE_NAME_EXTRACT:
                            pit = pb.get("path_item")
                            if pit is not None and not isinstance(pit, str):
                                errors.append(
                                    "items[%s].sources[%s].path_item（名前系 UI ブロック内）は文字列です" % (i, j)
                                )
                jp = it.get(KEY_JOIN_PATH_ITEM_ID)
                if jp is not None and not isinstance(jp, (str, int)):
                    errors.append("items[%s].join_path_item_id は文字列または id です" % i)

            for i, it in enumerate(items):
                if not isinstance(it, dict):
                    continue
                jp = it.get(KEY_JOIN_PATH_ITEM_ID)
                if jp is None or jp == "":
                    continue
                if isinstance(jp, (str, int)):
                    jps = str(jp).strip()
                    if jps and (all_item_ids or all_item_names):
                        if jps not in all_item_ids and jps not in all_item_names:
                            errors.append(
                                "items[%s].join_path_item_id=%r は items[].id または name と一致しません"
                                % (i, jps)
                            )

            hdrs = [
                (
                    str(it.get(KEY_ITEM_NAME) or it.get(KEY_ITEM_ID) or ("項目_%s" % ii))
                    if isinstance(it, dict)
                    else "項目_%s" % ii
                )
                for ii, it in enumerate(items)
            ]
            for ii, it in enumerate(items):
                if not isinstance(it, dict):
                    continue
                for j, src in enumerate(it.get(KEY_ITEM_SOURCES) or []):
                    if not isinstance(src, dict):
                        continue
                    if (src.get("type") or "").strip().lower() != SOURCE_TYPE_NAME_EXTRACT:
                        continue
                    pb = source_ui_block(src)
                    if not isinstance(pb, dict):
                        continue
                    pit = str(pb.get("path_item") or "").strip()
                    if not pit:
                        continue
                    path_h = _resolve_path_item_label_to_header(pit, items, hdrs)
                    try:
                        path_idx = hdrs.index(path_h)
                    except ValueError:
                        path_idx = 0
                    if path_idx >= ii:
                        errors.append(
                            "items[%s].sources[%s]: 名前から取得の関連付け列は、マスタ項目リストで"
                            "当該項目（書込み先）より上（先頭に近い側）にあり、同一列は不可です。"
                            % (ii, j)
                        )

    mk = data.get(KEY_MATCH_KEYS)
    if mk is not None and not isinstance(mk, list):
        errors.append("match_keys は配列である必要があります")
    elif isinstance(mk, list) and items and isinstance(items, list):
        for k, ref in enumerate(mk):
            if ref is None:
                continue
            r = str(ref).strip()
            if r and all_item_ids and r not in all_item_ids:
                errors.append(
                    "match_keys[%s]=%r は items[].id のいずれとも一致しません" % (k, ref)
                )

    scan = data.get(KEY_SCAN)
    if scan is not None and not isinstance(scan, dict):
        errors.append("scan はオブジェクトである必要があります")

    exo = data.get(KEY_EXCEL_OPTIONS)
    if exo is not None:
        if not isinstance(exo, dict):
            errors.append("excel_options はオブジェクトである必要があります")
        else:
            sk = exo.get("sort_keys")
            if sk is not None:
                if not isinstance(sk, list):
                    errors.append("excel_options.sort_keys は配列である必要があります")
                else:
                    for si, ent in enumerate(sk):
                        if not isinstance(ent, dict):
                            errors.append(
                                "excel_options.sort_keys[%s] はオブジェクトである必要があります" % si
                            )
            exn = normalize_excel_options(exo)
            if (
                exn.get("output_target") == "new_sheet"
                and exn.get("new_sheet_name_rule") == "custom_sheet_name"
                and not str(exn.get("new_sheet_custom_name") or "").strip()
            ):
                errors.append(
                    "excel_options: 新規シートで「シート名入力」のときはシート名（new_sheet_custom_name）を入力してください。"
                )

    return errors


def save_scenario(path: str | Path, data: dict[str, Any]) -> None:
    """
    シナリオ辞書を指定パスに JSON で保存する。

    【概要】
      UTF-8 で書き出し、インデント 2 で整形する。既存ファイルは上書きする。

    【引数】
      path: 保存先パス（.json / .scenario 等）。
      data: シナリオ辞書（validate_scenario で検証済みであることが望ましい）。

    【戻り値】
      なし
    """
    p = Path(path).resolve()
    p.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(data, ensure_ascii=False, indent=2)
    p.write_text(text, encoding="utf-8")
    logger.info("[DATA_AGG_SCENARIO] 保存 完了 path=%s", p)


def default_excel_options() -> dict[str, Any]:
    """メイン画面 Excel タブの既定値（保存 JSON の excel_options と一致）。"""
    return {
        "output_target": "active_sheet",
        "write_mode": "append",
        "anchor_cell": "",
        "new_sheet_name_rule": "scenario_name_seq",
        "new_sheet_custom_name": "",
        "jump_register_name": True,
        "freeze_header_row": True,
        "autofilter": True,
        "sort_keys": [{"item": "", "order": "asc", "natural": True}],
    }


def normalize_excel_options(raw: Any) -> dict[str, Any]:
    """
    excel_options を読込用に正規化する。欠損・未知値は default_excel_options で埋める。
    sort_keys が空リストのときは既定の1行にする。
    """
    d = default_excel_options()
    if not isinstance(raw, dict):
        return d
    ot = str(raw.get("output_target") or "").strip()
    if ot in ("active_sheet", "new_sheet"):
        d["output_target"] = ot
    wm = str(raw.get("write_mode") or "").strip()
    if wm in ("append", "overwrite", "clear_write", "anchor_cell"):
        d["write_mode"] = wm
    ac = raw.get("anchor_cell")
    if ac is not None:
        d["anchor_cell"] = str(ac).strip()
    nsr = str(raw.get("new_sheet_name_rule") or "").strip().lower()
    if nsr in ("scenario_name_seq", "scenario_datetime", "scenario_seq"):
        d["new_sheet_name_rule"] = "scenario_name_seq"
    elif nsr == "custom_sheet_name":
        d["new_sheet_name_rule"] = "custom_sheet_name"
    if "new_sheet_custom_name" in raw:
        d["new_sheet_custom_name"] = str(raw.get("new_sheet_custom_name") or "").strip()[:210]
    if "jump_register_name" in raw:
        d["jump_register_name"] = bool(raw.get("jump_register_name"))
    if "freeze_header_row" in raw:
        d["freeze_header_row"] = bool(raw.get("freeze_header_row"))
    if "autofilter" in raw:
        d["autofilter"] = bool(raw.get("autofilter"))
    sk_raw = raw.get("sort_keys")
    if isinstance(sk_raw, list) and sk_raw:
        rows: list[dict[str, Any]] = []
        for ent in sk_raw:
            if not isinstance(ent, dict):
                continue
            order = str(ent.get("order") or "asc").strip().lower()
            if order not in ("asc", "desc"):
                order = "asc"
            rows.append(
                {
                    "item": str(ent.get("item") or "").strip(),
                    "order": order,
                    "natural": bool(ent.get("natural")),
                }
            )
        if rows:
            d["sort_keys"] = rows
    return d


def create_empty_scenario() -> dict[str, Any]:
    """
    空のシナリオ構造を返す。項目一覧・走査条件のひな形。

    【概要】
      UI で新規シナリオを作成するときの初期値として使用する。

    【戻り値】
      items: [], match_keys: [], scan: { start_path, recursive, extensions, keyword }, master_path: "" 等を含む辞書。
    """
    return {
        KEY_VERSION: 1,
        KEY_ITEMS: [],
        KEY_MATCH_KEYS: [],
        KEY_SCAN: {
            KEY_START_PATH: "",
            KEY_RECURSIVE: False,
            KEY_EXTENSIONS: [".xlsx", ".xlsm", ".xls", ".csv"],
            KEY_KEYWORD: "",
        },
        KEY_MASTER_PATH: "",
        KEY_DEBUG_FLAGS: {
            "scenario_step": False,
            "item_preview": False,
        },
        KEY_EXCEL_OPTIONS: default_excel_options(),
    }
