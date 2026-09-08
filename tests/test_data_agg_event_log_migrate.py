# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg_write import (  # noqa: E402
    EVENT_LOG_HEADERS,
    EVENT_LOG_MIGRATE_LEAVE_PAST_PC_APL_BLANK,
    classify_event_log_layout,
    migrate_event_log_sheet_layout,
)


def test_classify_event_log_layout_variants() -> None:
    assert (
        classify_event_log_layout(["PC性能", "APL Ver", "記録日時"]) == "current"
    )
    assert classify_event_log_layout(["記録日時", "処理時間", "出力行数"]) == "insert2_shift"
    assert (
        classify_event_log_layout(
            ["記録日時", "処理時間", "出力行数", "APL Ver", "区分"]
        )
        == "insert2_shift"
    )
    assert (
        classify_event_log_layout(
            ["PC性能", "APL Ver", "記録日時"],
            first_data_a1="2026-09-08 12:00:00",
        )
        == "insert2_realign"
    )
    assert (
        classify_event_log_layout(
            ["PC性能", "APL Ver", "記録日時"],
            first_data_a1="Intel / 論理CPU 8 / メモリ 16GB / Windows 11",
        )
        == "current"
    )


class _LastCell:
    def __init__(self, row: int, column: int) -> None:
        self.row = row
        self.column = column


class _UsedRange:
    def __init__(self, sheet: "_GridSheet") -> None:
        self._sheet = sheet

    @property
    def last_cell(self) -> _LastCell:
        r, c = self._sheet._bounds()
        return _LastCell(r, c)


class _ColApi:
    def __init__(self, sheet: "_GridSheet", col: int) -> None:
        self._sheet = sheet
        self._col = col

    def Insert(self) -> None:
        self._sheet._insert_col(self._col)

    def Delete(self) -> None:
        self._sheet._delete_col(self._col)


class _ColumnsApi:
    def __init__(self, sheet: "_GridSheet") -> None:
        self._sheet = sheet

    def __call__(self, col: int) -> _ColApi:
        return _ColApi(self._sheet, int(col))


class _SheetApi:
    def __init__(self, sheet: "_GridSheet") -> None:
        self.Columns = _ColumnsApi(sheet)


class _GridRange:
    def __init__(
        self,
        sheet: "_GridSheet",
        r1: int,
        c1: int,
        r2: int | None = None,
        c2: int | None = None,
    ) -> None:
        self._sheet = sheet
        self._r1 = r1
        self._c1 = c1
        self._r2 = r2 if r2 is not None else r1
        self._c2 = c2 if c2 is not None else c1
        self.api = type("RApi", (), {"EntireColumn": _ColApi(sheet, c1)})()

    def resize(self, rows: int, cols: int) -> "_GridRange":
        return _GridRange(
            self._sheet,
            self._r1,
            self._c1,
            self._r1 + rows - 1,
            self._c1 + cols - 1,
        )

    @property
    def value(self) -> Any:
        out: list[list[Any]] = []
        for r in range(self._r1, self._r2 + 1):
            row: list[Any] = []
            for c in range(self._c1, self._c2 + 1):
                row.append(self._sheet._get(r, c))
            out.append(row)
        if len(out) == 1 and len(out[0]) == 1:
            return out[0][0]
        if len(out) == 1:
            return out[0]
        return out

    @value.setter
    def value(self, data: Any) -> None:
        if data is None:
            for r in range(self._r1, self._r2 + 1):
                for c in range(self._c1, self._c2 + 1):
                    self._sheet._set(r, c, None)
            return
        if not isinstance(data, list):
            self._sheet._set(self._r1, self._c1, data)
            return
        if data and not isinstance(data[0], list):
            # single row list
            for j, v in enumerate(data):
                self._sheet._set(self._r1, self._c1 + j, v)
            return
        for i, row in enumerate(data):
            for j, v in enumerate(row):
                self._sheet._set(self._r1 + i, self._c1 + j, v)


class _GridSheet:
    """列挿入／削除可能な簡易シート（移行テスト用）。"""

    def __init__(self, rows: list[list[Any]]) -> None:
        self._rows = [list(r) for r in rows]
        self.api = _SheetApi(self)

    @property
    def used_range(self) -> _UsedRange:
        return _UsedRange(self)

    def _bounds(self) -> tuple[int, int]:
        if not self._rows:
            return 0, 0
        max_c = max((len(r) for r in self._rows), default=0)
        # trim trailing empty rows
        last_r = 0
        for i, r in enumerate(self._rows, start=1):
            if any(str(x or "").strip() != "" for x in r):
                last_r = i
        return last_r, max_c

    def _get(self, r: int, c: int) -> Any:
        if r < 1 or c < 1:
            return None
        if r > len(self._rows):
            return None
        row = self._rows[r - 1]
        if c > len(row):
            return None
        return row[c - 1]

    def _set(self, r: int, c: int, v: Any) -> None:
        while len(self._rows) < r:
            self._rows.append([])
        row = self._rows[r - 1]
        while len(row) < c:
            row.append(None)
        row[c - 1] = v

    def _insert_col(self, col_1based: int) -> None:
        idx = max(0, int(col_1based) - 1)
        for row in self._rows:
            if idx >= len(row):
                row.append(None)
            else:
                row.insert(idx, None)

    def _delete_col(self, col_1based: int) -> None:
        idx = int(col_1based) - 1
        if idx < 0:
            return
        for row in self._rows:
            if idx < len(row):
                del row[idx]

    def range(self, a: Any, b: Any = None) -> _GridRange:
        if isinstance(a, str):
            # unused in tests
            return _GridRange(self, 1, 1)
        if b is None:
            r1, c1 = a
            return _GridRange(self, int(r1), int(c1))
        r1, c1 = a
        r2, c2 = b
        return _GridRange(self, int(r1), int(c1), int(r2), int(c2))


class _CoreXlcShim:
    @staticmethod
    def write_chunk(sheet: Any, start_y: int, start_x: int, data_list: list, text_mode: bool = False) -> None:
        if not data_list:
            return
        rng = sheet.range((start_y, start_x)).resize(len(data_list), len(data_list[0]))
        rng.value = data_list


def test_migrate_insert2_shift_old_9col() -> None:
    old_h = [
        "記録日時",
        "処理時間",
        "出力行数",
        "区分",
        "書込み方式",
        "出力シート名",
        "シナリオID",
        "対象パス",
        "詳細",
    ]
    ws = _GridSheet(
        [
            old_h,
            [
                "2026-09-01 10:00:00",
                "1.00 秒",
                3,
                "一括実行・完了",
                "追加",
                "Sheet1",
                "sid",
                r"C:\a.json",
                "{}",
            ],
        ]
    )
    action = migrate_event_log_sheet_layout(ws, core_xlc=_CoreXlcShim)
    assert action == "insert2_shift"
    assert ws._rows[0][:5] == EVENT_LOG_HEADERS[:5]
    assert ws._rows[1][2] == "2026-09-01 10:00:00"
    assert ws._rows[1][5] == "一括実行・完了"
    # 過去行の PC性能 / APL Ver は空欄
    assert not str(ws._rows[1][0] or "").strip()
    assert not str(ws._rows[1][1] or "").strip()


def test_migrate_insert2_shift_removes_mid_apl() -> None:
    old_h = [
        "記録日時",
        "処理時間",
        "出力行数",
        "APL Ver",
        "区分",
        "書込み方式",
        "出力シート名",
        "シナリオID",
        "対象パス",
        "詳細",
    ]
    ws = _GridSheet(
        [
            old_h,
            [
                "2026-09-01 10:00:00",
                "1 秒",
                1,
                "1.1.10.6",
                "一括実行・完了",
                "",
                "",
                "sid",
                r"C:\a.json",
                "{}",
            ],
        ]
    )
    action = migrate_event_log_sheet_layout(ws, core_xlc=_CoreXlcShim)
    assert action == "insert2_shift"
    assert ws._rows[0][0] == "PC性能"
    assert ws._rows[0][1] == "APL Ver"
    assert ws._rows[0][2] == "記録日時"
    assert ws._rows[0][5] == "区分"
    # APL Ver が中間に二重でない
    assert ws._rows[0].count("APL Ver") == 1
    assert ws._rows[1][2] == "2026-09-01 10:00:00"
    assert ws._rows[1][5] == "一括実行・完了"
    assert not str(ws._rows[1][0] or "").strip()
    assert not str(ws._rows[1][1] or "").strip()


def test_migrate_realign_miswritten_header() -> None:
    # ヘッダだけ新形式・データは旧のまま（左ずれ未実施）
    ws = _GridSheet(
        [
            list(EVENT_LOG_HEADERS),
            [
                "2026-09-01 10:00:00",
                "1 秒",
                2,
                "一括実行・完了",
                "",
                "",
                "sid",
                r"C:\a.json",
                "{}",
                "",
                "",
            ],
        ]
    )
    action = migrate_event_log_sheet_layout(ws, core_xlc=_CoreXlcShim)
    assert action == "insert2_realign"
    assert ws._rows[0][:3] == ["PC性能", "APL Ver", "記録日時"]
    assert ws._rows[1][2] == "2026-09-01 10:00:00"
    assert ws._rows[1][5] == "一括実行・完了"
    assert not str(ws._rows[1][0] or "").strip()
    assert not str(ws._rows[1][1] or "").strip()


def test_event_log_migrate_policy_flag_forbids_past_row_backfill() -> None:
    """再発防止: 過去行への PC/APL 埋め込み禁止フラグは True（空欄維持）のまま。"""
    assert EVENT_LOG_MIGRATE_LEAVE_PAST_PC_APL_BLANK is True


def test_ui_help_pc_apl_item_matches_blank_past_rows_policy() -> None:
    """再発防止: ヘルプが『過去行は空欄・新規から記録』から外れないこと。"""
    import json

    help_path = Path(__file__).resolve().parents[1] / "config" / "ui_help.json"
    data = json.loads(help_path.read_text(encoding="utf-8"))
    bins = ((data.get("VER_HISTORY") or {}).get("BIN") or [])
    items: list[str] = []
    for block in bins:
        if str(block.get("version") or "") == "1.1.11.7":
            items = [str(x) for x in (block.get("items") or [])]
            break
    hit = [s for s in items if "PC性能" in s and "APL Ver" in s]
    assert hit, "ui_help 1.1.11.7 に PC性能/APL Ver の履歴が無い"
    text = hit[0]
    assert "空欄" in text
    assert "新規" in text
    assert "既存行へも値" not in text
    assert "既存行へ値" not in text
