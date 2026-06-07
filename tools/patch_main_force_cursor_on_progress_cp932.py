# -*- coding: utf-8 -*-
"""Main.bas: ForceCursorOnProgress 追加に伴うモジュール/プロシージャヘッダを CP932 で追記する。

マスター: VBA/Main.bas（SHIFT-JIS / CP932）。_xlam_extract は正本ではない。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"

UPDATE_DATE = "2026-06-06"
MODULE_VERSION = "2.13.0"
MODULE_ENTRY = (
    f"'   {MODULE_VERSION} ({UPDATE_DATE}) [UX] ForceCursorOnProgress: 進捗ダイアログ表示時の砂時計 ON。"
    "保険タイマなし。外部 Python から Application.Run で呼ぶ（ProgressDialog show/teardown のみ制御）。\n"
)
PROC_HDR = (
    "' ---------------------------------------------------------------------------------------------------------------------\n"
    "' プロシージャ名: ForceCursorOnProgress\n"
    "' 改版番号および履歴:\n"
    f"'   1.0.0 ({UPDATE_DATE}) 進捗ダイアログ表示開始: xlWait ON（StartCursorGuardTimer は呼ばない）。\n"
    "' プロシージャの動作概要: ProgressDialog の showEvent からのみ砂時計を ON にする。\n"
    "' 呼出し例: Application.Run \"Main.ForceCursorOnProgress\", sheetId\n"
    "' ヘルパープロシージャの親子関係: (なし) ForceCursorOff は Python 側 notify_ui_ready で解除\n"
    "' ---------------------------------------------------------------------------------------------------------------------\n"
)
PROC_NEEDLE = 'Public Sub ForceCursorOnProgress(Optional ByVal sId As String = "progress")'
HIST_ANCHOR = "' 改版番号および履歴:\n"
OLD_MODULE_ENTRY_RE = re.compile(
    r"'   2\.13\.0 \(\d{4}-\d{2}-\d{2}\) \[UX\] ForceCursorOnProgress:.*\n"
)


def _read_cp932(path: Path) -> str:
    return path.read_bytes().decode("cp932").replace("\r\n", "\n")


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def patch_main(text: str) -> str:
    if HIST_ANCHOR not in text:
        raise ValueError("hist anchor not found")

    new_date_line = f"' 更新日: {UPDATE_DATE}\n"
    if re.search(r"' 更新日: \d{4}-\d{2}-\d{2}\n", text):
        text = re.sub(r"' 更新日: \d{4}-\d{2}-\d{2}\n", new_date_line, text, count=1)
    else:
        raise ValueError("更新日 line not found")

    if OLD_MODULE_ENTRY_RE.search(text):
        text = OLD_MODULE_ENTRY_RE.sub(MODULE_ENTRY, text, count=1)
    elif MODULE_VERSION not in text:
        text = text.replace(HIST_ANCHOR, HIST_ANCHOR + MODULE_ENTRY, 1)

    if PROC_NEEDLE in text and "' プロシージャ名: ForceCursorOnProgress" not in text:
        text = text.replace(PROC_NEEDLE, PROC_HDR + PROC_NEEDLE, 1)
    elif "' プロシージャ名: ForceCursorOnProgress" in text:
        text = re.sub(
            r"(' プロシージャ名: ForceCursorOnProgress\n"
            r"' 改版番号および履歴:\n"
            r")'   1\.0\.0 \(\d{4}-\d{2}-\d{2}\)",
            rf"\1'   1.0.0 ({UPDATE_DATE})",
            text,
            count=1,
        )

    return text


def main() -> int:
    t = patch_main(_read_cp932(MAIN))
    _write_cp932(MAIN, t)
    print(f"patched {MAIN} 更新日={UPDATE_DATE} {MODULE_VERSION}")
    _read_cp932(MAIN)
    return 0


if __name__ == "__main__":
    sys.exit(main())
