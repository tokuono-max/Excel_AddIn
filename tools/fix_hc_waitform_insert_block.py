# -*- coding: utf-8 -*-
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
WAITFORM = ROOT / "VBA" / "HC_WaitForm.bas"
PATCH = ROOT / "tools" / "patch_waitform_ready_signal_cp932.py"

t = WAITFORM.read_bytes().decode("cp932", errors="strict")
if "Public Sub WaitForUiReadySignal" in t:
    print("already ok")
    raise SystemExit(0)

block_src = PATCH.read_text(encoding="utf-8")
m = re.search(r'WAITFORM_NEW_BLOCK = """(.*?)"""', block_src, re.S)
if not m:
    raise SystemExit("block not found in patch file")
new_block = m.group(1)

needle = "Public Sub NotifyUiReady()"
if needle not in t:
    raise SystemExit("NotifyUiReady not found")

t = t.replace(needle, new_block.rstrip() + "\n" + needle, 1)
WAITFORM.write_bytes(t.replace("\n", "\r\n").encode("cp932"))
print(f"fixed {WAITFORM}")
