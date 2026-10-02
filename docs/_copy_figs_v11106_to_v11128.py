# -*- coding: utf-8 -*-
"""Ver1.1.10.6A の埋め込み図を、同図番号の Ver1.1.12.8 画像挿入枠へ流用する。"""
from __future__ import annotations

import io
import re
import shutil
import tempfile
from collections import defaultdict
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FIG_RE = re.compile(r"(図\d+-\d+)")
DOCS = Path(__file__).resolve().parent

PAIRINGS = [
    (
        "Ver1.1.10.6/データ集約_シナリオ編集画面説明書_V11106A.docx",
        "Ver1.1.12.8/データ集約_シナリオ編集画面説明書_V11128.docx",
    ),
    (
        "Ver1.1.10.6/データ集約_主キー・連携キー・結合キーの動作概念_V11106A.docx",
        "Ver1.1.12.8/データ集約_主キー・連携キー・結合キーの動作概念_V11128.docx",
    ),
    (
        "Ver1.1.10.6/データ集約_整形DSL_コマンドリファレンス_V11106A.docx",
        "Ver1.1.12.8/データ集約_整形DSL_コマンドリファレンス_V11128.docx",
    ),
]

# UI 配置が変わった枠は流用しない（差し替え推奨）
SKIP_FIGS: dict[str, set[str]] = {
    "データ集約_シナリオ編集画面説明書_V11128.docx": {
        "図3-4",  # 主キー：終結／件数／オフセット順が変更
    },
}


def _set_run_font(run, name="游ゴシック", size=9.5, bold=False, color=None):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = name
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)


def _rid_to_media(docx: Path) -> dict[str, bytes]:
    out: dict[str, bytes] = {}
    with ZipFile(docx) as z:
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
        for rid, target in re.findall(
            r'Id="(rId\d+)"[^>]*Target="(media/[^"]+)"', rels
        ):
            data = z.read("word/" + target)
            out[rid] = data
        # Target may be relative without media/ prefix variants
        for rid, target in re.findall(
            r'Target="(media/[^"]+)"[^>]*Id="(rId\d+)"', rels
        ):
            if rid not in out:
                out[rid] = z.read("word/" + target)
    return out


def extract_fig_images(src: Path) -> dict[str, list[bytes]]:
    """
    文書順で drawing の直後〜次の図番号キャプションまでに出た画像を、
    その図番号へ結びつける。
    """
    rid_media = _rid_to_media(src)
    with ZipFile(src) as z:
        xml = z.read("word/document.xml").decode("utf-8")

    # Split into paragraph-ish chunks
    chunks = re.split(r"</w:p>", xml)
    pending_rids: list[str] = []
    fig_imgs: dict[str, list[bytes]] = defaultdict(list)
    seen_assign: set[str] = set()

    for ch in chunks:
        embeds = re.findall(r'r:embed="(rId\d+)"', ch)
        pending_rids.extend(embeds)
        plain = re.sub(r"<[^>]+>", "", ch)
        plain = plain.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
        m = FIG_RE.search(plain)
        if not m:
            continue
        fig = m.group(1)
        # Prefer images accumulated just before this caption
        if pending_rids:
            for rid in pending_rids:
                blob = rid_media.get(rid)
                if blob and blob not in fig_imgs[fig]:
                    fig_imgs[fig].append(blob)
            pending_rids = []
            seen_assign.add(fig)
        elif fig not in seen_assign:
            # caption without preceding image in this pass — leave empty
            pass
    return dict(fig_imgs)


def _clear_paragraph(p) -> None:
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    # also remove drawings left in p
    for el in list(p._p):
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag in ("r", "hyperlink"):
            # keep empty structure minimal; remove runs with drawings
            pass


def fill_slots(dst: Path, fig_imgs: dict[str, list[bytes]], skip: set[str]) -> tuple[int, list[str]]:
    doc = Document(str(dst))
    filled = 0
    notes: list[str] = []
    with tempfile.TemporaryDirectory() as td:
        tdir = Path(td)
        for table in doc.tables:
            if len(table.rows) < 2 or len(table.columns) < 1:
                continue
            cell0 = table.cell(0, 0)
            texts = [(p.text or "").strip() for p in cell0.paragraphs]
            slot_line = next((t for t in texts if t.startswith("【画像挿入】")), "")
            if not slot_line:
                continue
            m = FIG_RE.search(slot_line)
            if not m:
                continue
            fig = m.group(1)
            if fig in skip:
                notes.append(f"SKIP {dst.name} {fig}（UI変更のため流用しない）")
                continue
            blobs = fig_imgs.get(fig) or []
            if not blobs:
                notes.append(f"MISS {dst.name} {fig}（前版に画像なし）")
                continue
            # Keep first paragraph as marker, clear rest of cell0, insert image
            # Rebuild cell0: marker + image + short note
            # Remove all paragraphs except we'll rewrite
            tc = cell0._tc
            for child in list(tc):
                if child.tag == qn("w:p"):
                    tc.remove(child)
            p_mark = cell0.add_paragraph()
            p_mark.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p_mark.add_run(f"【流用】{fig}（Ver1.1.10.6A）")
            _set_run_font(r, size=9, bold=True, color=RGBColor(0x0D, 0x5C, 0x63))
            # Use the largest blob if multiple (often duplicate small/large)
            blob = max(blobs, key=len)
            ext = ".png"
            if blob[:3] == b"\xff\xd8\xff":
                ext = ".jpg"
            img_path = tdir / f"{fig.replace('-', '_')}{ext}"
            img_path.write_bytes(blob)
            p_img = cell0.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_img.add_run()
            # Width ~14cm for A4 content
            run.add_picture(str(img_path), width=Cm(14.0))
            filled += 1
            notes.append(f"OK   {dst.name} {fig}")
    doc.save(str(dst))
    return filled, notes


def main() -> None:
    all_notes: list[str] = []
    total = 0
    for src_rel, dst_rel in PAIRINGS:
        src = DOCS / src_rel
        dst = DOCS / dst_rel
        if not src.exists() or not dst.exists():
            all_notes.append(f"ABSENT {src_rel} or {dst_rel}")
            continue
        # backup once
        bak = dst.with_suffix(dst.suffix + ".pre_img.bak")
        if not bak.exists():
            shutil.copy2(dst, bak)
        figs = extract_fig_images(src)
        skip = SKIP_FIGS.get(dst.name, set())
        n, notes = fill_slots(dst, figs, skip)
        total += n
        all_notes.extend(notes)
        print(dst.name, "filled", n, "source_figs", sorted(figs.keys()))
    print("---")
    for line in all_notes:
        print(line)
    print("TOTAL_FILLED", total)


if __name__ == "__main__":
    main()
