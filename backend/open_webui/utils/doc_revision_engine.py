# -*- coding: utf-8 -*-
"""
Word (.doc / .docx) Read, Write, and Track Changes (修订模式 / 审阅模式) Engine for Open WebUI.
Supports:
1. Reading full text (paragraphs + tables) from .doc and .docx files.
2. Revising existing .doc / .docx files with character-level OpenXML Track Changes (<w:del>, <w:ins>, <w:trackRevisions/>),
   preserving original typography, styles, tables, headers, and footers, and exporting both .doc and .docx.
3. Creating new .doc / .docx files from scratch with Chinese contract formatting and inline Track Changes markup
   (e.g. [-删除原文-]{+新增内容+}, ~~删除原文~~++新增内容++, <del>...</del><ins>...</ins>, or replacement pairs).
"""

from __future__ import annotations

import copy
import difflib
import hashlib
import json
import logging
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
import time
import uuid
from datetime import datetime, timezone
from typing import Any, Optional

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from lxml import etree

log = logging.getLogger(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
UPLOAD_DIR = "/app/backend/data/uploads"
DB_PATH = "/app/backend/data/webui.db"


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def run_soffice_convert(input_path: str, target_ext: str, out_dir: str) -> str:
    """Convert between .doc and .docx using an isolated LibreOffice profile."""
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_", dir="/tmp")
    try:
        cmd = [
            "soffice",
            f"-env:UserInstallation=file://{profile_dir}",
            "--headless",
            "--convert-to",
            target_ext,
            "--outdir",
            out_dir,
            input_path,
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        base = os.path.splitext(os.path.basename(input_path))[0]
        expected = os.path.join(out_dir, f"{base}.{target_ext}")
        if os.path.exists(expected):
            return expected
        candidates = [
            os.path.join(out_dir, f)
            for f in os.listdir(out_dir)
            if f.lower().endswith(f".{target_ext.lower()}") and os.path.join(out_dir, f) != input_path
        ]
        if candidates:
            return candidates[0]
        raise RuntimeError(f"LibreOffice conversion to .{target_ext} produced no output file.")
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def iter_all_paragraphs(doc: docx.document.Document):
    """Yield all unique paragraphs across body, tables (including nested), headers, and footers."""
    seen = set()

    def _from_container(container):
        for p in getattr(container, "paragraphs", []):
            if id(p._p) not in seen:
                seen.add(id(p._p))
                yield p
        for t in getattr(container, "tables", []):
            yield from _from_table(t)

    def _from_table(table):
        for row in table.rows:
            for cell in row.cells:
                yield from _from_container(cell)

    yield from _from_container(doc)
    for section in doc.sections:
        for hf in (section.header, section.footer):
            if hf is not None:
                yield from _from_container(hf)


def normalize_paragraph_runs(p_el: etree._Element) -> None:
    """Split any <w:r> containing multiple content children (e.g. w:t + w:br + w:t) into individual <w:r> elements."""
    for child in list(p_el):
        if child.tag != w("r"):
            continue
        rPr = child.find(w("rPr"))
        non_rpr = [c for c in child if c.tag != w("rPr")]
        if len(non_rpr) <= 1:
            continue
        idx = list(p_el).index(child)
        p_el.remove(child)
        for offset, sub in enumerate(non_rpr):
            new_r = etree.Element(w("r"))
            if rPr is not None:
                new_r.append(copy.deepcopy(rPr))
            new_r.append(sub)
            p_el.insert(idx + offset, new_r)


def make_text_run(text: str, rPr: Optional[etree._Element] = None, is_del: bool = False) -> etree._Element:
    """Create a <w:r> element with <w:t> or <w:delText> while preserving run formatting."""
    r = etree.Element(w("r"))
    if rPr is not None:
        r.append(copy.deepcopy(rPr))
    else:
        rPr_new = etree.SubElement(r, w("rPr"))
        rFonts = etree.SubElement(rPr_new, w("rFonts"))
        rFonts.set(w("ascii"), "宋体")
        rFonts.set(w("hAnsi"), "宋体")
        rFonts.set(w("eastAsia"), "宋体")
        rFonts.set(w("cs"), "宋体")
    t = etree.SubElement(r, w("delText") if is_del else w("t"))
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text
    return r


def enable_word_track_changes_settings(doc: docx.document.Document) -> None:
    """Enable <w:trackRevisions/> and <w:revisionView/> in settings.xml so Word opens in Track Changes view."""
    settings = doc.settings.element
    if settings.find(w("trackRevisions")) is None:
        etree.SubElement(settings, w("trackRevisions"))
    rev_view = settings.find(w("revisionView"))
    if rev_view is None:
        rev_view = etree.SubElement(settings, w("revisionView"))
    rev_view.set(w("markup"), "1")
    rev_view.set(w("comments"), "1")
    rev_view.set(w("insDel"), "1")
    rev_view.set(w("formatting"), "1")


def _clean_replacement_text(s: str) -> str:
    """Strip Markdown bold/code wrappers that LLMs often include when quoting contract clauses."""
    s = str(s).strip()
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"^`(.+)`$", r"\1", s)
    return s.strip()


def parse_replacements(raw_replacements: Any) -> dict[str, str]:
    """
    Flexibly parse replacements from:
    - dict: {"旧文本": "新文本"} or {"replacements": [...]}
    - list of dicts: [{"old": "...", "new": "..."}, {"original": "...", "revised": "..."}, {"from": "...", "to": "..."}]
    - JSON string encoding any of the above
    - Markdown table or arrow/colon separated lines
    """
    if not raw_replacements:
        return {}

    parsed = raw_replacements
    if isinstance(raw_replacements, str):
        text = raw_replacements.strip()
        # Try JSON first
        try:
            parsed = json.loads(text)
        except Exception:
            # Try extracting JSON block from string
            m = re.search(r"(\{[\s\S]*\}|\[[\s\S]*\])", text)
            if m:
                try:
                    parsed = json.loads(m.group(1))
                except Exception:
                    parsed = None
            else:
                parsed = None

        if parsed is None:
            result: dict[str, str] = {}
            for line in text.splitlines():
                line_s = line.strip()
                if not line_s:
                    continue
                # Check Markdown table row: | 1 | 原文 | 修改后 | or | 原文 | 修改后 |
                if line_s.startswith("|") and line_s.endswith("|"):
                    cells = [c.strip() for c in line_s.strip("|").split("|")]
                    if all(re.fullmatch(r"[-:]+", c) for c in cells if c):
                        continue
                    if any(h in cells[0] for h in ("序号", "原文", "原条款", "旧", "修改前")):
                        continue
                    if len(cells) >= 3 and re.fullmatch(r"\d+", cells[0]):
                        old_v, new_v = _clean_replacement_text(cells[1]), _clean_replacement_text(cells[2])
                    elif len(cells) >= 2:
                        old_v, new_v = _clean_replacement_text(cells[0]), _clean_replacement_text(cells[1])
                    else:
                        continue
                    if old_v and old_v != new_v and "…" not in old_v and "..." not in old_v:
                        # Handle slash-separated synonyms in table like "专家咨询费 / 专家费" -> "业务咨询服务费 / 咨询服务费"
                        if "/" in old_v and "/" in new_v:
                            old_parts = [p.strip() for p in old_v.split("/") if p.strip()]
                            new_parts = [p.strip() for p in new_v.split("/") if p.strip()]
                            for idx, op in enumerate(old_parts):
                                np = new_parts[idx] if idx < len(new_parts) else new_parts[-1]
                                if op and op != np:
                                    result[op] = np
                        else:
                            result[old_v] = new_v
                    continue

                # Check arrow format: old -> new or old => new or old → new
                for sep in ("->", "=>", "→", "替换为", "改为"):
                    if sep in line_s:
                        left, right = line_s.split(sep, 1)
                        left = re.sub(r"^[\d\.\-\*\s]+", "", left)
                        old_v, new_v = _clean_replacement_text(left), _clean_replacement_text(right)
                        if old_v and old_v != new_v:
                            result[old_v] = new_v
                        break
            return result

    result = {}
    if isinstance(parsed, dict):
        if "replacements" in parsed and isinstance(parsed["replacements"], (list, dict)):
            return parse_replacements(parsed["replacements"])
        # Check if single item dict {"old": "...", "new": "..."}
        old_keys = ("old", "original", "from", "target", "search", "原文", "原文本", "原条款", "修改前")
        new_keys = ("new", "revised", "to", "replacement", "replace", "修改后", "新文本", "建议条款", "替换为")
        if any(k in parsed for k in old_keys) and any(k in parsed for k in new_keys):
            old_v = next((parsed[k] for k in old_keys if k in parsed), "")
            new_v = next((parsed[k] for k in new_keys if k in parsed), "")
            old_c, new_c = _clean_replacement_text(old_v), _clean_replacement_text(new_v)
            if old_c and old_c != new_c:
                result[old_c] = new_c
            return result

        for k, v in parsed.items():
            if isinstance(v, (str, int, float)):
                old_c, new_c = _clean_replacement_text(k), _clean_replacement_text(v)
                if old_c and old_c != new_c and "…" not in old_c and "..." not in old_c:
                    result[old_c] = new_c
    elif isinstance(parsed, list):
        for item in parsed:
            if isinstance(item, dict):
                sub = parse_replacements(item)
                result.update(sub)
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                old_c, new_c = _clean_replacement_text(item[0]), _clean_replacement_text(item[1])
                if old_c and old_c != new_c:
                    result[old_c] = new_c
            elif isinstance(item, str):
                result.update(parse_replacements(item))

    return result


def apply_track_changes_to_doc(
    doc: docx.document.Document,
    replacements: dict[str, str],
    author: str = "AI法务审阅",
) -> tuple[int, list[dict[str, Any]]]:
    """
    Apply longest-first replacements across all paragraphs and tables in `doc`
    using character-level diffing (<w:del> and <w:ins>) while preserving run formatting.
    """
    clean_repls = {k: v for k, v in replacements.items() if k and k != v}
    if not clean_repls:
        return 0, []

    sorted_keys = sorted(clean_repls.keys(), key=len, reverse=True)
    pattern = re.compile("|".join(re.escape(k) for k in sorted_keys))
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = 100
    total_changes = 0
    change_counts: dict[tuple[str, str], int] = {}

    for p in iter_all_paragraphs(doc):
        p_el = p._p
        normalize_paragraph_runs(p_el)
        children = list(p_el)
        i = 0
        while i < len(children):
            group = []
            j = i
            while j < len(children):
                ch = children[j]
                if ch.tag == w("r"):
                    non_rpr = [c for c in ch if c.tag != w("rPr")]
                    if len(non_rpr) == 1 and non_rpr[0].tag == w("t"):
                        group.append(ch)
                        j += 1
                        continue
                break
            if not group:
                i += 1
                continue

            full_text = ""
            char_rpr = []
            for r in group:
                rPr = r.find(w("rPr"))
                t_el = r.find(w("t"))
                txt = t_el.text or ""
                full_text += txt
                char_rpr.extend([rPr] * len(txt))

            matches = list(pattern.finditer(full_text))
            if not matches:
                i = j
                continue

            new_elements = []
            cursor = 0

            def emit_plain_slice(start_idx: int, end_idx: int):
                if start_idx >= end_idx:
                    return
                seg_start = start_idx
                for k in range(start_idx + 1, end_idx + 1):
                    if k == end_idx or char_rpr[k] is not char_rpr[seg_start]:
                        new_elements.append(
                            make_text_run(full_text[seg_start:k], char_rpr[seg_start], is_del=False)
                        )
                        seg_start = k

            for m in matches:
                m_start, m_end = m.start(), m.end()
                emit_plain_slice(cursor, m_start)
                old_str = m.group(0)
                new_str = clean_repls[old_str]
                base_rpr = char_rpr[m_start] if m_start < len(char_rpr) else None
                total_changes += 1
                pair = (old_str, new_str)
                change_counts[pair] = change_counts.get(pair, 0) + 1

                for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, old_str, new_str).get_opcodes():
                    if tag == "equal":
                        emit_plain_slice(m_start + i1, m_start + i2)
                    else:
                        if tag in ("delete", "replace") and i2 > i1:
                            del_el = etree.Element(
                                w("del"),
                                {
                                    w("id"): str(rev_id),
                                    w("author"): author,
                                    w("date"): now_iso,
                                },
                            )
                            rev_id += 1
                            del_el.append(make_text_run(old_str[i1:i2], base_rpr, is_del=True))
                            new_elements.append(del_el)
                        if tag in ("insert", "replace") and j2 > j1:
                            ins_el = etree.Element(
                                w("ins"),
                                {
                                    w("id"): str(rev_id),
                                    w("author"): author,
                                    w("date"): now_iso,
                                },
                            )
                            rev_id += 1
                            ins_el.append(make_text_run(new_str[j1:j2], base_rpr, is_del=False))
                            new_elements.append(ins_el)
                cursor = m_end

            emit_plain_slice(cursor, len(full_text))

            insert_pos = list(p_el).index(group[0])
            for r in group:
                p_el.remove(r)
            for offset, el in enumerate(new_elements):
                p_el.insert(insert_pos + offset, el)

            children = list(p_el)
            i = insert_pos + len(new_elements)

    enable_word_track_changes_settings(doc)
    summary = [
        {"old": k[0], "new": k[1], "count": c}
        for k, c in change_counts.items()
    ]
    return total_changes, summary


# Regex for inline track changes markup in newly written documents:
# Matches:
# 1) [-deleted-]{+inserted+}
# 2) [-deleted-]
# 3) {+inserted+}
# 4) ~~deleted~~++inserted++
# 5) ~~deleted~~
# 6) ++inserted++
# 7) <del>deleted</del><ins>inserted</ins>
# 8) <del>deleted</del>
# 9) <ins>inserted</ins>
INLINE_REVISION_RE = re.compile(
    r"\[-(?P<del1>[\s\S]*?)-\]\{\+(?P<ins1>[\s\S]*?)\+\}"
    r"|\[-(?P<del2>[\s\S]*?)-\]"
    r"|\{\+(?P<ins2>[\s\S]*?)\+\}"
    r"|~~(?P<del3>[\s\S]+?)~~\+\+(?P<ins3>[\s\S]+?)\+\+"
    r"|~~(?P<del4>[\s\S]+?)~~"
    r"|\+\+(?P<ins4>[\s\S]+?)\+\+"
    r"|<del>(?P<del5>[\s\S]*?)</del>\s*<ins>(?P<ins5>[\s\S]*?)</ins>"
    r"|<del>(?P<del6>[\s\S]*?)</del>"
    r"|<ins>(?P<ins6>[\s\S]*?)</ins>"
)


def _build_rpr(font_name: str = "宋体", size_half_pt: int = 24, bold: bool = False) -> etree._Element:
    rPr = etree.Element(w("rPr"))
    rFonts = etree.SubElement(rPr, w("rFonts"))
    rFonts.set(w("ascii"), font_name)
    rFonts.set(w("hAnsi"), font_name)
    rFonts.set(w("eastAsia"), font_name)
    rFonts.set(w("cs"), font_name)
    sz = etree.SubElement(rPr, w("sz"))
    sz.set(w("val"), str(size_half_pt))
    szCs = etree.SubElement(rPr, w("szCs"))
    szCs.set(w("val"), str(size_half_pt))
    if bold:
        etree.SubElement(rPr, w("b"))
    return rPr


def append_rich_text_with_revisions(
    p_el: etree._Element,
    text: str,
    author: str = "AI法务审阅",
    font_name: str = "宋体",
    size_half_pt: int = 24,
    default_bold: bool = False,
    rev_counter: Optional[list[int]] = None,
) -> int:
    """
    Append text into paragraph element `p_el`, supporting both **bold** and inline Track Changes syntax.
    Returns number of revision marks emitted.
    """
    if rev_counter is None:
        rev_counter = [200]
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    revisions_emitted = 0

    def _emit_plain_with_bold(segment: str):
        if not segment:
            return
        parts = re.split(r"(\*\*.*?\*\*)", segment)
        for part in parts:
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                rPr = _build_rpr(font_name=font_name, size_half_pt=size_half_pt, bold=True)
                p_el.append(make_text_run(part[2:-2], rPr, is_del=False))
            else:
                rPr = _build_rpr(font_name=font_name, size_half_pt=size_half_pt, bold=default_bold)
                p_el.append(make_text_run(part, rPr, is_del=False))

    cursor = 0
    for m in INLINE_REVISION_RE.finditer(text):
        _emit_plain_with_bold(text[cursor : m.start()])
        gd = m.groupdict()
        del_txt = next((gd[k] for k in ("del1", "del2", "del3", "del4", "del5", "del6") if gd.get(k) is not None), None)
        ins_txt = next((gd[k] for k in ("ins1", "ins2", "ins3", "ins4", "ins5", "ins6") if gd.get(k) is not None), None)
        rPr = _build_rpr(font_name=font_name, size_half_pt=size_half_pt, bold=default_bold)

        if del_txt:
            del_el = etree.Element(
                w("del"),
                {
                    w("id"): str(rev_counter[0]),
                    w("author"): author,
                    w("date"): now_iso,
                },
            )
            rev_counter[0] += 1
            del_el.append(make_text_run(_clean_replacement_text(del_txt), rPr, is_del=True))
            p_el.append(del_el)
            revisions_emitted += 1

        if ins_txt:
            ins_el = etree.Element(
                w("ins"),
                {
                    w("id"): str(rev_counter[0]),
                    w("author"): author,
                    w("date"): now_iso,
                },
            )
            rev_counter[0] += 1
            ins_el.append(make_text_run(_clean_replacement_text(ins_txt), rPr, is_del=False))
            p_el.append(ins_el)
            revisions_emitted += 1

        cursor = m.end()

    _emit_plain_with_bold(text[cursor:])
    return revisions_emitted


def extract_full_text_from_doc_path(file_path: str) -> str:
    """Read all paragraphs and tables from a .doc or .docx file in document order."""
    ext = os.path.splitext(file_path)[1].lower()
    with tempfile.TemporaryDirectory(prefix="doc_read_") as tmp:
        if ext == ".doc":
            docx_path = run_soffice_convert(file_path, "docx", tmp)
        else:
            docx_path = file_path

        doc = docx.Document(docx_path)
        lines = []
        for idx, p in enumerate(doc.paragraphs):
            txt = p.text.strip()
            if txt:
                lines.append(f"[段落 {idx + 1}] {txt}")

        for t_idx, table in enumerate(doc.tables):
            lines.append(f"\n--- [表格 {t_idx + 1}] ---")
            for row in table.rows:
                row_cells = [ " ".join(c.text.split()) for c in row.cells ]
                lines.append(" | ".join(row_cells))

        return "\n".join(lines)


def resolve_uploaded_doc_file(
    file_id: Optional[str] = None,
    files_meta: Optional[list[dict]] = None,
    messages: Optional[list[dict]] = None,
    chat_id: Optional[str] = None,
    user_id: Optional[str] = None,
) -> Optional[dict[str, Any]]:
    """
    Find the target .doc / .docx file record from:
    1. Explicit file_id
    2. Current turn's __files__ metadata
    3. Chat record in DB (chat_file or chat.files)
    4. User's most recent uploaded .doc / .docx file in DB
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        candidate_ids = []
        if file_id:
            clean_id = file_id.strip()
            if "/" in clean_id:
                parts = [p for p in clean_id.split("/") if p]
                for p in parts:
                    if re.fullmatch(r"[0-9a-fA-F-]{36}", p):
                        clean_id = p
                        break
            candidate_ids.append(clean_id)

        for f in files_meta or []:
            fid = f.get("id") or (f.get("file") or {}).get("id")
            if fid and fid not in candidate_ids:
                candidate_ids.append(fid)

        if chat_id:
            row = cur.execute("SELECT chat FROM chat WHERE id = ?", (chat_id,)).fetchone()
            if row and row[0]:
                try:
                    chat_obj = json.loads(row[0])
                    for f in chat_obj.get("files", []):
                        fid = f.get("id")
                        if fid and fid not in candidate_ids:
                            candidate_ids.append(fid)
                    msgs = (chat_obj.get("history") or {}).get("messages") or {}
                    for m in msgs.values():
                        for f in m.get("files") or []:
                            fid = f.get("id")
                            if fid and fid not in candidate_ids:
                                candidate_ids.append(fid)
                except Exception:
                    pass

        for cid in candidate_ids:
            row = cur.execute(
                "SELECT id, user_id, filename, path, meta, data FROM file WHERE id = ?",
                (cid,),
            ).fetchone()
            if row and row[3] and os.path.exists(row[3]):
                fname = (row[2] or "").lower()
                if fname.endswith((".doc", ".docx")) and "修订版" not in fname:
                    return {
                        "id": row[0],
                        "user_id": row[1],
                        "filename": row[2],
                        "path": row[3],
                        "meta": json.loads(row[4]) if row[4] else {},
                        "data": json.loads(row[5]) if row[5] else {},
                    }

        # Fallback: latest uploaded original .doc / .docx file in DB
        rows = cur.execute(
            "SELECT id, user_id, filename, path, meta, data FROM file ORDER BY created_at DESC"
        ).fetchall()
        for row in rows:
            fname = row[2] or ""
            fpath = row[3] or ""
            if (
                fname.lower().endswith((".doc", ".docx"))
                and "修订版" not in fname
                and "审阅版" not in fname
                and os.path.exists(fpath)
            ):
                return {
                    "id": row[0],
                    "user_id": row[1],
                    "filename": row[2],
                    "path": row[3],
                    "meta": json.loads(row[4]) if row[4] else {},
                    "data": json.loads(row[5]) if row[5] else {},
                }
        # If only revised files exist, allow the latest .doc/.docx
        for row in rows:
            fname = row[2] or ""
            fpath = row[3] or ""
            if fname.lower().endswith((".doc", ".docx")) and os.path.exists(fpath):
                return {
                    "id": row[0],
                    "user_id": row[1],
                    "filename": row[2],
                    "path": row[3],
                    "meta": json.loads(row[4]) if row[4] else {},
                    "data": json.loads(row[5]) if row[5] else {},
                }
        return None
    finally:
        conn.close()


def register_file_in_open_webui(
    source_file_path: str,
    display_filename: str,
    user_id: Optional[str] = None,
    extracted_text: str = "",
) -> dict[str, Any]:
    """Copy a generated .doc or .docx file to UPLOAD_DIR and insert it into Open WebUI's `file` table."""
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    file_id = str(uuid.uuid4())
    safe_name = os.path.basename(display_filename)
    dest_path = os.path.join(UPLOAD_DIR, f"{file_id}_{safe_name}")
    shutil.copy2(source_file_path, dest_path)

    size = os.path.getsize(dest_path)
    with open(dest_path, "rb") as f:
        file_hash = hashlib.sha256(f.read()).hexdigest()

    ext = os.path.splitext(safe_name)[1].lower()
    if ext == ".doc":
        content_type = "application/msword"
    elif ext == ".docx":
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        content_type = "application/octet-stream"

    now_ts = int(time.time())
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        if not user_id:
            row = cur.execute("SELECT id FROM user ORDER BY created_at ASC LIMIT 1").fetchone()
            user_id = row[0] if row else "system"

        meta_dict = {
            "name": safe_name,
            "content_type": content_type,
            "size": size,
            "file_hash": file_hash,
            "data": {},
        }
        data_dict = {
            "status": "completed",
            "content": extracted_text,
            "error": None,
        }
        cur.execute(
            """
            INSERT INTO file (id, user_id, hash, filename, path, data, meta, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                file_id,
                user_id,
                file_hash,
                safe_name,
                dest_path,
                json.dumps(data_dict, ensure_ascii=False),
                json.dumps(meta_dict, ensure_ascii=False),
                now_ts,
                now_ts,
            ),
        )
        conn.commit()
    finally:
        conn.close()

    download_url = f"/api/v1/files/{file_id}/content?attachment=true"
    return {
        "id": file_id,
        "filename": safe_name,
        "path": dest_path,
        "size": size,
        "content_type": content_type,
        "download_url": download_url,
    }


def revise_document_file(
    source_path: str,
    original_filename: str,
    raw_replacements: Any,
    author: str = "AI法务审阅",
    output_filename_stem: Optional[str] = None,
    user_id: Optional[str] = None,
) -> dict[str, Any]:
    """
    Convert source (.doc or .docx) -> .docx, apply Track Changes for all replacements,
    and save + register both .doc and .docx in Open WebUI.
    """
    replacements = parse_replacements(raw_replacements)
    if not replacements:
        raise ValueError("未解析到有效的修改替换项（replacements 为空）。请提供需要替换的『原文 -> 修改后文本』映射。")

    stem = output_filename_stem or os.path.splitext(os.path.basename(original_filename))[0]
    stem = re.sub(r"^[0-9a-fA-F-]{36}_", "", stem)
    if not stem.endswith(("修订版", "审阅版")):
        stem = f"{stem}_修订版"

    with tempfile.TemporaryDirectory(prefix="doc_rev_") as tmp:
        ext = os.path.splitext(source_path)[1].lower()
        if ext == ".doc":
            working_docx = run_soffice_convert(source_path, "docx", tmp)
        else:
            working_docx = os.path.join(tmp, "working.docx")
            shutil.copy2(source_path, working_docx)

        doc = docx.Document(working_docx)
        total_changes, change_summary = apply_track_changes_to_doc(doc, replacements, author=author)

        out_docx_tmp = os.path.join(tmp, f"{stem}.docx")
        doc.save(out_docx_tmp)
        out_doc_tmp = run_soffice_convert(out_docx_tmp, "doc", tmp)

        doc_info = register_file_in_open_webui(
            out_doc_tmp,
            f"{stem}.doc",
            user_id=user_id,
        )
        docx_info = register_file_in_open_webui(
            out_docx_tmp,
            f"{stem}.docx",
            user_id=user_id,
        )

    md_links = (
        f"### 📄 已生成修订模式（Track Changes）Word 文档\n\n"
        f"- 📥 [**点击下载 Word 97-2003 修订版 (.doc)**：`{doc_info['filename']}`]({doc_info['download_url']})\n"
        f"- 📥 [**点击下载 Word 修订版 (.docx)**：`{docx_info['filename']}`]({docx_info['download_url']})\n\n"
        f"> **修订统计**：共完成 **{total_changes}** 处修订标记（审阅人：`{author}`）。用 Microsoft Word 或 WPS 打开文件即可在「审阅 -> 修订」模式下查看所有删除与新增痕迹，支持一键接受或拒绝修订。"
    )

    return {
        "status": "success",
        "total_changes": total_changes,
        "author": author,
        "change_summary": change_summary,
        "doc_file": doc_info,
        "docx_file": docx_info,
        "download_markdown": md_links,
    }


def create_document_file(
    title: str,
    content: str,
    filename: Optional[str] = None,
    author: str = "AI法务审阅",
    raw_replacements: Any = None,
    user_id: Optional[str] = None,
) -> dict[str, Any]:
    """
    Create a new .doc and .docx file from Markdown/structured text, supporting inline Track Changes
    ([-删除-]{+新增+}, ~~删除~~++新增++, <del>...</del><ins>...</ins>) as well as optional replacement pairs.
    """
    doc = docx.Document()

    # Configure default Normal style for Chinese legal documents
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "宋体"
    normal_style.font.size = Pt(12)
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    rev_counter = [200]
    inline_revisions_count = 0

    if title and title.strip():
        clean_title = re.sub(r"^#+\s*", "", title.strip())
        p_title = doc.add_paragraph()
        p_title.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(6)
        p_title.paragraph_format.space_after = Pt(14)
        inline_revisions_count += append_rich_text_with_revisions(
            p_title._p,
            clean_title,
            author=author,
            font_name="宋体",
            size_half_pt=32,  # 16pt
            default_bold=True,
            rev_counter=rev_counter,
        )

    lines = (content or "").splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if not stripped or stripped == "---":
            i += 1
            continue

        # Markdown Table block
        if stripped.startswith("|") and stripped.endswith("|"):
            table_rows = []
            while i < len(lines):
                r_line = lines[i].strip()
                if not (r_line.startswith("|") and r_line.endswith("|")):
                    break
                cells = [c.strip() for c in r_line.strip("|").split("|")]
                if not all(re.fullmatch(r"[-:\s]+", c) for c in cells if c):
                    table_rows.append(cells)
                i += 1

            if table_rows:
                col_count = max(len(r) for r in table_rows)
                t = doc.add_table(rows=len(table_rows), cols=col_count)
                t.style = "Table Grid"
                for r_idx, r_cells in enumerate(table_rows):
                    for c_idx in range(col_count):
                        cell_txt = r_cells[c_idx] if c_idx < len(r_cells) else ""
                        cell_p = t.cell(r_idx, c_idx).paragraphs[0]
                        inline_revisions_count += append_rich_text_with_revisions(
                            cell_p._p,
                            cell_txt,
                            author=author,
                            font_name="宋体",
                            size_half_pt=21,  # 10.5pt for tables
                            default_bold=(r_idx == 0),
                            rev_counter=rev_counter,
                        )
            continue

        # Headings
        if stripped.startswith("#"):
            m_h = re.match(r"^(#+)\s*(.*)$", stripped)
            level = len(m_h.group(1)) if m_h else 1
            h_text = m_h.group(2) if m_h else stripped
            if level == 1 and title and h_text.strip() == title.strip():
                i += 1
                continue
            p_h = doc.add_paragraph()
            p_h.paragraph_format.space_before = Pt(10)
            p_h.paragraph_format.space_after = Pt(4)
            sz_map = {1: 28, 2: 26, 3: 24}
            inline_revisions_count += append_rich_text_with_revisions(
                p_h._p,
                h_text,
                author=author,
                font_name="宋体",
                size_half_pt=sz_map.get(level, 24),
                default_bold=True,
                rev_counter=rev_counter,
            )
            i += 1
            continue

        # Bullet / quote cleanup
        if stripped.startswith("> "):
            stripped = stripped[2:].strip()

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        if not re.match(r"^([第\d一二三四五六七八九十]+[条章节、\.]|[（\(]\d+[）\)]|甲方|乙方|日期|签字|附件)", stripped):
            p.paragraph_format.first_line_indent = Pt(24)

        inline_revisions_count += append_rich_text_with_revisions(
            p._p,
            stripped,
            author=author,
            font_name="宋体",
            size_half_pt=24,
            default_bold=False,
            rev_counter=rev_counter,
        )
        i += 1

    repls = parse_replacements(raw_replacements)
    repl_changes_count = 0
    change_summary = []
    if repls:
        repl_changes_count, change_summary = apply_track_changes_to_doc(doc, repls, author=author)

    enable_word_track_changes_settings(doc)
    total_changes = inline_revisions_count + repl_changes_count

    stem = filename or title or "生成文档"
    stem = os.path.splitext(os.path.basename(stem))[0].strip()
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem) or "生成文档"

    with tempfile.TemporaryDirectory(prefix="doc_create_") as tmp:
        out_docx_tmp = os.path.join(tmp, f"{stem}.docx")
        doc.save(out_docx_tmp)
        out_doc_tmp = run_soffice_convert(out_docx_tmp, "doc", tmp)

        doc_info = register_file_in_open_webui(
            out_doc_tmp,
            f"{stem}.doc",
            user_id=user_id,
            extracted_text=content,
        )
        docx_info = register_file_in_open_webui(
            out_docx_tmp,
            f"{stem}.docx",
            user_id=user_id,
            extracted_text=content,
        )

    mode_desc = (
        f"包含 **{total_changes}** 处修订模式标记（审阅人：`{author}`）"
        if total_changes > 0
        else "已开启 Word 修订跟踪模式（Track Changes）"
    )
    md_links = (
        f"### 📄 已生成 Word 文档（支持 .doc 与 .docx）\n\n"
        f"- 📥 [**点击下载 Word 97-2003 文件 (.doc)**：`{doc_info['filename']}`]({doc_info['download_url']})\n"
        f"- 📥 [**点击下载 Word 文件 (.docx)**：`{docx_info['filename']}`]({docx_info['download_url']})\n\n"
        f"> **文档状态**：{mode_desc}。用 Microsoft Word 或 WPS 打开即可直接审阅或编辑。"
    )

    return {
        "status": "success",
        "total_changes": total_changes,
        "author": author,
        "change_summary": change_summary,
        "doc_file": doc_info,
        "docx_file": docx_info,
        "download_markdown": md_links,
    }
