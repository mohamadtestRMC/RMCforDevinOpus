"""
Base RMC template filler — in-place ZIP+XML edit.

The user wants their uploaded `1 Base RMC _ <month>.xlsx` returned with all
27 sheets populated, **preserving every visual aspect of the template**:
sheet colours, cell colours, borders, conditional formatting, defined
charts, pivot tables, etc. They explicitly asked for the template "as it
is, same colors, same column names" but with values filled in.

The previous implementation rewrote the workbook from scratch via
xlsxwriter. That was fast but stripped every cell style and lost all
pivot tables / charts. Loading the workbook with `openpyxl(read_only=False)`
preserves everything but takes >10 minutes on the 14MB template.

This implementation operates directly on the .xlsx ZIP archive:

  1. Open the template as a ZipFile and load every entry into memory.
  2. Resolve sheet name → worksheet XML path via `xl/workbook.xml`
     and `xl/_rels/workbook.xml.rels`.
  3. Parse only the worksheets we need to modify (process sheets,
     Jobtrack, RMC summary). All other entries — including
     `xl/styles.xml`, `xl/charts/*.xml`, `xl/pivotTables/*.xml`,
     themes, conditional formats, drawings — are passed through
     byte-for-byte. Nothing visual is touched.
  4. For each cell we write, we *preserve* the existing `s=`
     (style index) attribute so the cell keeps its template colour,
     number format, border, etc. We replace only the cell's value
     (and clear any existing `<f>` formula on cells we overwrite).
  5. For columns we don't compute, we leave the template's pre-filled
     formula intact — Excel recalculates on open from the data we
     wrote into the process sheets.
  6. New strings go into `xl/sharedStrings.xml` (or a new shared
     strings part if the template lacks one).
  7. Re-emit the ZIP with the modified sheet XML and shared strings;
     every other archive member is preserved.

End result: the user gets a workbook visually identical to their
template, with our computed data populating it.

All values written are derived from the user's monthly source files
at runtime via the existing UnifiedRMCPipeline. No data values are
hardcoded.
"""
from __future__ import annotations

import io
import logging
import re
import zipfile
from datetime import date, datetime, time as _time
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET

import openpyxl

logger = logging.getLogger(__name__)


# ── XML namespaces ──────────────────────────────────────────────────────────

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XML = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def _qn(tag: str) -> str:
    """Qualify a tag name with the main spreadsheetml namespace."""
    return f"{{{NS_MAIN}}}{tag}"


_QN_R = f"{{{NS_REL}}}id"
_QN_XML_SPACE = f"{{{NS_XML}}}space"


# ── Sheet groups ────────────────────────────────────────────────────────────

PROCESS_SHEET_TO_IDX: Dict[str, str] = {
    "BFL": "BFL",
    "Bag&Pouch": "Bag&Pouch",
    "PTR Rew": "PTR Rew",
    "Slit": "Slit",
    "Lam": "Lam",
    "Print": "Print",
    "Printing Work": "Print",
    "HCI Rew": "HCI Rew",
    "Spout&Valve": "Spout&Valve",
    "Embossing": "Embossing",
}

# Header / data row layout (1-based Excel rows)
PROCESS_HEADER_ROW = 6
PROCESS_DATA_ROW = 7

JOBTRACK_HEADER_ROW = 4
JOBTRACK_DATA_ROW = 5

RMC_HEADER_ROW = 6
RMC_DATA_ROW = 7

_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


# ── Helpers ─────────────────────────────────────────────────────────────────

def _norm_header(s: Any) -> str:
    """Normalize a header label for fuzzy matching across templates."""
    if s is None:
        return ""
    s = str(s).lower()
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.rstrip(":.")
    return s


def _excel_col_letter(col_1based: int) -> str:
    """1 → 'A', 27 → 'AA', etc."""
    s = ""
    n = col_1based
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


_CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _parse_cell_ref(ref: str) -> Tuple[int, int]:
    """'C7' → (col=3, row=7), both 1-based."""
    m = _CELL_REF_RE.match(ref)
    if not m:
        return 0, 0
    letters, digits = m.group(1), m.group(2)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return col, int(digits)


def _detect_report_month_from_jt(jt_headers: List[Any], jt_rows: List[List[Any]]) -> str:
    """Pick the modal (year, month) from the Jobtrack date column."""
    if not jt_rows:
        return ""
    date_ci = -1
    for ci, h in enumerate(jt_headers):
        if "date" in str(h).lower():
            date_ci = ci
            break
    if date_ci < 0:
        return ""
    counts: Dict[Tuple[int, int], int] = {}
    for row in jt_rows:
        if date_ci >= len(row):
            continue
        v = row[date_ci]
        if isinstance(v, (datetime, date)):
            key = (v.year, v.month)
            counts[key] = counts.get(key, 0) + 1
    if not counts:
        return ""
    (yr, mo), _n = max(counts.items(), key=lambda kv: kv[1])
    return f"{_MONTH_NAMES[mo - 1][:3]} {yr}"


# ── XLSX in-place editor ────────────────────────────────────────────────────

class XLSXEditor:
    """Edit cell values inside an .xlsx archive while preserving formatting.

    The archive is read into memory, only the sheet XMLs we touch are
    parsed, and the rest of the bytes pass through verbatim. This keeps
    cell styles (colour, borders, number formats), conditional formats,
    pivot tables, charts, defined names, drawings, etc., all intact.
    """

    def __init__(self, template_bytes: bytes):
        self._files: Dict[str, bytes] = {}
        with zipfile.ZipFile(io.BytesIO(template_bytes), "r") as z:
            self._file_order: List[str] = list(z.namelist())
            for name in self._file_order:
                self._files[name] = z.read(name)

        # Parse workbook.xml + workbook rels to map sheet name → XML path
        wb_root = ET.fromstring(self._files["xl/workbook.xml"])
        sheets_elem = wb_root.find(_qn("sheets"))
        if sheets_elem is None:
            raise ValueError("workbook.xml has no <sheets> element")

        self._sheet_meta: List[Dict[str, str]] = []
        for s in sheets_elem.findall(_qn("sheet")):
            self._sheet_meta.append({
                "name": s.get("name", ""),
                "rid": s.get(_QN_R, ""),
            })

        rels_root = ET.fromstring(self._files["xl/_rels/workbook.xml.rels"])
        rid_to_target: Dict[str, str] = {}
        for rel in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship"):
            rid_to_target[rel.get("Id", "")] = rel.get("Target", "")

        for sm in self._sheet_meta:
            target = rid_to_target.get(sm["rid"], "")
            if target.startswith("/"):
                sm["path"] = target.lstrip("/")
            else:
                sm["path"] = "xl/" + target

        self._sheet_by_name: Dict[str, Dict[str, str]] = {sm["name"]: sm for sm in self._sheet_meta}

        # Lazily-parsed worksheets
        self._sheet_roots: Dict[str, ET.Element] = {}
        # Sheet name → row_number_1based → row Element (cache, built when sheet is parsed)
        self._row_cache: Dict[str, Dict[int, ET.Element]] = {}

        # Shared strings table
        ss_path = "xl/sharedStrings.xml"
        self._ss_path = ss_path
        if ss_path in self._files:
            try:
                self._ss_root = ET.fromstring(self._files[ss_path])
            except ET.ParseError:
                self._ss_root = ET.Element(_qn("sst"))
        else:
            self._ss_root = ET.Element(_qn("sst"))
        # Map text → first index (only single-text-run shared strings can be reused)
        self._ss_index: Dict[str, int] = {}
        for i, si in enumerate(self._ss_root.findall(_qn("si"))):
            t = si.find(_qn("t"))
            if t is not None and t.text is not None:
                # Skip rich-text runs (they have multiple <r> children)
                if not si.findall(_qn("r")):
                    self._ss_index.setdefault(t.text, i)

    # ── public API ─────────────────────────────────────────────────────

    @property
    def sheet_names(self) -> List[str]:
        return [sm["name"] for sm in self._sheet_meta]

    def has_sheet(self, name: str) -> bool:
        return name in self._sheet_by_name

    def get_sheet_root(self, name: str) -> ET.Element:
        if name not in self._sheet_roots:
            sm = self._sheet_by_name[name]
            self._sheet_roots[name] = ET.fromstring(self._files[sm["path"]])
            sd = self._sheet_roots[name].find(_qn("sheetData"))
            self._row_cache[name] = {}
            if sd is not None:
                for row_el in sd.findall(_qn("row")):
                    r_attr = row_el.get("r")
                    if r_attr:
                        try:
                            self._row_cache[name][int(r_attr)] = row_el
                        except ValueError:
                            pass
        return self._sheet_roots[name]

    def read_header_row(self, sheet_name: str, header_row_1based: int) -> List[Optional[str]]:
        """Read a header row's cell values from the template."""
        if not self.has_sheet(sheet_name):
            return []
        self.get_sheet_root(sheet_name)
        row_el = self._row_cache[sheet_name].get(header_row_1based)
        if row_el is None:
            return []
        out: Dict[int, Optional[str]] = {}
        max_col = 0
        si_list = self._ss_root.findall(_qn("si"))
        for c_el in row_el.findall(_qn("c")):
            ref = c_el.get("r", "")
            col_1, _row_1 = _parse_cell_ref(ref)
            if col_1 == 0:
                continue
            t_attr = c_el.get("t")
            v_el = c_el.find(_qn("v"))
            is_el = c_el.find(_qn("is"))
            value: Optional[str] = None
            if t_attr == "s" and v_el is not None and v_el.text:
                try:
                    idx = int(v_el.text)
                    if 0 <= idx < len(si_list):
                        t = si_list[idx].find(_qn("t"))
                        if t is not None and t.text is not None:
                            value = t.text
                        else:
                            parts = []
                            for r in si_list[idx].findall(_qn("r")):
                                rt = r.find(_qn("t"))
                                if rt is not None and rt.text:
                                    parts.append(rt.text)
                            value = "".join(parts) if parts else None
                except ValueError:
                    pass
            elif t_attr == "inlineStr" and is_el is not None:
                t = is_el.find(_qn("t"))
                if t is not None and t.text is not None:
                    value = t.text
            elif v_el is not None and v_el.text:
                value = v_el.text
            out[col_1] = value
            if col_1 > max_col:
                max_col = col_1
        return [out.get(c) for c in range(1, max_col + 1)]

    def write_cell(self, sheet_name: str, row_1based: int, col_1based: int,
                   value: Any) -> None:
        """Write a value to a cell, preserving its style index and clearing any formula."""
        if not self.has_sheet(sheet_name):
            return
        if value is None or value == "":
            return

        root = self.get_sheet_root(sheet_name)
        sd = root.find(_qn("sheetData"))
        if sd is None:
            return

        row_cache = self._row_cache[sheet_name]
        row_el = row_cache.get(row_1based)
        if row_el is None:
            row_el = ET.Element(_qn("row"))
            row_el.set("r", str(row_1based))
            inserted = False
            for i, existing in enumerate(list(sd)):
                if existing.tag != _qn("row"):
                    continue
                try:
                    existing_r = int(existing.get("r", "0"))
                except ValueError:
                    continue
                if existing_r > row_1based:
                    sd.insert(list(sd).index(existing), row_el)
                    inserted = True
                    break
            if not inserted:
                sd.append(row_el)
            row_cache[row_1based] = row_el

        cell_ref = _excel_col_letter(col_1based) + str(row_1based)
        cell_el = None
        for c in row_el.findall(_qn("c")):
            if c.get("r") == cell_ref:
                cell_el = c
                break
        if cell_el is None:
            cell_el = ET.Element(_qn("c"))
            cell_el.set("r", cell_ref)
            inserted = False
            for i, existing in enumerate(list(row_el)):
                ref = existing.get("r", "")
                col_1, _ = _parse_cell_ref(ref)
                if col_1 > col_1based:
                    row_el.insert(i, cell_el)
                    inserted = True
                    break
            if not inserted:
                row_el.append(cell_el)

        # Preserve 's' (style index) — only clear value/formula/type
        for child in list(cell_el):
            cell_el.remove(child)
        cell_el.attrib.pop("t", None)

        # Encode the value
        if isinstance(value, bool):
            cell_el.set("t", "b")
            v = ET.SubElement(cell_el, _qn("v"))
            v.text = "1" if value else "0"
        elif isinstance(value, (int, float)):
            f = float(value)
            if f != f:  # NaN
                return
            v = ET.SubElement(cell_el, _qn("v"))
            v.text = repr(f) if isinstance(value, float) else str(int(value))
        elif isinstance(value, datetime):
            serial = (value - datetime(1899, 12, 30)).total_seconds() / 86400.0
            v = ET.SubElement(cell_el, _qn("v"))
            v.text = repr(serial)
        elif isinstance(value, date):
            dt = datetime(value.year, value.month, value.day)
            serial = (dt - datetime(1899, 12, 30)).total_seconds() / 86400.0
            v = ET.SubElement(cell_el, _qn("v"))
            v.text = repr(serial)
        elif isinstance(value, _time):
            cell_el.set("t", "inlineStr")
            is_el = ET.SubElement(cell_el, _qn("is"))
            t = ET.SubElement(is_el, _qn("t"))
            t.text = value.isoformat()
        elif isinstance(value, str) and value.startswith("=") and len(value) <= 8000:
            f_el = ET.SubElement(cell_el, _qn("f"))
            f_el.text = value[1:]
        else:
            s = str(value)
            idx = self._add_string(s)
            cell_el.set("t", "s")
            v = ET.SubElement(cell_el, _qn("v"))
            v.text = str(idx)

    def write_string_inline(self, sheet_name: str, row_1based: int,
                            col_1based: int, value: str) -> None:
        """Write a string as inline (avoid touching shared strings)."""
        if not self.has_sheet(sheet_name) or value is None:
            return
        # Make sure cell exists with a clean payload
        self.write_cell(sheet_name, row_1based, col_1based, " ")
        row_el = self._row_cache[sheet_name].get(row_1based)
        if row_el is None:
            return
        cell_ref = _excel_col_letter(col_1based) + str(row_1based)
        cell_el = next((c for c in row_el.findall(_qn("c")) if c.get("r") == cell_ref), None)
        if cell_el is None:
            return
        for child in list(cell_el):
            cell_el.remove(child)
        cell_el.attrib.pop("t", None)
        cell_el.set("t", "inlineStr")
        is_el = ET.SubElement(cell_el, _qn("is"))
        t = ET.SubElement(is_el, _qn("t"))
        if value and (value[:1] == " " or value[-1:] == " "):
            t.set(_QN_XML_SPACE, "preserve")
        t.text = value

    def clear_cell_formula(self, sheet_name: str, row_1based: int, col_1based: int) -> None:
        """Remove any <f> formula element from a cell, leaving everything else intact."""
        if not self.has_sheet(sheet_name):
            return
        self.get_sheet_root(sheet_name)
        row_el = self._row_cache[sheet_name].get(row_1based)
        if row_el is None:
            return
        cell_ref = _excel_col_letter(col_1based) + str(row_1based)
        cell_el = next((c for c in row_el.findall(_qn("c")) if c.get("r") == cell_ref), None)
        if cell_el is None:
            return
        for f in cell_el.findall(_qn("f")):
            cell_el.remove(f)

    def save(self) -> bytes:
        """Re-emit the .xlsx archive with our modifications."""
        # Re-serialize modified sheets
        for name, root in self._sheet_roots.items():
            sm = self._sheet_by_name[name]
            self._files[sm["path"]] = self._serialize(root)

        # Update shared strings count and serialize
        if self._ss_index or self._ss_path in self._files:
            self._ss_root.set("count", str(len(self._ss_index)))
            self._ss_root.set("uniqueCount", str(len(self._ss_index)))
            new_ss = self._ss_path not in self._files
            self._files[self._ss_path] = self._serialize(self._ss_root)
            if new_ss:
                self._file_order.append(self._ss_path)
                self._ensure_shared_strings_relationship()

        # Repack ZIP, preserving original entry order
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for name in self._file_order:
                if name in self._files and name not in written:
                    zout.writestr(name, self._files[name])
                    written.add(name)
            for name, data in self._files.items():
                if name not in written:
                    zout.writestr(name, data)
                    written.add(name)
        return out.getvalue()

    # ── internals ──────────────────────────────────────────────────────

    def _add_string(self, s: str) -> int:
        if s in self._ss_index:
            return self._ss_index[s]
        idx = len(self._ss_root.findall(_qn("si")))
        si = ET.SubElement(self._ss_root, _qn("si"))
        t = ET.SubElement(si, _qn("t"))
        if s and (s[:1] == " " or s[-1:] == " "):
            t.set(_QN_XML_SPACE, "preserve")
        t.text = s
        self._ss_index[s] = idx
        return idx

    def _ensure_shared_strings_relationship(self) -> None:
        """If we created a new sharedStrings.xml, register it in workbook rels and content types."""
        rels_path = "xl/_rels/workbook.xml.rels"
        rels_root = ET.fromstring(self._files[rels_path])
        has_ss_rel = any(
            rel.get("Type", "").endswith("/sharedStrings")
            for rel in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship")
        )
        if not has_ss_rel:
            existing_ids = {rel.get("Id", "") for rel in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship")}
            i = 1
            while f"rId_ss_added{i}" in existing_ids:
                i += 1
            rel = ET.SubElement(rels_root, f"{{{NS_PKG_REL}}}Relationship")
            rel.set("Id", f"rId_ss_added{i}")
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
            rel.set("Target", "sharedStrings.xml")
            self._files[rels_path] = self._serialize(rels_root)

        ct_path = "[Content_Types].xml"
        if ct_path in self._files:
            ct_root = ET.fromstring(self._files[ct_path])
            ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"
            has_ss_ct = any(
                ov.get("PartName", "") == "/xl/sharedStrings.xml"
                for ov in ct_root.findall(f"{{{ns_ct}}}Override")
            )
            if not has_ss_ct:
                ov = ET.SubElement(ct_root, f"{{{ns_ct}}}Override")
                ov.set("PartName", "/xl/sharedStrings.xml")
                ov.set("ContentType",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
                self._files[ct_path] = self._serialize(ct_root)

    @staticmethod
    def _serialize(root: ET.Element) -> bytes:
        buf = io.BytesIO()
        tree = ET.ElementTree(root)
        tree.write(buf, xml_declaration=True, encoding="UTF-8")
        data = buf.getvalue()
        # Force standalone="yes" — Excel sometimes complains if missing.
        if data.startswith(b"<?xml") and b"standalone=" not in data[:200]:
            data = data.replace(b"<?xml version='1.0' encoding='UTF-8'?>",
                                b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", 1)
            data = data.replace(b'<?xml version="1.0" encoding="UTF-8"?>',
                                b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", 1)
        return data


# ── Enriched JT reader ──────────────────────────────────────────────────────

def _read_enriched_jobtrack(jt_bytes: bytes) -> Tuple[List[Any], List[List[Any]]]:
    """Return (headers_at_row_4, data_rows starting at row 5)."""
    wb = openpyxl.load_workbook(io.BytesIO(jt_bytes),
                                data_only=True, read_only=True,
                                keep_links=False)
    ws = wb.active
    headers: List[Any] = []
    data: List[List[Any]] = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row is None:
            continue
        if r_idx < JOBTRACK_HEADER_ROW:
            continue
        elif r_idx == JOBTRACK_HEADER_ROW:
            headers = list(row)
        else:
            if any(v is not None and v != "" for v in row):
                data.append(list(row))
    wb.close()
    return headers, data


# ── Fill orchestration ──────────────────────────────────────────────────────

def _build_col_map(engine_headers: List[Any],
                   template_headers: List[Optional[str]]) -> List[int]:
    """Map each engine column → template column 1-based (or -1 if no match)."""
    norm_to_col1: Dict[str, int] = {}
    for ci_0, h in enumerate(template_headers):
        nh = _norm_header(h)
        if nh and nh not in norm_to_col1:
            norm_to_col1[nh] = ci_0 + 1
    out: List[int] = []
    for h in engine_headers:
        out.append(norm_to_col1.get(_norm_header(h), -1))
    return out


def _fill_process_sheet(editor: XLSXEditor, sheet_name: str,
                        engine_headers: List[Any],
                        engine_rows: List[tuple]) -> int:
    """Write engine rows starting at row 7. Returns row count written."""
    if not editor.has_sheet(sheet_name):
        return 0
    template_headers = editor.read_header_row(sheet_name, PROCESS_HEADER_ROW)
    if not template_headers:
        return 0
    col_map = _build_col_map(engine_headers, template_headers)

    written = 0
    for ri, row_tuple in enumerate(engine_rows):
        target_row = PROCESS_DATA_ROW + ri
        for src_ci, val in enumerate(row_tuple):
            if src_ci >= len(col_map):
                break
            tgt_ci = col_map[src_ci]
            if tgt_ci < 0 or val is None or val == "":
                continue
            editor.write_cell(sheet_name, target_row, tgt_ci, val)
        written += 1
    return written


def _fill_jobtrack(editor: XLSXEditor, jt_headers: List[Any],
                   jt_rows: List[List[Any]]) -> int:
    """Write enriched Jobtrack rows starting at row 5."""
    if not editor.has_sheet("Jobtrack"):
        return 0
    template_headers = editor.read_header_row("Jobtrack", JOBTRACK_HEADER_ROW)
    if not template_headers:
        col_map = list(range(1, len(jt_headers) + 1))
    else:
        col_map = _build_col_map(jt_headers, template_headers)
        # Fallback to 1:1 column index for engine columns the template lacks
        for i, m in enumerate(col_map):
            if m < 0:
                col_map[i] = i + 1

    written = 0
    for ri, row in enumerate(jt_rows):
        target_row = JOBTRACK_DATA_ROW + ri
        for src_ci, val in enumerate(row):
            if src_ci >= len(col_map):
                break
            tgt_ci = col_map[src_ci]
            if tgt_ci < 0 or val is None or val == "":
                continue
            editor.write_cell("Jobtrack", target_row, tgt_ci, val)
        written += 1
    return written


def _fill_rmc_summary(editor: XLSXEditor, rmc_rows: List[Dict],
                      rmc_col_order: List[str], text_cols: set,
                      month_str: str) -> int:
    """Write our computed values into the RMC summary sheet.

    For columns we compute (Order info, all input/output Kgs and Values,
    Total Cost, Prod RMC/Kg) we OVERWRITE the template's row-7-style
    SUMIF formula with our exact value (engine is 100% accurate vs the
    manual file).

    For columns we don't compute (Diff, Input/Output check, wastage by
    process, RMC per process, Overall Consumption) we leave the
    template's pre-filled formula intact — Excel recalculates it on
    open from the data we wrote into the process sheets.
    """
    if not editor.has_sheet("RMC summary"):
        return 0

    # Refresh the title row 1 to show the detected report month
    if month_str:
        title_cells = editor.read_header_row("RMC summary", 1)
        for ci_0, val in enumerate(title_cells):
            if isinstance(val, str) and "RMC -" in val:
                editor.write_string_inline("RMC summary", 1, ci_0 + 1,
                                           f"RMC - {month_str.upper()}")
                break

    template_headers = editor.read_header_row("RMC summary", RMC_HEADER_ROW)
    if not template_headers:
        return 0

    col_map: Dict[str, int] = {}
    for cn in rmc_col_order:
        nh = _norm_header(cn)
        col_map[cn] = -1
        for ci_0, h in enumerate(template_headers):
            if _norm_header(h) == nh:
                col_map[cn] = ci_0 + 1
                break

    written = 0
    for ri, row in enumerate(rmc_rows):
        target_row = RMC_DATA_ROW + ri
        for cn in rmc_col_order:
            tgt_ci = col_map.get(cn, -1)
            if tgt_ci < 0:
                continue
            val = row.get(cn)
            if val is None or val == "":
                continue
            if cn in text_cols:
                editor.write_cell("RMC summary", target_row, tgt_ci, str(val))
            else:
                try:
                    editor.write_cell("RMC summary", target_row, tgt_ci, float(val))
                except (ValueError, TypeError):
                    editor.write_cell("RMC summary", target_row, tgt_ci, str(val))
        written += 1
    return written


# ── Public entry point ──────────────────────────────────────────────────────

def fill_base_rmc_template(
    template_bytes: bytes,
    enriched_jt_bytes: bytes,
    idx: Dict[str, Any],
    rmc_rows: List[Dict],
    rmc_col_order: List[str],
    text_cols: Iterable[str],
    progress_cb: Optional[Any] = None,
) -> bytes:
    """Fill the user's uploaded Base RMC template in-place and return bytes.

    The template's visual aspects (cell colours, borders, conditional
    formats, charts, pivot tables, defined names) are preserved
    unchanged. Only data cells are modified.

    Args:
      template_bytes: User's uploaded Base RMC workbook (.xlsx bytes).
      enriched_jt_bytes: Output of `engine.fill_jobtrack.fill_jobtrack(...)`.
      idx: Process-sheet OrderIndex dict from `build_all_from_jobtrack(...)`.
      rmc_rows: RMC summary rows from `compute_rmc_summary(...)`.
      rmc_col_order: Column order for the RMC summary (RMC_COL_ORDER).
      text_cols: Set of column names that should be written as strings.
      progress_cb: Optional callable(pct: int, msg: str).

    Returns:
      Bytes of the filled .xlsx workbook.
    """
    text_cols_set = set(text_cols)

    def _p(pct: int, msg: str) -> None:
        if progress_cb:
            progress_cb(pct, msg)
        else:
            logger.info(f"[template_filler:{pct}%] {msg}")

    _p(0, "Opening template archive...")
    editor = XLSXEditor(template_bytes)

    _p(8, "Reading enriched Jobtrack rows...")
    jt_headers, jt_rows = _read_enriched_jobtrack(enriched_jt_bytes)

    month_str = _detect_report_month_from_jt(jt_headers, jt_rows)
    _p(12, f"Detected report month: {month_str or 'unknown'}")

    # 1. Process sheets
    process_sheets_present = [s for s in PROCESS_SHEET_TO_IDX if editor.has_sheet(s)]
    for i, sheet_name in enumerate(process_sheets_present):
        idx_key = PROCESS_SHEET_TO_IDX[sheet_name]
        oi = idx.get(idx_key)
        pct = 15 + int(40 * (i + 1) / max(len(process_sheets_present), 1))
        if oi is None:
            _p(pct, f"{sheet_name}: no engine data, leaving template intact")
            continue
        n = _fill_process_sheet(editor, sheet_name, oi.headers, oi.all_rows)
        _p(pct, f"{sheet_name}: wrote {n} rows")

    # 2. Jobtrack
    if editor.has_sheet("Jobtrack"):
        n = _fill_jobtrack(editor, jt_headers, jt_rows)
        _p(60, f"Jobtrack: wrote {n} enriched rows")

    # 3. RMC summary
    if editor.has_sheet("RMC summary"):
        n = _fill_rmc_summary(editor, rmc_rows, rmc_col_order,
                              text_cols_set, month_str)
        _p(80, f"RMC summary: wrote {n} order rows")

    _p(92, "Saving workbook archive...")
    out_bytes = editor.save()
    _p(100, f"Filled template ready: {len(out_bytes):,} bytes, "
            f"{len(editor.sheet_names)} sheets")
    return out_bytes
