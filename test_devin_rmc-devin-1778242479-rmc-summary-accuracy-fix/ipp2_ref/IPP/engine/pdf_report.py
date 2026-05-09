"""
PDF Report Generator for IPP Row Explanations.
Uses fpdf2 to produce clean audit-grade PDFs.
"""
import io
import logging
from datetime import datetime

import warnings
warnings.filterwarnings("ignore", message=".*PyFPDF.*fpdf2.*")

from fpdf import FPDF

logger = logging.getLogger(__name__)


class RowReportPDF(FPDF):
    """Custom PDF class with branded header/footer."""

    def __init__(self):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.set_auto_page_break(auto=True, margin=20)

    @staticmethod
    def _safe_text(text) -> str:
        """Sanitize text for Helvetica (latin-1 only)."""
        s = str(text) if text is not None else ''
        # Replace common Unicode chars with ASCII equivalents
        replacements = {
            '\u2014': '-', '\u2013': '-', '\u2018': "'", '\u2019': "'",
            '\u201c': '"', '\u201d': '"', '\u2026': '...', '\u2022': '*',
            '\u2265': '>=', '\u2264': '<=', '\u2248': '~', '\u00b2': '2',
        }
        for uni, asc in replacements.items():
            s = s.replace(uni, asc)
        return s.encode('latin-1', 'replace').decode('latin-1')

    def header(self):
        self.set_font('Helvetica', 'B', 10)
        self.set_text_color(79, 70, 229)  # Indigo
        self.cell(0, 8, 'IPP Jobtrack MRR - Row Audit Report', ln=True, align='L')
        self.set_draw_color(226, 232, 240)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(100, 116, 139)
        self.cell(0, 10, f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}  |  '
                         f'Page {self.page_no()}/{{nb}}', align='C')

    def section_title(self, title: str):
        """Render a section header."""
        self.ln(3)
        self.set_font('Helvetica', 'B', 12)
        self.set_text_color(30, 41, 59)
        self.cell(0, 8, title, ln=True)
        self.set_draw_color(79, 70, 229)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 80, self.get_y())
        self.set_line_width(0.2)
        self.ln(3)

    def key_value(self, key: str, value: str, bold_value: bool = False):
        """Render a key: value pair."""
        self.set_font('Helvetica', 'B', 9)
        self.set_text_color(71, 85, 105)
        key_safe = self._safe_text(key)
        key_w = self.get_string_width(key_safe + ': ') + 2
        self.cell(key_w, 6, key_safe + ': ')
        self.set_font('Helvetica', 'B' if bold_value else '', 9)
        self.set_text_color(30, 41, 59)
        val_str = self._safe_text(value)[:100]
        self.cell(0, 6, val_str, ln=True)

    def table_row(self, cells: list, widths: list, header: bool = False):
        """Render a table row."""
        if header:
            self.set_font('Helvetica', 'B', 8)
            self.set_fill_color(241, 245, 249)
            self.set_text_color(30, 41, 59)
        else:
            self.set_font('Helvetica', '', 8)
            self.set_fill_color(255, 255, 255)
            self.set_text_color(71, 85, 105)

        for i, (cell, w) in enumerate(zip(cells, widths)):
            cell_str = self._safe_text(cell)[:40]
            self.cell(w, 6, cell_str, border=1, fill=header)
        self.ln()

    def confidence_badge(self, level: str, color_hex: str, reason: str):
        """Render a confidence badge."""
        r, g, b = int(color_hex[1:3], 16), int(color_hex[3:5], 16), int(color_hex[5:7], 16)
        self.set_font('Helvetica', 'B', 14)
        self.set_text_color(r, g, b)
        self.cell(30, 10, level)
        self.set_font('Helvetica', '', 9)
        self.set_text_color(71, 85, 105)
        reason_safe = self._safe_text(reason)
        self.cell(0, 10, f'  -  {reason_safe}', ln=True)


def generate_row_pdf(explanation: dict) -> bytes:
    """
    Generate a PDF report for a single row explanation.

    Args:
        explanation: dict from explainer.explain_row()

    Returns:
        PDF bytes
    """
    pdf = RowReportPDF()
    pdf.alias_nb_pages()
    pdf.add_page()

    # ── Section 1: Row Summary ──
    pdf.section_title('Row Summary')
    pdf.key_value('Row Number', str(explanation['row_num']))
    pdf.key_value('UID', str(explanation['uid']))
    pdf.key_value('Process', str(explanation['process']))
    pdf.key_value('Order No', str(explanation['order_no']))

    raw = explanation['raw_row']
    process = str(explanation.get('process', '')).upper().strip()

    if process == 'PRINTING':
        pdf.key_value('Material', str(raw.get('Input Name', '')))
        pdf.key_value('Size (mm)', str(raw.get('Input Size', '')))
        pdf.key_value('Micron', str(raw.get('Input Mic', '')))
        pdf.key_value('Total Qty', str(raw.get('Total Input', '')), bold_value=True)
        pdf.key_value('Film MR#', str(raw.get('Film MR#', '')), bold_value=True)
        pdf.key_value('Film Rate', str(raw.get('Film Rate', '')), bold_value=True)
        pdf.key_value('Film Value', str(raw.get('Film Value', '')), bold_value=True)
    elif process == 'LAM':
        # Fresh 1
        if raw.get('Fresh1 Name'):
            pdf.key_value('1st Fresh Material', str(raw.get('Fresh1 Name', '')))
            pdf.key_value('Fresh1 MR#', str(raw.get('Fresh1 MR#', '')), bold_value=True)
            pdf.key_value('Fresh1 Rate', str(raw.get('Fresh1 Rate', '')), bold_value=True)
            pdf.key_value('Fresh1 Value', str(raw.get('Fresh1 Value', '')), bold_value=True)
        # Fresh 2
        if raw.get('Fresh2 Name'):
            pdf.key_value('2nd Fresh Material', str(raw.get('Fresh2 Name', '')))
            pdf.key_value('Fresh2 MR#', str(raw.get('Fresh2 MR#', '')), bold_value=True)
            pdf.key_value('Fresh2 Rate', str(raw.get('Fresh2 Rate', '')), bold_value=True)
            pdf.key_value('Fresh2 Value', str(raw.get('Fresh2 Value', '')), bold_value=True)
        # Chemicals
        if raw.get('Adh Name'):
            pdf.key_value('Adhesive', str(raw.get('Adh Name', '')))
            pdf.key_value('Adh Rate', str(raw.get('Adh Rate', '')), bold_value=True)
            pdf.key_value('Hard Rate', str(raw.get('Hard Rate', '')), bold_value=True)
            pdf.key_value('Sol Rate', str(raw.get('Sol Rate', '')), bold_value=True)

    # ── Section 2: Raw Excel Data ──
    pdf.section_title('Excel Source Data')
    # Show as key-value pairs (only non-empty)
    for label, val in raw.items():
        if val is not None:
            pdf.key_value(label, str(val))

    # ── Section 3: Rate Breakdown ──
    pdf.section_title('Rate Breakdown')
    breakdown = explanation.get('rate_breakdown', [])
    if breakdown:
        widths = [25, 20, 145]
        pdf.table_row(['Type', 'Status', 'Detail'], widths, header=True)
        for entry in breakdown:
            pdf.table_row(
                [entry.get('type', ''), entry.get('status', ''), entry.get('detail', '')],
                widths
            )
    else:
        pdf.set_font('Helvetica', 'I', 9)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(0, 6, 'No rate calculation entries for this row.', ln=True)

    # ── Section 4: Mismatch / Risk Report ──
    pdf.section_title('Mismatch / Risk Report')
    issues = explanation.get('issues', [])
    if issues:
        for i, issue in enumerate(issues, 1):
            pdf.set_font('Helvetica', 'B', 9)
            sev = issue.get('severity', '')
            if sev == 'LOW':
                pdf.set_text_color(220, 38, 38)
            elif sev == 'MEDIUM':
                pdf.set_text_color(217, 119, 6)
            else:
                pdf.set_text_color(5, 150, 105)
            pdf.cell(8, 6, f'{i}.')
            pdf.cell(50, 6, pdf._safe_text(issue['title']))
            pdf.set_font('Helvetica', '', 8)
            pdf.set_text_color(71, 85, 105)
            pdf.cell(0, 6, f'[{sev}]', ln=True)
            pdf.set_x(18)
            pdf.multi_cell(172, 5, pdf._safe_text(issue.get('detail', '')))
            pdf.ln(1)
    else:
        pdf.set_font('Helvetica', '', 9)
        pdf.set_text_color(5, 150, 105)
        pdf.cell(0, 6, 'No issues detected. All values are direct matches.', ln=True)

    # ── Section 5: Confidence Score ──
    pdf.section_title('Confidence Score')
    conf = explanation.get('confidence', {})
    pdf.confidence_badge(
        conf.get('level', 'N/A'),
        conf.get('color', '#64748B'),
        conf.get('reason', '')
    )

    # Output
    output = io.BytesIO()
    pdf_bytes = pdf.output()
    output.write(pdf_bytes)
    output.seek(0)
    return output.read()
