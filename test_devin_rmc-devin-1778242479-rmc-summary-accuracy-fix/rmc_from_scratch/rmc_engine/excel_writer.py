"""
Excel Writer — produces the final RMC output workbook using xlsxwriter.
"""
from __future__ import annotations

import io
import logging
from datetime import datetime
from typing import Any, Dict, List

from rmc_engine.data_reader import OrderIndex, safe_float, safe_str
from rmc_engine.rmc_compute import RMC_COL_ORDER, TEXT_COLS

logger = logging.getLogger(__name__)


def write_rmc_output(
    idx: Dict[str, OrderIndex],
    rmc_rows: List[Dict],
    output_path: str = None,
) -> bytes:
    """
    Write the filled RMC workbook.
    Returns the workbook bytes. If output_path given, also writes to disk.
    """
    import xlsxwriter

    buf = io.BytesIO()
    xwb = xlsxwriter.Workbook(buf, {"in_memory": True})

    hdr_fmt = xwb.add_format({
        "bold": True, "bg_color": "#4472C4", "font_color": "white",
        "border": 1, "text_wrap": True, "valign": "vcenter",
    })
    num_fmt = xwb.add_format({"num_format": "#,##0.00"})
    date_fmt = xwb.add_format({"num_format": "dd-mmm-yyyy"})
    pct_fmt = xwb.add_format({"num_format": "0.00%"})

    # --- RMC Summary sheet ---
    ws = xwb.add_worksheet("RMC summary")
    ws.set_column(0, 0, 15)
    ws.set_column(1, 1, 45)
    ws.set_column(2, 2, 25)
    ws.set_column(3, 6, 15)
    ws.set_column(7, len(RMC_COL_ORDER) - 1, 14)
    ws.freeze_panes(6, 1)

    title_fmt = xwb.add_format({
        "bold": True, "font_size": 14, "font_color": "#1E3A5F",
    })
    subtitle_fmt = xwb.add_format({
        "bold": True, "font_size": 11, "font_color": "#4472C4",
    })
    ws.write(0, 0, "INTEGRATED PLASTICS PACKAGING", title_fmt)
    ws.write(1, 0, "Raw Material Consumption Summary", subtitle_fmt)
    ws.write(2, 0, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    for ci, cn in enumerate(RMC_COL_ORDER):
        ws.write(5, ci, cn, hdr_fmt)

    for ri, row in enumerate(rmc_rows):
        for ci, cn in enumerate(RMC_COL_ORDER):
            val = row.get(cn, "")
            if cn in TEXT_COLS:
                ws.write(6 + ri, ci, safe_str(val))
            else:
                ws.write_number(6 + ri, ci, safe_float(val), num_fmt)

    # --- Process & WIP sheets ---
    for sn, oi in idx.items():
        if sn == "FG":
            continue
        wsx = xwb.add_worksheet(sn[:31])
        sr = 5 if sn not in ("OPN_WIP", "CLS_WIP") else 4
        for ci, h in enumerate(oi.headers):
            wsx.write(sr, ci, h, hdr_fmt)
        for ri, row_tuple in enumerate(oi.all_rows):
            for ci, val in enumerate(row_tuple):
                if val is None:
                    continue
                if isinstance(val, datetime):
                    wsx.write_datetime(sr + 1 + ri, ci, val, date_fmt)
                elif isinstance(val, (int, float)):
                    wsx.write_number(sr + 1 + ri, ci, val, num_fmt)
                else:
                    wsx.write(sr + 1 + ri, ci, str(val))

    # --- FG sheet ---
    fg_oi = idx.get("FG")
    if fg_oi:
        ws_fg = xwb.add_worksheet("FG")
        for ci, h in enumerate(fg_oi.headers):
            ws_fg.write(2, ci, h, hdr_fmt)
        for ri, row_tuple in enumerate(fg_oi.all_rows):
            for ci, val in enumerate(row_tuple):
                if val is None:
                    continue
                if isinstance(val, (int, float)):
                    ws_fg.write_number(3 + ri, ci, val, num_fmt)
                else:
                    ws_fg.write(3 + ri, ci, str(val))

    xwb.close()
    output_bytes = buf.getvalue()

    if output_path:
        with open(output_path, "wb") as f:
            f.write(output_bytes)
        logger.info(f"  Output written: {output_path}")

    return output_bytes
