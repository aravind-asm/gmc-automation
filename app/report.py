"""Excel report generation for the download and the Report Drop (see docs/adr/ADR-006.md).

This module is a **pure formatter**: it is handed data and returns bytes. It never touches
the database, so the sheets can never disagree with the endpoints that serve the same data —
``insights.py`` builds each dataset with the same helper its own routes use and passes it in.

``openpyxl`` is confined to this module; nothing else imports it.
"""
import logging
import math
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname

from app.db import KST  # the constant only: this module still never queries the database

logger = logging.getLogger("gmc.report")

# Rates are written as floats 0..1 carrying this format, never as pre-formatted strings,
# so Excel can sort, chart and compute with them (ADR-006 §5).
RATE_FORMAT = "0.0%"

# The Overview and Spike Radar sheets show rates to 1 decimal place, as the pages do; status and
# spikes are decided on that same rounded figure (insights.RATE_DP; operator, 2026-11-16, was 2).
OVERVIEW_RATE_FORMAT = "0.0%"

SHEET_OVERVIEW = "Overview"
SHEET_TOP_ISSUES = "Top 10 Issues"
SHEET_PRODUCTS = "Disapproved Products"
SHEET_HISTORY = "Approval Rate Over Time"

# Overview status: a label plus Excel's standard good / neutral / bad fills. The label is
# always written too, so the status survives printing in black and white (ADR-009 §2).
STATUS_LABELS = {
    "critical": "Critical",
    "warning": "Warning",
    "minor": "Minor",
    "healthy": "Healthy",
    "not_synced": "Not in last sync",
}
STATUS_STYLES = {
    "critical": (PatternFill("solid", fgColor="FFC7CE"), Font(bold=True, color="9C0006")),
    "warning": (PatternFill("solid", fgColor="FFCC99"), Font(bold=True, color="9C3B00")),  # orange (ADR-011 §3)
    "minor": (PatternFill("solid", fgColor="FFEB9C"), Font(bold=True, color="9C5700")),    # Excel's yellow
    "healthy": (PatternFill("solid", fgColor="C6EFCE"), Font(bold=True, color="006100")),
    "not_synced": (PatternFill(fill_type=None), Font(italic=True, color="7F7F7F")),
}
# Neutral grey for things that are neither good nor bad: Pending review, No earlier figure.
GREY_STYLE = (PatternFill("solid", fgColor="EDEDED"), Font(color="595959"))

# Sheet 1 columns (ADR-006 S4). Deliberately an 11-column superset of the UI's nine:
# `sub_account_id` is the join key across sheets, `product_url` is hidden behind the title
# link in the UI, and `description`/`detail` stay apart (the UI merges them into "Reason").
# `severity` is omitted — DISAPPROVED is its only value, by construction in the extractor.
# Order is the operator's reading order (ADR-006 §4 amendment); the last three are hidden.
PRODUCT_COLUMNS: list[tuple[str, str, int]] = [
    ("sub_account_id", "Account ID", 18),  # header plus its filter arrow
    ("sub_account_name", "Account", 30),
    ("attribute", "Attribute", 20),
    ("description", "Description", 34),
    ("offer_id", "Offer ID", 22),
    ("product_url", "Product URL", 44),
    ("detail", "Detail", 62),
    ("documentation_uri", "Help", 40),
    ("product_title", "Product", 46),
    ("issue_code", "Issue code", 34),
    ("reporting_context", "Surfaces", 40),
]

# Hidden, not removed: the report is the only durable record of a run's product detail
# (ADR-006 Context), so these stay in every row for anyone who unhides them.
HIDDEN_PRODUCT_FIELDS = {"product_title", "issue_code", "reporting_context"}

_TITLE_FONT = Font(bold=True, size=14)
_META_FONT = Font(italic=True, size=9, color="666666")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")


def _fmt_kst_utc(iso: str | None) -> str:
    """A sync time in Korea time with UTC beside it: ``2026-11-12 15:14 KST · 06:14 UTC``.

    The UTC date is added only when it differs from the KST one (ADR-011 §1).
    """
    if not iso:
        return "no sync yet"
    try:
        utc = datetime.fromisoformat(iso).astimezone(timezone.utc)
    except ValueError:
        return iso
    kst = utc.astimezone(KST)
    utc_part = f"{utc:%H:%M}" if utc.date() == kst.date() else f"{utc:%Y-%m-%d %H:%M}"
    return f"{kst:%Y-%m-%d %H:%M} KST · {utc_part} UTC"


def _fmt_kst(iso: str) -> str:
    """A sync time in Korea time only — the rate-over-time column headings."""
    try:
        return f"{datetime.fromisoformat(iso).astimezone(KST):%Y-%m-%d %H:%M} KST"
    except ValueError:
        return iso


def _generated_line(meta: dict) -> str:
    """The one time line every sheet carries: when the data was fetched, not when the file was
    built (ADR-011 §1)."""
    return f"Generated: {_fmt_kst_utc(meta.get('last_synced_at'))}"


def _write_provenance(ws, title: str, meta: dict, scope_lines: list[str]) -> int:
    """Write the per-sheet header block; return the 1-based row the table should start on.

    Every sheet carries its own provenance because the sheets have deliberately different
    scopes — sheets 1-2 describe one sync, sheet 3 spans all of them (ADR-006 §6). A reader
    looking at one sheet in isolation must still be able to tell which.
    """
    ws.cell(row=1, column=1, value=title).font = _TITLE_FONT
    ws.cell(row=2, column=1, value=_generated_line(meta)).font = _META_FONT
    row = 3
    for line in scope_lines:
        ws.cell(row=row, column=1, value=line).font = _META_FONT
        row += 1
    return row + 1  # one blank spacer row between the header block and the table


def _write_table(ws, start_row: int, headers: list[str], rows: list[list], widths: list[int],
                 freeze_cols: int = 0) -> None:
    """Write a header row + data rows, then freeze and autofilter below the provenance block.

    ``freeze_cols`` keeps that many left columns (the store) in view while scrolling right
    (ADR-011 §8).
    """
    for col, (name, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=start_row, column=col, value=name)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = width

    for r, values in enumerate(rows, start=start_row + 1):
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)

    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A{start_row}:{last_col}{start_row + len(rows)}"
    # Freeze the provenance block and the header row so the table scrolls under them, and the
    # first ``freeze_cols`` columns so the store stays beside its figures.
    ws.freeze_panes = f"{get_column_letter(freeze_cols + 1)}{start_row + 1}"


# --- Overview header: a dashboard-style banner of cards, mirroring the Overview page ---------
# Each card spans two of the table's first eight columns, so the cards line up with the table
# below; the ninth (Account warnings) runs on past them. Wide enough for each bold header plus
# its filter arrow ("Sub-account ID", "Total products").
_OVERVIEW_WIDTHS = [14, 18, 20, 34, 19, 17, 19, 20, 60]  # "Approval rate" + filter arrow

# Account warnings column (ADR-009 §3): the text says what it is; the font colour matches the
# Status fills so a red warning reads as the reason for a red status.
_WARNING_LABELS = {"CRITICAL": "Critical", "ERROR": "Error"}
_WARNING_FONTS = {
    "CRITICAL": Font(bold=True, color="9C0006"),
    "ERROR": Font(color="9C5700"),
    "failed": Font(italic=True, color="7F7F7F"),
    "not_checked": Font(italic=True, color="7F7F7F"),
}
_CARD_SPANS = [(1, 2), (3, 4), (5, 6), (7, 8)]
_BRAND = "1F4E79"  # the table header colour, reused so the sheet reads as one design
_CARD_FILL = PatternFill("solid", fgColor="EEF3F8")
_CARD_EDGE = Side(style="thin", color="C9D3DE")
_CARD_BORDER = Border(left=_CARD_EDGE, right=_CARD_EDGE, top=_CARD_EDGE, bottom=_CARD_EDGE)
_OVERVIEW_TITLE_FONT = Font(bold=True, size=20, color=_BRAND)
_SUBTITLE_FONT = Font(size=11, color="595959")
_CARD_LABEL_FONT = Font(bold=True, size=9, color=_BRAND)
_CARD_VALUE_FONT = Font(bold=True, size=20, color="1A1A1A")
_NOTE_FONT = Font(size=9, color="595959")
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _card(ws, row: int, span: tuple[int, int], label: str, value, number_format: str | None,
          fill: PatternFill, label_font: Font, value_font: Font) -> None:
    """One card: a small label row above a big value row, two merged cells with a border.

    The value is written as a real number (with a display format), never as text, so it
    stays usable in Excel. Merged cells are fine here — the "no merged cells" rule
    (ADR-006 §4) protects the filterable product table, not this banner.
    """
    first, last = span
    for r, text, font in ((row, label.upper(), label_font), (row + 1, value, value_font)):
        if last > first:
            ws.merge_cells(start_row=r, start_column=first, end_row=r, end_column=last)
        cell = ws.cell(row=r, column=first, value=text)
        cell.font = font
        cell.alignment = _CENTER
        if r == row + 1 and number_format:
            cell.number_format = number_format
        for c in range(first, last + 1):  # fill and border every cell under a merge
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = _CARD_BORDER


# --- The Overview letterhead (operator, 2026-11-12; ADR-011 amendment) ----------------------
# Overview sheet only: the workbook opens on it, so it is the report's cover. Both logo files are
# committed with the code; a missing file is logged and the sheet is built without it.
REPORT_TITLE = "GMC Feed Global Governance report"
_ASSETS = Path(__file__).parent / "assets"
_STAGWELL_LOGO = _ASSETS / "stagwell_media_logo.png"   # black background removed from the supplied file
_LG_LOGO = _ASSETS / "lg_logo.png"
_LOGO_ROW_HEIGHT = 36                                  # points; 48 px
_STAGWELL_SIZE = (300, 36)                             # px, the file's 1000 × 120 aspect
_LG_SIZE = (40, 40)


def _column_px(width: float) -> int:
    """Excel's on-screen width of a column, in pixels, for a width in characters (Calibri 11)."""
    return int(width * 7 + 5)


def _place_logo(ws, path: Path, size: tuple[int, int], col: int, col_off_px: int, row_off_px: int) -> None:
    """Put one image in row 1, offset inside a cell; skip it (logged) if the file is missing."""
    if not path.is_file():
        logger.warning("report logo missing, sheet built without it: %s", path)
        return
    img = Image(str(path))
    img.width, img.height = size
    img.anchor = OneCellAnchor(
        _from=AnchorMarker(col=col, colOff=pixels_to_EMU(col_off_px), row=0, rowOff=pixels_to_EMU(row_off_px)),
        ext=XDRPositiveSize2D(pixels_to_EMU(size[0]), pixels_to_EMU(size[1])),
    )
    ws.add_image(img)


def _write_letterhead(ws) -> None:
    """Row 1: Stagwell Media on the left (A–C), LG centred in column D — the last frozen column,
    so both logos stay in view when the sheet is scrolled right (operator, 2026-11-12)."""
    ws.row_dimensions[1].height = _LOGO_ROW_HEIGHT
    row_px = _LOGO_ROW_HEIGHT * 4 // 3
    _place_logo(ws, _STAGWELL_LOGO, _STAGWELL_SIZE, 0, 4, (row_px - _STAGWELL_SIZE[1]) // 2)
    lg_col = 3  # column D (0-based)
    _place_logo(ws, _LG_LOGO, _LG_SIZE, lg_col, (_column_px(_OVERVIEW_WIDTHS[lg_col]) - _LG_SIZE[0]) // 2,
                (row_px - _LG_SIZE[1]) // 2)


def _write_overview_header(ws, rows: list[dict], kpis: dict, thresholds: dict, meta: dict,
                           radar: dict | None = None) -> int:
    """Title, KPI cards, status cards and the Spike Radar card; return the row the table starts on.

    Same figures as the page's top banner. The status cards reuse the Status column's fills
    so a colour means the same thing everywhere on the sheet. Everything above the table is
    frozen, so it is kept short: the explanatory notes go **under** the table
    (``_overview_notes``) — with them up here a small window showed about five store rows.
    """
    for col, width in enumerate(_OVERVIEW_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    _write_letterhead(ws)
    ws.cell(row=2, column=1, value=REPORT_TITLE).font = _OVERVIEW_TITLE_FONT
    ws.row_dimensions[2].height = 30
    ws.cell(row=3, column=1, value="Overview  ·  Account health by region and country").font = _SUBTITLE_FONT
    ws.cell(
        row=4,
        column=1,
        value=f"{_generated_line(meta)}  ·  Test accounts are not shown.",
    ).font = _META_FONT

    counts = kpis["status_counts"]
    avg = kpis["average_approval_rate"]
    kpi_cards = [
        ("Total products", kpis["total_products"], "#,##0"),
        ("Disapproved products", kpis["disapproved_products"], "#,##0"),
        ("Average approval rate", "—" if avg is None else avg, None if avg is None else OVERVIEW_RATE_FORMAT),
        ("Accounts", kpis["store_count"], "#,##0"),
    ]
    for span, (label, value, fmt) in zip(_CARD_SPANS, kpi_cards):
        _card(ws, 6, span, label, value, fmt, _CARD_FILL, _CARD_LABEL_FONT, _CARD_VALUE_FONT)

    # Healthy first, then worst to mildest (operator, 2026-11-12) — the same order as the page.
    for span, status in zip(_CARD_SPANS, ["healthy", "critical", "warning", "minor"]):
        fill, font = STATUS_STYLES[status]
        value_font = Font(bold=True, size=20, color=font.color)
        _card(ws, 9, span, STATUS_LABELS[status], counts[status], "#,##0", fill, font, value_font)

    for r, height in ((6, 18), (7, 36), (8, 8), (9, 18), (10, 36)):
        ws.row_dimensions[r].height = height

    ws.row_dimensions[11].height = 8
    row = 12
    if radar is not None:
        row = _write_spike_card(ws, row, radar)
    return row + 1  # one blank spacer row before the table


# Notes under the Overview table (operator, 2026-11-10): the table's own size, not the small
# grey note font, with each status word in its banner colour and bold, and the term each note
# explains in bold. A note is a list of pieces: plain text, or ``(text, style)``.
_NOTE_SIZE = 11
_NOTE_TEXT_COLOR = "404040"
_NOTE_STYLES = {
    "critical": "9C0006",
    "warning": "9C3B00",
    "minor": "9C5700",
    "healthy": "006100",
    "grey": "595959",
    "term": "1A1A1A",
}


def _overview_notes(rows: list[dict], kpis: dict, thresholds: dict, radar: dict | None) -> list[list]:
    """The Overview sheet's explanations, written under the store table."""
    counts = kpis["status_counts"]
    # Approval cut-offs: the disapproval thresholds turned round, so the same stores sit in each
    # band (ADR-011 amendment, 2026-11-12). 5 % disapproved = 95.0 % approved, which is Minor.
    minor, warn, crit = (_approval_pct(thresholds[k]) for k in ("minor", "warning", "critical"))
    notes: list[list] = [
        [("Healthy", "healthy"), f" = approval rate above {minor:.1f}%  ·  ",
         ("Critical", "critical"), f" = {crit:.1f}% or lower, or a critical account warning from Google  ·  ",
         ("Warning", "warning"), f" = above {crit:.1f}% up to {warn:.1f}%  ·  ",
         ("Minor", "minor"), f" = above {warn:.1f}% up to {minor:.1f}%."],
        [("Approval rate", "term"), " = 100% minus the account's disapproval rate: the share of its products "
         "with no disapproval (pending review counts as not approved)."],
        [("Average approval rate", "term"), " = 100% minus the mean of each account's disapproval rate. Approved "
         "= not disapproved, so pending review counts as not approved."],
        [("Account warnings", "term"), " = problems Google reports for the whole account, not one product. Error "
         "warnings are listed but do not change the colour. An account red from its rate says so there first "
         "(\"Low approval rate\"). Blank = none."],
    ]
    if radar is not None and radar.get("state") == "ok":
        th = radar["thresholds"]
        notes.append(
            [("Spike Radar", "term"), ": each account's approval rate now against the last sync of an earlier day "
             f"(Korea time). Spike = at least {th['min_disapproved']} disapproved products now and approval down "
             f"by at least {th['min_diff_pp']:.1f} percentage points (pp)."]
        )
    failed = sum(1 for r in rows if r["account_check"] == "failed")
    if failed:
        notes.append([(f"{failed} account(s): warning check failed", "term"),
                      " — Google did not answer, so their status comes from the rate only."])
    if counts["not_synced"]:
        notes.append([(f"{counts['not_synced']} account(s) Not in last sync", "term"),
                      ": listed in the Account Directory but no figures in this sync (newly added, or its fetch "
                      "failed). Their cells are blank, not zero."])
    unassigned = sum(1 for r in rows if not r["in_directory"])
    if unassigned:
        notes.append([(f"{unassigned} account(s) are not in the Account Directory", "term"), " and show as Unassigned."])
    return notes


def _rich_note(pieces: list) -> CellRichText:
    """One note as a single cell with mixed styling (Excel rich text), so it stays one line."""
    plain = InlineFont(rFont="Calibri", sz=_NOTE_SIZE, color=_NOTE_TEXT_COLOR)
    blocks = []
    for piece in pieces:
        if isinstance(piece, tuple):
            text, style = piece
            blocks.append(TextBlock(InlineFont(rFont="Calibri", sz=_NOTE_SIZE, b=True, color=_NOTE_STYLES[style]), text))
        else:
            blocks.append(TextBlock(plain, piece))
    return CellRichText(blocks)


def _approval_pct(disapproval_threshold: float) -> float:
    """A disapproval threshold (fraction) as the approval cut-off people read (percent, 1 dp)."""
    return round(100 - disapproval_threshold * 100, 1)


def _write_notes(ws, row: int, notes: list[list]) -> None:
    """A "Notes" heading and one line per note, starting at ``row``."""
    ws.cell(row=row, column=1, value="Notes").font = Font(bold=True, size=12, color=_BRAND)
    for i, note in enumerate(notes, start=1):
        ws.cell(row=row + i, column=1, value=_rich_note(note))
        ws.row_dimensions[row + i].height = 18


def _spike_card_lines(radar: dict) -> tuple[str, list[str]]:
    """Heading and body lines of the Spike Radar card — the same words as the page's card."""
    if radar["state"] != "ok":
        body = (
            "Only one day of syncs so far — the radar needs a sync on an earlier day (Korea time) to compare with."
            if radar["state"] == "one_day" else "No sync yet, so there is nothing to compare."
        )
        return "SPIKE RADAR", [body]
    t0, t1 = radar["t0"], radar["t1"]
    days = t1["days_back"]
    count = radar["spike_count"]
    result = "NO SPIKES" if count == 0 else f"{count} SPIKE{'' if count == 1 else 'S'}"
    # One heading line carries the dates and the pointer, so the frozen card stays short.
    heading = (f"SPIKE RADAR — {result}  ·  {t0['label']} compared with {t1['label']} "
               f"({days} day{'s' if days != 1 else ''} earlier)  ·  Every account's change: Spike Radar sheet")
    lines = []
    if count == 0:
        drop = radar.get("biggest_drop")
        lines.append(
            "No account passed both tests this time."
            + (f" Biggest approval drop: {drop['sub_account_name']} {drop['approval_diff_pp']:+.1f} pp "
               f"(approval {drop['t1_approval'] * 100:.1f}% → {drop['t0_approval'] * 100:.1f}%, "
               f"{drop['t0_disapproved']} disapproved)." if drop else " No account's approval fell.")
        )
    else:
        for r in radar["spikes"]:
            lines.append(
                f"{r['approval_diff_pp']:+.1f} pp  {r['sub_account_name']} ({r['country']}) — approval "
                f"{r['t1_approval'] * 100:.1f}% → {r['t0_approval'] * 100:.1f}%, {r['t0_disapproved']} disapproved"
            )
        if count > len(radar["spikes"]):
            lines.append(f"…and {count - len(radar['spikes'])} more on the Spike Radar sheet.")
    return heading, lines


def _write_spike_card(ws, row: int, radar: dict) -> int:
    """The Spike Radar card, as on the Overview page: one box across the eight table columns,
    red-tinted when there is a spike. Returns the row after it."""
    heading, lines = _spike_card_lines(radar)
    spikes = radar.get("spike_count", 0) > 0
    fill = STATUS_STYLES["critical"][0] if spikes else _CARD_FILL
    heading_font = Font(bold=True, size=11, color="9C0006" if spikes else _BRAND)
    first, last = 1, len(_OVERVIEW_WIDTHS) - 1  # columns A..H, like the cards above
    for i, (text, font) in enumerate([(heading, heading_font)] + [(t, _NOTE_FONT) for t in lines]):
        r = row + i
        ws.merge_cells(start_row=r, start_column=first, end_row=r, end_column=last)
        cell = ws.cell(row=r, column=first, value=text)
        cell.font = font
        cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        for c in range(first, last + 1):
            ws.cell(row=r, column=c).fill = fill
            ws.cell(row=r, column=c).border = Border(
                left=_CARD_EDGE if c == first else None,
                right=_CARD_EDGE if c == last else None,
                top=_CARD_EDGE if i == 0 else None,
                bottom=_CARD_EDGE if i == len(lines) else None,
            )
        ws.row_dimensions[r].height = 20 if i == 0 else 16
    return row + len(lines) + 1


def _build_overview_sheet(ws, rows: list[dict], kpis: dict, thresholds: dict, meta: dict,
                          radar: dict | None = None) -> None:
    """Overview — every real store in the latest sync, by region and country (ADR-009).

    Rows and KPIs come from the same helpers as ``GET /api/overview``, so the sheet and the
    page cannot disagree. Replaces the Leaderboard sheet (ADR-006 amendment).
    """
    start = _write_overview_header(ws, rows, kpis, thresholds, meta, radar)
    notes = _overview_notes(rows, kpis, thresholds, radar)

    if not rows:
        ws.cell(row=start, column=1, value="No accounts to show for this sync.").font = _NOTE_FONT
        _write_notes(ws, start + 2, notes)
        return

    _write_table(
        ws,
        start,
        ["Region", "Country", "Account ID", "Account", "Total products", "Disapproved", "Approval rate", "Status",
         "Account warnings"],
        [
            [
                r["region"],
                r["country"],
                r["sub_account_id"],
                r["sub_account_name"] or r["sub_account_id"],
                r["total_products"],  # None for not_synced -> an empty cell, never 0
                r["disapproved_products"],
                r["approval_rate"],  # None for not_synced -> an empty cell
                STATUS_LABELS[r["status"]],
                _warnings_text(r, thresholds["critical"]),
            ]
            for r in rows
        ],
        _OVERVIEW_WIDTHS,
        freeze_cols=4,
    )
    for row, r in enumerate(rows, start=start + 1):
        ws.cell(row=row, column=7).number_format = OVERVIEW_RATE_FORMAT
        fill, font = STATUS_STYLES[r["status"]]
        status_cell = ws.cell(row=row, column=8)
        status_cell.fill = fill
        status_cell.font = font
        warning_font = _warnings_font(r)
        if warning_font is not None:
            ws.cell(row=row, column=9).font = warning_font
    _write_notes(ws, start + len(rows) + 2, notes)  # one blank row after the last store


def _warnings_text(row: dict, critical_threshold: float) -> str | None:
    """One line for the Account warnings cell. ``None`` (an empty cell) means checked, none
    found — the banner note says so. One line, not wrapped, so the table stays filterable.

    A store red from its rate says so first, labelled as the rate, so a Critical row never has
    an empty reason and is never mistaken for one Google flagged (operator request, 2026-11-10).
    """
    parts = []
    if row["critical_by_rate"]:
        parts.append(
            f"Low approval rate: {row['approval_rate'] * 100:.1f}% "
            f"(Critical at {_approval_pct(critical_threshold):.1f}% or lower)"
        )
    if row["account_check"] == "failed":
        parts.append("Warning check failed")
    elif row["account_check"] == "not_checked":
        parts.append("Not checked")
    else:
        parts.extend(f"{_WARNING_LABELS[i['severity']]}: {i['title']}" for i in row["account_issues"])
    return "; ".join(parts) or None


def _warnings_font(row: dict) -> Font | None:
    if row["critical_by_rate"]:
        return _WARNING_FONTS["CRITICAL"]
    if row["account_check"] != "ok":
        return _WARNING_FONTS[row["account_check"]]
    if not row["account_issues"]:
        return None
    return _WARNING_FONTS[row["account_issues"][0]["severity"]]  # worst first, so this is the worst


# --- Spike Radar (ADR-009 §5) ------------------------------------------------------------
SHEET_SPIKES = "Spike Radar"
_SPIKE_WIDTHS = [12, 16, 18, 34, 18, 20, 19, 18, 16, 20]  # bold headers plus their filter arrows
# Columns A-D (Region to Account) are frozen (ADR-011 §8), so no card may straddle D and E: Spikes
# spans A-C, Accounts compared sits in D alone (the widest column), the other two follow.
_SPIKE_CARD_SPANS = [(1, 3), (4, 4), (5, 6), (7, 8)]
PP_FORMAT = '+0.0" pp";-0.0" pp";0.0" pp"'


def _build_spike_sheet(ws, radar: dict, meta: dict) -> None:
    """Spike Radar — each store's approval rate now against the last sync of an earlier KST day.

    Rows come from the same helper as ``GET /api/spikes``. Rates are numbers with a % format;
    a store with no earlier figure has blank before/change cells, never zeros. Shown in approval
    terms, biggest gain first (ADR-011 §5).
    """
    th = radar["thresholds"]
    rule = [("Spike", "critical"), f" = at least {th['min_disapproved']} disapproved products now AND approval down "
            f"by at least {th['min_diff_pp']:.1f} percentage points (pp). Both bounds count."]

    if radar["state"] != "ok":
        _write_banner_title(ws, "Spike Radar", "Accounts whose approval rate fell sharply since an earlier day",
                            meta, _SPIKE_WIDTHS)
        message = (
            "Only one day of syncs so far — the radar needs a sync on an earlier day (Korea time) to "
            "compare with. It fills in after the next Refresh on a later day."
            if radar["state"] == "one_day" else "No sync yet, so there is nothing to compare."
        )
        ws.cell(row=5, column=1, value=message).font = Font(size=_NOTE_SIZE, color=_NOTE_TEXT_COLOR)
        _write_notes(ws, 7, [rule])
        return

    rows, t0, t1 = radar["rows"], radar["t0"], radar["t1"]
    days = t1["days_back"]
    # The dates stay in the frozen banner (subtitle); the explanations go under the table.
    _write_banner_title(
        ws, "Spike Radar",
        f"Accounts whose approval rate fell sharply:  {t0['label']} compared with {t1['label']} "
        f"({days} day{'s' if days != 1 else ''} earlier)",
        meta, _SPIKE_WIDTHS,
    )
    compared = [r for r in rows if not r["baseline_missing"]]
    no_baseline = len(rows) - len(compared)
    average = radar.get("average_approval_diff_pp")
    drop = radar.get("biggest_drop")

    spike_style = STATUS_STYLES["critical" if radar["spike_count"] else "healthy"]
    grey_fill, grey_font = GREY_STYLE
    cards = [
        ("Spikes", radar["spike_count"], "#,##0", spike_style[0], spike_style[1],
         Font(bold=True, size=20, color=spike_style[1].color)),
        ("Accounts compared", len(compared), "#,##0", _CARD_FILL, _CARD_LABEL_FONT, _CARD_VALUE_FONT),
        ("Approval change", "—" if average is None else average,
         None if average is None else PP_FORMAT, _CARD_FILL, _CARD_LABEL_FONT, _CARD_VALUE_FONT),
        ("No earlier figure", no_baseline, "#,##0", grey_fill, grey_font,
         Font(bold=True, size=20, color=grey_font.color)),
    ]
    for span, (label, value, fmt, fill, label_font, value_font) in zip(_SPIKE_CARD_SPANS, cards):
        _card(ws, 5, span, label, value, fmt, fill, label_font, value_font)
    for r, height in ((5, 18), (6, 36)):
        ws.row_dimensions[r].height = height

    notes = [rule]
    if not radar["spike_count"]:
        drop_text = (
            f" Biggest approval drop: {drop['sub_account_name']} {drop['approval_diff_pp']:+.1f} pp "
            f"({drop['t0_disapproved']} disapproved)." if drop else " No account's approval fell."
        )
        notes.append([("No spikes this time", "term"), ": no account passed both tests." + drop_text])
    notes += [
        [("Approval before / Approval now", "term"), f" = 100% minus the disapproval rate at the last sync of "
         f"{t1['label']} and of {t0['label']}. Days are Korea time; several syncs on one day count as the last one."],
        [("Approval change", "term"), " (card) = the average change of the accounts compared. Table: biggest gain "
         "first, so the sharpest drops are at the bottom."],
        [("No earlier figure", "grey"), " = the account had no rate on the earlier day (new, or its fetch "
         "failed, or it had no products), so it is not compared. Test accounts "
         "are left out."],
    ]
    start = 8  # straight after the cards: the notes live under the table

    _write_table(
        ws,
        start,
        ["Region", "Country", "Account ID", "Account", "Total products", "Disapproved now", "Approval before",
         "Approval now", "Change (pp)", "Tag"],
        [
            [
                r["region"],
                r["country"],
                r["sub_account_id"],  # text, like the Overview (ADR-011 amendment 2026-11-23)
                r["sub_account_name"] or r["sub_account_id"],
                r["t0_total"],
                r["t0_disapproved"],
                r["t1_approval"],
                r["t0_approval"],
                r["approval_diff_pp"],
                "Spike" if r["spike"] else ("No earlier figure" if r["baseline_missing"] else None),
            ]
            for r in rows
        ],
        _SPIKE_WIDTHS,
        freeze_cols=4,
    )
    for row, r in enumerate(rows, start=start + 1):
        ws.cell(row=row, column=5).number_format = "#,##0"
        ws.cell(row=row, column=7).number_format = OVERVIEW_RATE_FORMAT
        ws.cell(row=row, column=8).number_format = OVERVIEW_RATE_FORMAT
        ws.cell(row=row, column=9).number_format = PP_FORMAT
        tag = ws.cell(row=row, column=10)
        if r["spike"]:
            tag.fill, tag.font = STATUS_STYLES["critical"]
        elif r["baseline_missing"]:
            tag.fill, tag.font = grey_fill, grey_font
    _write_notes(ws, start + len(rows) + 2, notes)  # one blank row after the last store


# --- Top 10 Issues: same banner style as Overview (operator: every sheet like the dashboard) ---
PENDING_LABEL = "Pending review — not a fault"
_TOP_ISSUES_WIDTHS = [10, 40, 18, 40, 13, 13, 13, 30]  # eight columns, so the four cards line up; "Accounts" + arrow fits


def _write_banner_title(ws, title: str, subtitle: str, meta: dict, widths: list[int]) -> None:
    """Rows 1-3 of a dashboard-style sheet: big title, subtitle, provenance line."""
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.cell(row=1, column=1, value=title).font = _OVERVIEW_TITLE_FONT
    ws.row_dimensions[1].height = 30
    ws.cell(row=2, column=1, value=subtitle).font = _SUBTITLE_FONT
    ws.cell(
        row=3,
        column=1,
        value=f"{_generated_line(meta)}  ·  Test accounts are not counted.",
    ).font = _META_FONT


def _build_top_issues_sheet(ws, issues: list[dict], totals: dict, pending_codes: list[str], meta: dict) -> None:
    """Top 10 Issues — the issue codes affecting the most products (ADR-010 §3).

    Rows and totals come from the same helper as ``GET /api/issues/top``, so the sheet and the
    page cannot disagree. Pending-review codes are counted but tagged grey.
    """
    _write_banner_title(ws, "Top 10 Issues", "The problems affecting the most products, all accounts",
                        meta, _TOP_ISSUES_WIDTHS)

    disapproved = totals["disapproved_products"]
    top_share = totals.get("top3_share")
    top_codes = totals.get("top3_codes") or []
    cards = [
        ("Disapproved products", disapproved, "#,##0", _CARD_FILL, _CARD_LABEL_FONT, _CARD_VALUE_FONT),
        ("Issue types", totals["issue_codes"], "#,##0", _CARD_FILL, _CARD_LABEL_FONT, _CARD_VALUE_FONT),
        ("Top 3 issues share", "—" if top_share is None else top_share,
         None if top_share is None else "0%", _CARD_FILL, _CARD_LABEL_FONT, _CARD_VALUE_FONT),
    ]
    grey_fill, grey_font = GREY_STYLE
    cards.append(("Pending review products", totals["pending_products"], "#,##0", grey_fill, grey_font,
                  Font(bold=True, size=20, color=grey_font.color)))
    for span, (label, value, fmt, fill, label_font, value_font) in zip(_CARD_SPANS, cards):
        _card(ws, 5, span, label, value, fmt, fill, label_font, value_font)
    for r, height in ((5, 18), (6, 36)):
        ws.row_dimensions[r].height = height

    notes = [
        [("Products", "term"), " = distinct products with this issue in the latest sync. A product with several "
         "issues counts under each, so the column adds up to more than the total."],
        [("Top 3 issues share", "term"), " = products with any of the three biggest fixable issues"
         + (f" ({', '.join(top_codes)})" if top_codes else "")
         + " ÷ all disapproved products. A product with two of them counts once; pending review is skipped."],
        [("Countries", "term"), " come from the Account Directory; ", ("Accounts", "term"),
         " = Merchant Center accounts. Test accounts are not counted."],
        [(PENDING_LABEL, "grey"), ": Google has not finished its review "
         f"({', '.join(pending_codes) or 'none set'}). Counted like the rest, but there is nothing to fix."],
    ]
    start = 8  # straight after the cards: the notes live under the table

    if not issues:
        empty = ws.cell(row=start, column=1, value="No disapproved products in this sync, so there are no issues to rank.")
        empty.font = Font(size=_NOTE_SIZE, color=_NOTE_TEXT_COLOR)
        _write_notes(ws, start + 2, notes)
        return

    _write_table(
        ws,
        start,
        ["Rank", "Issue", "Attribute", "Issue code", "Products", "Countries", "Accounts", "Tag"],
        [
            [i["rank"], i["description"], i["attribute"], i["issue_code"], i["products"], i["countries"],
             i["stores"], PENDING_LABEL if i["pending"] else None]
            for i in issues
        ],
        _TOP_ISSUES_WIDTHS,
        freeze_cols=2,
    )
    for row, issue in enumerate(issues, start=start + 1):
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
        if issue["pending"]:
            tag = ws.cell(row=row, column=8)
            tag.fill, tag.font = grey_fill, grey_font
    last = start + len(issues)
    notes_row = last + 2  # one blank row after the last issue
    _write_notes(ws, notes_row, notes)

    # A native bar chart of the same ten rows, biggest at the top, so it reads like the page.
    chart = BarChart()
    chart.type = "bar"
    chart.title = "Products affected, by issue"
    chart.legend = None
    chart.height = 9
    chart.width = 26
    chart.add_data(Reference(ws, min_col=5, min_row=start, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=start + 1, max_row=last))
    chart.x_axis.scaling.orientation = "maxMin"  # rank 1 at the top
    chart.y_axis.crosses = "max"  # ...which would push the number scale to the top; keep it at the bottom
    chart.y_axis.majorGridlines = None
    # Just the count at the end of each bar. Excel shows series name, category and legend key
    # too unless each is switched off explicitly.
    chart.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                                     showLegendKey=False, showPercent=False)
    chart.title.overlay = False  # keep the title above the bars, not on top of rank 1
    series = chart.series[0]
    series.graphicalProperties.solidFill = _BRAND
    for idx, issue in enumerate(issues):
        if issue["pending"]:  # grey, like its Tag cell: nothing to fix
            point = DataPoint(idx=idx)
            point.graphicalProperties.solidFill = "BFBFBF"
            series.dPt.append(point)
    for axis, position in ((chart.x_axis, "l"), (chart.y_axis, "b")):
        axis.delete = False  # openpyxl leaves axes unlabelled otherwise (see the history chart)
        axis.axPos = position
        axis.tickLblPos = "nextTo"
    ws.add_chart(chart, f"A{notes_row + len(notes) + 2}")  # under the notes


# --- Priority List (ADR-010) ------------------------------------------------------------
SHEET_PRIORITY = "Priority List"
_PRIORITY_WIDTHS = [30, 12, 16, 32, 38, 16, 36, 14, 16, 20, 26, 12, 14, 21]
PRIORITY_LABELS = {"critical": "Critical", "high": "High", "medium": "Medium", "low": "Low", "pending": PENDING_LABEL}
# The store-status colours, so a colour means one thing in the workbook: red, orange, yellow, green, grey.
_PRIORITY_STYLES = {"critical": "critical", "high": "warning", "medium": "minor", "low": "healthy", "pending": "grey"}
_PRIORITY_FILLS = {tag: GREY_STYLE if style == "grey" else STATUS_STYLES[style] for tag, style in _PRIORITY_STYLES.items()}
# Six cards (Problems + four levels + Pending review), two columns each, over the first twelve columns.
_PRIORITY_CARD_SPANS = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
_PERF_FIELDS = ["impressions", "ads_impressions", "free_impressions", "clicks", "ads_clicks", "free_clicks"]
_PERF_COL = 9  # Impressions; the five other figures follow it
NO_PERFORMANCE = "No performance data"


def _pct(share: float) -> str:
    return f"{share * 100:g} %"


def _priority_notes(priority: dict) -> list[list]:
    """The Priority List sheet's explanations, written under the table — the page's words."""
    shares = priority.get("shares") or {"critical": 0.10, "high": 0.30}
    perf = priority.get("performance")
    crit, high = _pct(shares["critical"]), _pct(shares["high"])
    window = f" ({perf['label']})" if perf and perf.get("label") else ""
    notes = [
        [("Critical", "critical"), f" = in the top {crit} of problems by impressions, or the account has a critical "
         "account warning from Google  ·  ",
         ("High", "warning"), f" = top {crit} to {high}  ·  ",
         ("Medium", "minor"), " = the rest with impressions, and accounts without performance data  ·  ",
         ("Low", "healthy"), " = 0 impressions  ·  ",
         (PENDING_LABEL, "grey"), ": Google is still reviewing, so there is nothing to fix."],
        [("Ranking", "term"), f": {priority.get('ranked_groups', 0)} problem(s) ranked. Pending review and accounts "
         "without performance data are not ranked. Problems with equal impressions share a place, and the shares "
         "are rounded up, so with any impressions at all at least one problem is Critical."],
        [("Order", "term"), ": Critical, High, Medium, Low, then Pending review; within each, most impressions "
         "first, then most affected products."],
        [("Impressions", "term"), " / ", ("Clicks", "term"),
         f" = how often these products were shown / clicked on Google over the last 30 complete days{window}, "
         "from Merchant Center's performance report. Ads = Shopping ads; Free listings = unpaid listings; the "
         "total also counts affiliate reach. A disapproval often blocks only some places, so a disapproved "
         "product can still be shown."],
    ]
    if perf is None:
        notes.append([(NO_PERFORMANCE, "term"), ": this sync was made before performance data was fetched. The "
                      "next Refresh fetches it."])
    elif perf["stores_without_data"]:
        notes.append([(NO_PERFORMANCE, "term"), f": Google's performance report could not be read for "
                      f"{perf['stores_without_data']} account(s) in this sync. Their problems are not ranked (they "
                      "show as Medium) and their figures are blank, not zero."])
    # No "Open (days)" here: the operator dropped it from the sheet (2026-11-11); the page keeps it.
    notes.append([("Products", "term"), " = distinct products in this account with this issue. Test accounts are "
                  "left out."])
    return notes


def _build_priority_sheet(ws, priority: dict, meta: dict) -> None:
    """Priority List — problems grouped by store × issue, what to fix first (ADR-010 §1-3).

    Rows come from the same helper as ``GET /api/priority``. Layout as every dashboard sheet:
    banner and cards above the table, notes under it (ADR-006 amendment). Figures are numbers;
    a store without performance data says so in its Impressions cell and leaves the rest blank.
    Days open is on the page only — the operator left it off the sheet (2026-11-11).
    """
    _write_banner_title(ws, "Priority List", "What to fix first: each account's problems, grouped by issue",
                        meta, _PRIORITY_WIDTHS)
    groups, counts = priority["groups"], priority["tag_counts"]

    cards = [("Problems", len(groups), _CARD_FILL, _CARD_LABEL_FONT, _CARD_VALUE_FONT)]
    for tag in ("critical", "high", "medium", "low", "pending"):
        fill, font = _PRIORITY_FILLS[tag]
        label = "Pending review" if tag == "pending" else PRIORITY_LABELS[tag]
        cards.append((label, counts[tag], fill, font, Font(bold=True, size=20, color=font.color)))
    for span, (label, value, fill, label_font, value_font) in zip(_PRIORITY_CARD_SPANS, cards):
        _card(ws, 5, span, label, value, "#,##0", fill, label_font, value_font)
    for r, height in ((5, 18), (6, 36)):
        ws.row_dimensions[r].height = height

    notes = _priority_notes(priority)
    start = 8  # straight after the cards: the notes live under the table

    if not groups:
        empty = ws.cell(row=start, column=1, value="No disapproved products in this sync, so there is nothing to prioritise.")
        empty.font = Font(size=_NOTE_SIZE, color=_NOTE_TEXT_COLOR)
        _write_notes(ws, start + 2, notes)
        return

    _write_table(
        ws,
        start,
        ["Tag", "Region", "Country", "Account", "Issue", "Attribute", "Issue code", "Products",
         "Impressions", "Ads impressions", "Free listing impressions", "Clicks", "Ads clicks", "Free listing clicks"],
        [
            [PRIORITY_LABELS[g["tag"]], g["region"], g["country"], g["sub_account_name"], g["description"],
             g["attribute"], g["issue_code"], g["products"]]
            + ([g[f] for f in _PERF_FIELDS] if g.get("has_performance") else [NO_PERFORMANCE] + [None] * 5)
            for g in groups
        ],
        _PRIORITY_WIDTHS,
        freeze_cols=4,
    )
    for row, g in enumerate(groups, start=start + 1):
        tag = ws.cell(row=row, column=1)
        tag.fill, tag.font = _PRIORITY_FILLS[g["tag"]]
        if g.get("has_performance"):
            for col in range(_PERF_COL, _PERF_COL + len(_PERF_FIELDS)):
                ws.cell(row=row, column=col).number_format = "#,##0"
        else:
            ws.cell(row=row, column=_PERF_COL).font = Font(italic=True, color=_NOTE_STYLES["grey"])
    _write_notes(ws, start + len(groups) + 2, notes)  # one blank row after the last group


def _build_products_sheet(ws, rows: list[dict], meta: dict) -> None:
    """Sheet 1 — every disapproval row from the last sync, flat (ADR-006 S4)."""
    products = len({(r["sub_account_id"], r["offer_id"]) for r in rows})
    accounts = len({r["sub_account_id"] for r in rows})
    scope = [
        "Scope: every disapproved product from the last sync, across all accounts.",
        # Row count is not product count — one product repeats once per disapproving reason.
        # Both are stated so the difference cannot be misread (ADR-006 Consequences).
        f"{len(rows)} row(s) · {products} product(s) · {accounts} account(s). "
        "One row per product + reason, so a product with several reasons repeats.",
    ]
    start = _write_provenance(ws, "Disapproved Products", meta, scope)

    if not rows:
        ws.cell(row=start, column=1, value="No disapproved products in this sync.")
        ws.column_dimensions["A"].width = 60
        return

    _write_table(
        ws,
        start,
        [label for _, label, _ in PRODUCT_COLUMNS],
        [[r.get(field) or "" for field, _, _ in PRODUCT_COLUMNS] for r in rows],
        [width for _, _, width in PRODUCT_COLUMNS],
        freeze_cols=2,
    )
    for col, (field, _, _) in enumerate(PRODUCT_COLUMNS, start=1):
        if field in HIDDEN_PRODUCT_FIELDS:
            ws.column_dimensions[get_column_letter(col)].hidden = True


_HISTORY_LEFT = ["Region", "Country", "Account", "Latest approval rate"]
_HISTORY_LEFT_WIDTHS = [12, 16, 32, 22]


def _approval_axis_min(values: list[float]) -> float | None:
    """The chart's floor: the lowest approval shown, rounded down to the nearest 5 % (ADR-011
    amendment, 2026-11-12). Excel's automatic scale drops to 0 % once a value is below about 5/6 of
    the top, squashing every line into the top strip; a fixed floor would cut low stores off."""
    if not values:
        return None
    return math.floor(round(min(values) * 100, 6) / 5) * 5 / 100


def _build_history_sheet(ws, history: list[dict], meta: dict) -> None:
    """Approval rate over time for every store, plus a native line chart (ADR-006 §2, §5; in
    approval terms since the ADR-011 amendment of 2026-11-12).

    One row per store (Region, Country, Store, Latest approval rate, then one column per sync),
    lowest latest approval first (a store whose latest sync had no products — a gap — last). Each row is one series of the chart, and the chart plots visible cells
    only, so the table's filter arrows choose which stores' lines are drawn (operator,
    2026-11-11): filter a store out and its line leaves the chart. Nothing is filtered by default.
    """
    stores_: dict[str, dict] = {}
    for r in history:  # oldest first, so the last write per store is its latest
        entry = stores_.setdefault(r["sub_account_id"], {"rates": {}})
        entry.update(name=r["sub_account_name"] or r["sub_account_id"], region=r.get("region", ""),
                     country=r.get("country", ""), latest=r["approval_rate"])
        entry["rates"][r["synced_at"]] = r["approval_rate"]
    syncs = sorted({r["synced_at"] for r in history})
    rows = sorted(stores_.values(),
                  key=lambda e: (e["latest"] is None, e["latest"] or 0, e["name"].lower()))

    scope = [
        "Scope: every account (test accounts left out), across all recorded history. One row and one line per "
        "account, lowest latest approval rate first. Approval rate = 100% minus the disapproval rate.",
        "To choose the accounts on the chart, use the filter arrows on Region, Country or Account: an account "
        "filtered out of the table leaves the chart too. Clear the filter to show them all again.",
        f"{len(syncs)} sync(s) recorded."
        + (" A line needs two syncs — with one, each account shows a single point." if len(syncs) < 2 else ""),
        "A blank cell (a gap in a line) means no rate was recorded for that account at that sync — its fetch "
        "failed, or it had no products. It is not a zero.",
    ]
    start = _write_provenance(ws, "Approval Rate Over Time", meta, scope)

    if not rows or not syncs:
        ws.cell(row=start, column=1, value="No approval rates recorded yet, so there is nothing to plot.")
        ws.column_dimensions["A"].width = 70
        return

    _write_table(
        ws,
        start,
        _HISTORY_LEFT + [_fmt_kst(sync) for sync in syncs],
        [
            # A store absent from a sync gets None — a genuinely empty cell, never 0. ADR-002
            # omits it rather than record a rate from a partial fetch, and 0 is a real rate that
            # would draw a confident line to the floor for a moment we never measured.
            [e["region"], e["country"], e["name"], e["latest"]] + [e["rates"].get(sync) for sync in syncs]
            for e in rows
        ],
        _HISTORY_LEFT_WIDTHS + [22] * len(syncs),  # "2026-11-11 16:01 KST" plus the filter arrow
    )
    first_col = len(_HISTORY_LEFT) + 1
    last_col = first_col + len(syncs) - 1
    last_row = start + len(rows)
    for row in range(start + 1, last_row + 1):
        for col in range(len(_HISTORY_LEFT), last_col + 1):
            ws.cell(row=row, column=col).number_format = RATE_FORMAT
    # Keep the store names in view while scrolling across the syncs.
    ws.freeze_panes = f"{get_column_letter(first_col)}{start + 1}"

    chart = LineChart()
    chart.title = "Approval rate over time"
    chart.y_axis.title = "Approval rate"
    chart.y_axis.numFmt = RATE_FORMAT
    chart.x_axis.title = "Sync"
    floor = _approval_axis_min([v for e in rows for v in e["rates"].values() if v is not None])
    if floor is not None:
        chart.y_axis.scaling.min = floor
        chart.y_axis.scaling.max = 1.0
        # Whole-percent gridlines over a narrow range (Excel picked odd 2 % steps that left 100 %
        # unlabelled); 5 % steps once the range is wide, so the labels never crowd.
        chart.y_axis.majorUnit = 0.01 if 1.0 - floor <= 0.20 else 0.05
    # Tall enough for one legend line per store (about 0.55 cm each), never shorter than before.
    chart.height = max(15, 0.55 * len(rows) + 4)
    chart.width = 32
    # The legend keeps its own strip on the right and never draws over the lines (operator,
    # 2026-11-12: it overlapped the plot). Fractions of the chart: plot on the left ~76 %,
    # legend in the right ~21 %.
    chart.legend.position = "r"
    chart.legend.overlay = False
    # openpyxl writes `chart.layout` as the plot area's layout (it overwrites plot_area.layout).
    chart.layout = Layout(manualLayout=ManualLayout(
        layoutTarget="inner", xMode="edge", yMode="edge", x=0.07, y=0.08, w=0.68, h=0.80))
    chart.legend.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.78, y=0.06, w=0.21, h=0.90))
    # Empty cells render as gaps, never spanned or zeroed. This is the ADR-002 invariant:
    # "a visible gap is strictly better than an invisible lie." openpyxl already defaults to
    # "gap"; set explicitly so the requirement is stated here rather than inherited silently.
    chart.display_blanks = "gap"
    # The filter is the store picker: Excel leaves hidden (filtered-out) rows off the chart only
    # when it plots visible cells only. openpyxl's default, stated here because it is the feature.
    chart.visible_cells_only = True

    sheet = quote_sheetname(ws.title)
    first, last = get_column_letter(first_col), get_column_letter(last_col)
    # Categories must be a *string* reference: the header holds formatted timestamps, and a
    # numeric reference would leave the x-axis unlabelled.
    categories = f"{sheet}!${first}${start}:${last}${start}"
    for row in range(start + 1, last_row + 1):
        series = Series(Reference(ws, min_col=first_col, max_col=last_col, min_row=row, max_row=row))
        series.val.numRef.f = f"{sheet}!${first}${row}:${last}${row}"  # one row = one store's line
        series.tx = SeriesLabel(strRef=StrRef(f=f"{sheet}!$C${row}"))   # named by its Store cell
        series.cat = AxDataSource(strRef=StrRef(f=categories))
        # Markers are not decoration. A line needs two points, so until a second sync exists
        # an unmarked series draws *nothing at all* — the chart would look broken rather than
        # sparse. They also stay useful afterwards: syncs are irregular, so the markers show
        # where a rate was actually measured as opposed to interpolated between measurements.
        series.marker = Marker(symbol="circle", size=7)
        series.smooth = False  # straight segments: a curve would bend past the measured rates
        chart.series.append(series)

    # openpyxl leaves these unset, and Excel then renders the axes without their tick labels —
    # the numbers and dates vanish while the plot area still draws.
    for axis, position in ((chart.x_axis, "b"), (chart.y_axis, "l")):
        axis.delete = False
        axis.axPos = position
        axis.tickLblPos = "nextTo"
        axis.majorTickMark = "out"
        axis.title.overlay = False  # keep the axis title clear of the tick labels

    ws.add_chart(chart, f"A{last_row + 3}")


def build_workbook(
    *,
    overview: list[dict],
    overview_kpis: dict,
    thresholds: dict,
    products: list[dict],
    history: list[dict],
    meta: dict,
    spike_radar: dict | None = None,
    top_issues: list[dict] | None = None,
    top_issue_totals: dict | None = None,
    pending_issue_codes: list[str] | None = None,
    priority: dict | None = None,
) -> BytesIO:
    """Build the report workbook and return it as a rewound in-memory stream.

    ``overview`` / ``overview_kpis`` are what ``GET /api/overview`` serves; ``thresholds`` is
    ``{"minor", "warning", "critical"}`` as fractions. ``products`` is the flat disapproval table from
    the last sync. ``history`` is every recorded snapshot of every real store, oldest first, each
    row carrying a precomputed ``rate`` and the store's ``region`` / ``country``.
    ``top_issues`` / ``top_issue_totals`` are what ``GET /api/issues/top`` serves.
    """
    wb = Workbook()
    ws_overview = wb.active
    ws_overview.title = SHEET_OVERVIEW
    radar = spike_radar or {"state": "no_sync", "thresholds": {"min_disapproved": 0, "min_diff_pp": 0.0}}
    _build_overview_sheet(ws_overview, overview, overview_kpis, thresholds, meta, radar)
    _build_spike_sheet(wb.create_sheet(SHEET_SPIKES), radar, meta)
    _build_top_issues_sheet(
        wb.create_sheet(SHEET_TOP_ISSUES),
        top_issues or [],
        top_issue_totals or {"disapproved_products": 0, "issue_codes": 0, "pending_products": 0},
        pending_issue_codes or [],
        meta,
    )
    _build_priority_sheet(
        wb.create_sheet(SHEET_PRIORITY),
        priority or {"groups": [], "tag_counts": {"critical": 0, "high": 0, "medium": 0, "low": 0, "pending": 0},
                     "tracking_since": None},
        meta,
    )

    _build_products_sheet(wb.create_sheet(SHEET_PRODUCTS), products, meta)
    _build_history_sheet(wb.create_sheet(SHEET_HISTORY), history, meta)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    logger.info(
        "report built: %d overview store(s), %d product row(s), %d charted store(s), %d history point(s)",
        len(overview),
        len(products),
        len({r["sub_account_id"] for r in history}),
        len(history),
    )
    return buf
