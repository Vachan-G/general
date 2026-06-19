"""
split_opp_column.py

Reads OPP_Quality_Analysis_MiCo_IAM.xlsx and produces OPP_Quality_Analysis_Updated.xlsx.

For the 'OPP Quality Analysis' data sheet:
  - Keeps cols 1-4 (original columns A-D: #, Workstream, Service, OPP Document)
  - Splits col 5 ("Part that needs to be updated") into two new columns:
      Col 5: "Gap/Problem of the OPP"
      Col 6: "Update/Change to be made to the OPP"

Split logic:
  - Tokenise the cell text into sentences/clauses by splitting on '. ' and '\n'.
  - Find the first token that starts with an action verb (case-insensitive):
    Add, Update, Include, Ensure, Remove, Change, Modify, Replace, Consider, Revise
  - Everything before that token → Gap column
  - That token and everything after → Update column
  - If no action verb found, entire text → Gap column, Update column is empty

Formatting:
  - Header row (row 3): dark blue background (1F3864), white bold font, wrap text
  - Data rows: OPP colour on ALL cells in that row (same fill as original col 4)
  - Wrap text on all data cells
  - Row height 120 for all data rows
  - Freeze panes at A4 (below the two title rows and the header row)
  - Column width 40 for all columns
  - Rows 1-2 (title/subtitle merged rows) preserved as-is

Summary sheet copied unchanged.
"""

import re
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

SRC = '/root/.claude/uploads/c57d67a1-7ab7-5d6e-b9c2-3e553b76b6f0/b64a28d4-OPP_Quality_Analysis_MiCo_IAM.xlsx'
DST = '/home/user/general/OPP_Quality_Analysis_Updated.xlsx'

ACTION_VERBS = [
    'Add', 'Update', 'Include', 'Ensure', 'Remove',
    'Change', 'Modify', 'Replace', 'Consider', 'Revise',
    'Restructure', 'Align', 'Expand', 'Create', 'Insert',
    'Introduce', 'Clarify', 'Define', 'Document', 'Extend',
    'Merge', 'Move', 'Rename', 'Rewrite', 'Split',
]
# Build a regex pattern that matches at the start of a sentence/token
VERB_PATTERN = re.compile(
    r'^(' + '|'.join(ACTION_VERBS) + r')\b',
    re.IGNORECASE
)

HEADER_FILL = PatternFill(fill_type='solid', fgColor='1F3864')
HEADER_FONT = Font(bold=True, color='FFFFFF', name='Calibri', size=11)
WRAP_ALIGN = Alignment(wrap_text=True, vertical='top')

NEW_HEADERS = [
    '#',
    'Workstream',
    'Service',
    'OPP Document',
    'Gap/Problem of the OPP',
    'Update/Change to be made to the OPP',
]


def split_text(text):
    """
    Split text into (gap, update) by finding the first sentence/clause that
    begins with an action verb at a sentence boundary.

    Sentence boundaries are: start of text, after '\n', or after '. '.
    The verb must be the FIRST word of the sentence (after stripping whitespace).

    Returns (gap_text, update_text).
    """
    if not text:
        return ('', '')

    # We'll scan the text character by character to find sentence boundaries
    # and check if the sentence starting there begins with an action verb.
    # Sentence boundaries:
    #   - position 0 (start)
    #   - right after a '\n'
    #   - right after '. ' (period + space)

    # Build list of (start_char_index, sentence_text) by splitting carefully.
    # Strategy: collect boundary positions, then extract candidate sentences.

    boundaries = [0]  # indices where a new sentence starts

    i = 0
    while i < len(text):
        if text[i] == '\n':
            if i + 1 < len(text):
                boundaries.append(i + 1)
        elif text[i] == '.':
            # Handle '. ' and also '.' followed by quote/apostrophe then space
            # e.g.  .' Add ...  or  ." Add ...
            j = i + 1
            # skip closing quote characters
            while j < len(text) and text[j] in ('"', "'", '’', '”'):
                j += 1
            if j < len(text) and text[j] == ' ' and j + 1 < len(text):
                boundaries.append(j + 1)
        i += 1

    # For each boundary, check if the sentence starting there begins with an action verb
    split_pos = None
    for b in boundaries:
        if b == 0:
            continue  # don't split at position 0 (nothing would go to gap)
        remainder = text[b:]
        stripped = remainder.lstrip()
        if VERB_PATTERN.match(stripped):
            # Make sure this is truly at a sentence start, not mid-sentence
            # (boundary already guarantees this — b follows \n or '. ')
            split_pos = b
            break

    if split_pos is None:
        return (text.strip(), '')

    gap = text[:split_pos].strip()
    update = text[split_pos:].strip()
    return (gap, update)


def make_fill(rgb_with_prefix):
    """Create PatternFill from a full 8-char RGB like '00FDEBD0'."""
    # Strip leading '00' alpha prefix if present, get last 6 chars
    color = rgb_with_prefix[-6:]
    return PatternFill(fill_type='solid', fgColor=color)


def copy_cell_style(src_cell, dst_cell):
    """Copy fill, font, alignment, border from src to dst."""
    if src_cell.fill:
        dst_cell.fill = copy(src_cell.fill)
    if src_cell.font:
        dst_cell.font = copy(src_cell.font)
    if src_cell.alignment:
        dst_cell.alignment = copy(src_cell.alignment)
    if src_cell.border:
        dst_cell.border = copy(src_cell.border)


def process_data_sheet(src_ws, dst_ws):
    """
    Process the 'OPP Quality Analysis' sheet:
    - Rows 1-2: merged title rows, copy as-is
    - Row 3: header row, write new headers with dark blue style, 6 columns
    - Rows 4+: data rows, split col 5 into cols 5+6
    """

    # --- Rows 1-2: Copy title/subtitle merged rows ---
    for row_num in (1, 2):
        src_row = list(src_ws.iter_rows(min_row=row_num, max_row=row_num))[0]
        for src_cell in src_row:
            dst_cell = dst_ws.cell(row=row_num, column=src_cell.column)
            dst_cell.value = src_cell.value
            copy_cell_style(src_cell, dst_cell)
        dst_ws.row_dimensions[row_num].height = src_ws.row_dimensions[row_num].height or 15

    # Merge title rows across 6 columns (was 5)
    # First remove any existing merges then add new ones
    dst_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    dst_ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)

    # --- Row 3: Header row ---
    for col_idx, header in enumerate(NEW_HEADERS, start=1):
        cell = dst_ws.cell(row=3, column=col_idx)
        cell.value = header
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP_ALIGN
    dst_ws.row_dimensions[3].height = 30

    # --- Rows 4+: Data rows ---
    sample_splits = []  # collect a few for reporting

    for src_row in src_ws.iter_rows(min_row=4, max_row=src_ws.max_row):
        row_num = src_row[0].row

        # Get OPP Document fill (col 4, index 3) for row colour coding
        opp_fill = copy(src_row[3].fill) if src_row[3].fill else None

        # Extract values for cols 1-4
        vals = [src_row[i].value for i in range(4)]
        # Col 5 text to split
        raw_text = src_row[4].value if len(src_row) > 4 else None

        gap, update = split_text(raw_text)

        if len(sample_splits) < 3 and raw_text:
            sample_splits.append({
                'row': row_num,
                'gap_preview': gap[:120],
                'update_preview': update[:120],
            })

        row_values = vals + [gap, update]

        for col_idx, val in enumerate(row_values, start=1):
            cell = dst_ws.cell(row=row_num, column=col_idx)
            cell.value = val
            cell.alignment = WRAP_ALIGN
            if opp_fill and opp_fill.patternType == 'solid':
                cell.fill = copy(opp_fill)

        dst_ws.row_dimensions[row_num].height = 120

    # --- Column widths ---
    for col_idx in range(1, 7):
        dst_ws.column_dimensions[get_column_letter(col_idx)].width = 40

    # --- Freeze panes at A4 (below header) ---
    dst_ws.freeze_panes = 'A4'

    return sample_splits


def copy_summary_sheet(src_ws, dst_ws):
    """Copy the Summary sheet verbatim including merged cells and styles."""
    # Copy merged cell ranges first
    for merge in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merge))

    for row in src_ws.iter_rows():
        for src_cell in row:
            # Skip merged cells (read-only proxies); only write to top-left of each merge
            if src_cell.__class__.__name__ == 'MergedCell':
                continue
            dst_cell = dst_ws.cell(row=src_cell.row, column=src_cell.column)
            dst_cell.value = src_cell.value
            copy_cell_style(src_cell, dst_cell)

    # Copy row heights and column widths
    for row_num, rd in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_num].height = rd.height

    for col_letter, cd in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = cd.width

    dst_ws.freeze_panes = src_ws.freeze_panes


def main():
    print(f"Loading source: {SRC}")
    src_wb = load_workbook(SRC)

    dst_wb = Workbook()
    # Remove default sheet
    dst_wb.remove(dst_wb.active)

    # Process data sheet
    src_data_ws = src_wb['OPP Quality Analysis']
    dst_data_ws = dst_wb.create_sheet('OPP Quality Analysis')
    print("Processing 'OPP Quality Analysis' sheet...")
    sample_splits = process_data_sheet(src_data_ws, dst_data_ws)

    # Copy summary sheet
    src_summary_ws = src_wb['Summary by OPP']
    dst_summary_ws = dst_wb.create_sheet('Summary by OPP')
    print("Copying 'Summary by OPP' sheet...")
    copy_summary_sheet(src_summary_ws, dst_summary_ws)

    dst_wb.save(DST)
    print(f"\nSaved output: {DST}")

    # Report
    print("\n--- Sample Split Results ---")
    for s in sample_splits:
        print(f"\nRow {s['row']}:")
        print(f"  GAP    : {s['gap_preview']!r}")
        print(f"  UPDATE : {s['update_preview']!r}")

    print("\nDone.")


if __name__ == '__main__':
    main()
