"""
Export helpers for the KISJ Grade Calculator.

Writes a finished report to PDF, JPG, Excel (.xlsx), or CSV (the file you hand
to Google Sheets) with no third-party libraries:

  * CSV   - plain text, opens anywhere.
  * XLSX  - an .xlsx file is a zip of XML parts, written here by hand.
  * PDF   - a small PDF writer using the built-in Helvetica/Courier fonts,
            laid out on A4 portrait pages.
  * JPG   - a one-page, A4-wide PDF rasterized by macOS `sips`.

A "document" is the plain dictionary produced by build_document():

    {"title": str,
     "subtitle": str,
     "tables": [{"name": str,
                 "headers": [str, ...],
                 "rows": [[str | float | None, ...], ...],
                 "bold_rows": {row index, ...}}, ...]}
"""

import csv
import datetime
import io
import os
import shutil
import subprocess
import tempfile
import zipfile


class ExportError(Exception):
    """Raised when an export cannot be completed, with a message for the user."""


# ---------------------------------------------------------------------------
# Document assembly
# ---------------------------------------------------------------------------

def build_document(reports, letter_and_gpa, average, quarter_count):
    """Turn [(class name, report)] into the export document.

    letter_and_gpa and average are passed in so this module stays independent
    of the app module (and easy to test on its own).
    """
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    quarter_labels = ["Q%d (%%)" % n for n in range(1, quarter_count + 1)]

    summary_headers = (["Class"] + quarter_labels +
                       ["Semester 1 (%)", "Semester 2 (%)", "Total (%)", "Letter", "GPA"])
    summary_rows = []
    totals, gpas = [], []

    for name, report in reports:
        quarters, semesters = report["quarters"], report["semesters"]
        row = [name]
        row.extend(_round(quarters.get(n)) for n in range(1, quarter_count + 1))
        row.append(_round(semesters.get(1)))
        row.append(_round(semesters.get(2)))
        if report["total"] is None:
            row.extend([None, "no scores", None])
        else:
            letter, gpa = letter_and_gpa(report["total"])
            totals.append(report["total"])
            gpas.append(gpa)
            row.extend([_round(report["total"]), letter, gpa])
        summary_rows.append(row)

    bold = set()
    if totals:
        combined = average(totals)
        blanks = [None] * (quarter_count + 2)
        summary_rows.append(["ALL CLASSES (%d)" % len(totals)] + blanks +
                            [_round(combined), letter_and_gpa(combined)[0],
                             round(average(gpas), 2)])
        bold.add(len(summary_rows) - 1)

    tables = [{"name": "Summary", "headers": summary_headers,
               "rows": summary_rows, "bold_rows": bold}]

    for name, report in reports:
        rows, bold_rows = [], set()
        for index, row in enumerate(report["rows"]):
            score = row["score"]
            letter, gpa = ("", None)
            if score is not None:
                letter, gpa = letter_and_gpa(score)
            rows.append([row["period"], _round(row["formative"]), _round(row["summative"]),
                         _round(score), letter, gpa, row["note"]])
            if not row["period"].startswith("Quarter"):
                bold_rows.add(index)
        tables.append({
            "name": name,
            "headers": ["Period", "Formative avg", "Summative avg", "Weighted score (%)",
                        "Letter", "GPA", "Note"],
            "rows": rows,
            "bold_rows": bold_rows,
        })

    return {
        "title": "KISJ Grade Report",
        "subtitle": "Generated %s  -  Formative 20%% / Summative 80%%  -  "
                    "unweighted 4.0 GPA scale" % stamp,
        "tables": tables,
    }


def _round(value):
    return None if value is None else round(value, 2)


# ---------------------------------------------------------------------------
# CSV  (also the file used for Google Sheets)
# ---------------------------------------------------------------------------

def csv_text(doc):
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow([doc["title"]])
    writer.writerow([doc["subtitle"]])
    for table in doc["tables"]:
        writer.writerow([])
        writer.writerow([table["name"]])
        writer.writerow(table["headers"])
        for row in table["rows"]:
            writer.writerow(["" if cell is None else cell for cell in row])
    return buffer.getvalue()


def write_csv(path, doc):
    with io.open(path, "w", newline="", encoding="utf-8-sig") as handle:
        handle.write(csv_text(doc))
    return path


# ---------------------------------------------------------------------------
# XLSX  (zip of XML parts, one sheet per table)
# ---------------------------------------------------------------------------

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
    'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.spreadsheetml.styles+xml"/>%s</Types>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
)

STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="3">'
    '<font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="12"/><color rgb="FF1B3A6B"/><name val="Calibri"/></font>'
    '</fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border/></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
    '</cellStyleXfs>'
    '<cellXfs count="3">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '</cellXfs></styleSheet>'
)

STYLE_NORMAL, STYLE_BOLD, STYLE_TITLE = 0, 1, 2


def write_xlsx(path, doc):
    sheets = [(_sheet_name(table["name"], index), table)
              for index, table in enumerate(doc["tables"], start=1)]

    overrides = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % index
        for index in range(1, len(sheets) + 1))

    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><sheets>%s</sheets></workbook>'
                % "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                          % (_escape(name), index, index)
                          for index, (name, _t) in enumerate(sheets, start=1)))

    workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                     'relationships">%s<Relationship Id="rId%d" Type="http://schemas.'
                     'openxmlformats.org/officeDocument/2006/relationships/styles" '
                     'Target="styles.xml"/></Relationships>'
                     % ("".join('<Relationship Id="rId%d" Type="http://schemas.'
                                'openxmlformats.org/officeDocument/2006/relationships/'
                                'worksheet" Target="worksheets/sheet%d.xml"/>'
                                % (index, index)
                                for index in range(1, len(sheets) + 1)),
                        len(sheets) + 1))

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES % overrides)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/styles.xml", STYLES)
        for index, (_name, table) in enumerate(sheets, start=1):
            archive.writestr("xl/worksheets/sheet%d.xml" % index,
                             _sheet_xml(doc, table, first=index == 1))
    return path


def _sheet_xml(doc, table, first):
    rows = []
    line = 1

    def add(values, style):
        cells = []
        for column, value in enumerate(values):
            reference = "%s%d" % (_column_letter(column), line)
            if value is None or value == "":
                continue
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append('<c r="%s" s="%d"><v>%s</v></c>' % (reference, style, value))
            else:
                cells.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">'
                             '%s</t></is></c>' % (reference, style, _escape(str(value))))
        rows.append('<row r="%d">%s</row>' % (line, "".join(cells)))

    if first:
        add([doc["title"]], STYLE_TITLE)
        line += 1
        add([doc["subtitle"]], STYLE_NORMAL)
        line += 2
    add([table["name"]], STYLE_TITLE)
    line += 1
    add(table["headers"], STYLE_BOLD)
    line += 1
    for index, row in enumerate(table["rows"]):
        add(row, STYLE_BOLD if index in table["bold_rows"] else STYLE_NORMAL)
        line += 1

    widths = _column_widths(table)
    cols = "".join('<col min="%d" max="%d" width="%d" customWidth="1"/>'
                   % (index + 1, index + 1, width) for index, width in enumerate(widths))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
            '<cols>%s</cols><sheetData>%s</sheetData>'
            '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" '
            'header="0.3" footer="0.3"/>'
            # paperSize 9 is A4; fit every column onto one page across.
            '<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" '
            'fitToHeight="0"/>'
            '</worksheet>' % (cols, "".join(rows)))


def _column_widths(table):
    widths = []
    for column, header in enumerate(table["headers"]):
        longest = len(str(header))
        for row in table["rows"]:
            if column < len(row) and row[column] is not None:
                longest = max(longest, len(str(row[column])))
        widths.append(min(46, max(10, longest + 2)))
    return widths


def _sheet_name(name, index):
    """Excel sheet names: 31 chars max, no []:*?/\\ and must be unique."""
    cleaned = "".join("-" if character in "[]:*?/\\" else character for character in name)
    cleaned = cleaned.strip()[:28] or "Sheet"
    return "%d %s" % (index, cleaned) if index > 1 else cleaned[:31]


def _column_letter(index):
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _escape(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

PAGE_WIDTH, PAGE_HEIGHT = 595, 842          # A4 portrait (210 x 297 mm)
MARGIN = 40
FONTS = [("F1", "Helvetica"), ("F2", "Helvetica-Bold"),
         ("F3", "Courier"), ("F4", "Courier-Bold")]

MONO_SIZE = 8.0
MONO_WIDTH = MONO_SIZE * 0.6                # Courier advance width


def write_pdf(path, doc, stacked=False, scale=1.0):
    """Write the report.

    scale enlarges the page itself rather than the drawing inside it, so a
    rasteriser renders the text at that size from the vector outlines. Blowing
    up a 72-dpi bitmap afterwards, which is what --resampleWidth does, only
    produces a blurry version of the same small render.
    """
    pages = _paginate(_layout(doc))
    if stacked:
        pages = _stack(pages)
    with io.open(path, "wb") as handle:
        handle.write(_pdf_bytes(pages, scale))
    return path


def table_lines(table):
    """The monospaced lines for one table, shared by the PDF and the previews."""
    widths = _text_widths(table)
    lines = [("row_bold", _text_row(table["headers"], widths)),
             ("rule", "-" * min(_line_capacity(), sum(widths) + 2 * len(widths)))]
    for index, row in enumerate(table["rows"]):
        lines.append(("row_bold" if index in table["bold_rows"] else "row",
                      _text_row(row, widths)))
    return lines


def _layout(doc):
    """Flatten the document into (kind, text) lines."""
    items = [("title", doc["title"]), ("subtitle", doc["subtitle"])]
    for table in doc["tables"]:
        items.append(("gap", ""))
        items.append(("heading", table["name"]))
        items.extend(table_lines(table))
    return items


def _line_capacity():
    return int((PAGE_WIDTH - 2 * MARGIN) / MONO_WIDTH)


def _text_widths(table):
    widths = []
    for column, header in enumerate(table["headers"]):
        longest = len(_cell_text(header))
        for row in table["rows"]:
            if column < len(row):
                longest = max(longest, len(_cell_text(row[column])))
        widths.append(longest)
    return widths


def _cell_text(value):
    if value is None:
        return "-"
    if isinstance(value, float):
        return "%.2f" % value
    return str(value)


def _text_row(values, widths):
    parts = []
    for column, width in enumerate(widths):
        text = _cell_text(values[column]) if column < len(values) else ""
        parts.append(text.ljust(width) if column == 0 else text.rjust(width))
    line = "  ".join(parts)
    return line[:_line_capacity()]


LEADING = {"title": 24, "subtitle": 18, "heading": 20, "row": 11.5,
           "row_bold": 11.5, "rule": 11.5, "gap": 10}
STYLE = {"title": ("F2", 17), "subtitle": ("F1", 9.5), "heading": ("F2", 12),
         "row": ("F3", MONO_SIZE), "row_bold": ("F4", MONO_SIZE),
         "rule": ("F3", MONO_SIZE), "gap": ("F1", 1)}


# Laying the report out on the sheet.
#
# Everything sits at a fixed size in a fixed place: the title at the top margin
# of the first page, the lines one under the next at the spacing set out in
# LEADING above, and the left edge of every line at the left margin. Adding a
# class adds lines to the foot of the report and, when the page is full, a
# further sheet - it never moves or resizes what came before it. A short report
# therefore leaves the lower part of its last sheet empty, which is the price
# of every report being set identically.


def _blocks(items):
    """Group the lines into the heading block and one block per table.

    _layout() puts a gap before every table, so a gap starts a new group.
    """
    blocks, current = [], []
    for kind, text in items:
        if kind == "gap" and current:
            blocks.append(current)
            current = []
        current.append((kind, text))
    if current:
        blocks.append(current)
    return blocks


# A table that has to be broken keeps its heading with at least this many rows,
# so a page never ends with a heading and a line or two stranded under it.
ORPHAN_ROWS = 4


def _split_pages(items):
    """Break (kind, text) lines into one list per A4 page.

    Lines simply run from the top of a page to the bottom and on to the next
    sheet. The one refinement is that a table which would fit on a page of its
    own is moved there whole, rather than leaving one stray row behind.
    """
    limit = PAGE_HEIGHT - 2 * MARGIN
    # Half a point of tolerance: a report that exactly fills the page would
    # otherwise drop its last line onto a sheet of its own.
    slack = 0.5
    pages, page, used = [], [], 0.0

    def start_new_page(block):
        pages.append(page)
        return [], 0.0, block[1:] if block and block[0][0] == "gap" else block

    for block in _blocks(items):
        height = sum(LEADING[kind] for kind, _text in block)
        if page and used + height > limit + slack and height <= limit + slack:
            page, used, block = start_new_page(block)
            height = sum(LEADING[kind] for kind, _text in block)

        if used + height <= limit + slack:
            page.extend(block)
            used += height
            continue

        # Longer than a whole page: lay it out line by line.
        for index, (kind, text) in enumerate(block):
            needed = LEADING[kind]
            if kind == "heading":
                # Take the heading over with its first rows or not at all.
                needed += sum(LEADING[next_kind]
                              for next_kind, _t in block[index + 1:index + 1 + ORPHAN_ROWS])
            if page and used + needed > limit + slack:
                page, used, _block = start_new_page([])
                if kind == "gap":
                    continue        # a gap at the top of a page is a blank line
            page.append((kind, text))
            used += LEADING[kind]

    pages.append(page)
    return [page for page in pages if page]


def _place(items, top):
    """Position (kind, text) lines down a page from `top`, as PDF text runs."""
    placed, y = [], top
    for kind, text in items:
        y -= LEADING[kind]
        if text:
            font, size = STYLE[kind]
            placed.append((MARGIN, y, font, size, text))
    return placed


def _paginate(items):
    """Return [(page height, [(x, y, font, size, text), ...]), ...]."""
    return [(PAGE_HEIGHT, _place(page, PAGE_HEIGHT - MARGIN))
            for page in _split_pages(items)]


def _stack(pages):
    """The pages laid end to end as one tall page, for the JPG.

    The JPG is a single image, but it is an image of A4 sheets: the same pages
    the PDF has, in the same order, one under the next.
    """
    total = sum(height for height, _lines in pages)
    lines, top = [], total
    for height, page_lines in pages:
        base = top - height          # PDF y is measured up from the page foot
        lines.extend((x, y + base, font, size, text)
                     for x, y, font, size, text in page_lines)
        top = base
    return [(total, lines)]


def _pdf_escape(text):
    ascii_text = text.encode("ascii", "replace").decode("ascii")
    return ascii_text.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")


def _pdf_bytes(pages, scale=1.0):
    page_count = len(pages)
    font_first = 3 + 2 * page_count
    font_refs = " ".join("/%s %d 0 R" % (name, font_first + index)
                         for index, (name, _base) in enumerate(FONTS))

    objects = ["<< /Type /Catalog /Pages 2 0 R >>"]
    kids = " ".join("%d 0 R" % (3 + 2 * index) for index in range(page_count))
    objects.append("<< /Type /Pages /Kids [%s] /Count %d >>" % (kids, page_count))

    for index, (height, lines) in enumerate(pages):
        # An explicit white page. Without it the page is transparent, which
        # some PDF viewers show as grey and which makes the rasterised preview
        # take on whatever colour sits behind it.
        body = []
        if scale != 1.0:
            body.append("%.4f 0 0 %.4f 0 0 cm" % (scale, scale))
        body.append("1 1 1 rg 0 0 %.0f %.0f re f" % (PAGE_WIDTH, height))
        body.append("0 0 0 rg")
        body.extend(
            "BT /%s %.1f Tf %.1f %.1f Td (%s) Tj ET" % (font, size, x, y, _pdf_escape(text))
            for x, y, font, size, text in lines)
        content = "\n".join(body)
        objects.append(
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %.0f %.0f] /Resources "
            "<< /Font << %s >> >> /Contents %d 0 R >>"
            % (PAGE_WIDTH * scale, height * scale, font_refs, 4 + 2 * index))
        objects.append("<< /Length %d >>\nstream\n%s\nendstream"
                       % (len(content.encode("latin-1")) + 1, content))

    for _name, base in FONTS:
        objects.append("<< /Type /Font /Subtype /Type1 /BaseFont /%s "
                       "/Encoding /WinAnsiEncoding >>" % base)

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += ("%d 0 obj\n" % number).encode("latin-1")
        out += body.encode("latin-1")
        out += b"\nendobj\n"

    xref_at = len(out)
    out += ("xref\n0 %d\n" % (len(objects) + 1)).encode("latin-1")
    out += b"0000000000 65535 f \n"
    for offset in offsets:
        out += ("%010d 00000 n \n" % offset).encode("latin-1")
    out += ("trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n"
            % (len(objects) + 1, xref_at)).encode("latin-1")
    return bytes(out)


# ---------------------------------------------------------------------------
# JPG  (one tall page, rasterized by macOS sips)
# ---------------------------------------------------------------------------

# A4 portrait is 210 mm wide, so 200 dpi of it is 210 / 25.4 * 200 pixels.
A4_JPG_WIDTH = int(round(PAGE_WIDTH / 72.0 * 200))


def jpg_available():
    return shutil.which("sips") is not None


def write_jpg(path, doc, pixel_width=A4_JPG_WIDTH):
    if not jpg_available():
        raise ExportError(
            "Saving as JPG needs the macOS `sips` tool, which was not found on this "
            "computer. PDF, Excel and CSV all still work.")

    handle, temp_pdf = tempfile.mkstemp(suffix=".pdf")
    os.close(handle)
    try:
        write_pdf(temp_pdf, doc, stacked=True,
                  scale=float(pixel_width) / PAGE_WIDTH)
        result = subprocess.run(
            ["sips", "-s", "format", "jpeg", "-s", "formatOptions", "best",
             temp_pdf, "--out", path],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode != 0 or not os.path.exists(path):
            raise ExportError("The JPG could not be created:\n%s"
                              % result.stderr.decode("utf-8", "replace").strip())
    finally:
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)
    return path


# ---------------------------------------------------------------------------
# Previews - what the saved file will contain, before it is written
# ---------------------------------------------------------------------------

def page_count(doc):
    return len(_paginate(_layout(doc)))


def layout_blocks(doc):
    """(kind, line) pairs describing the report, for a styled preview."""
    return _layout(doc)


def placed_pages(doc):
    """The report drawn to scale, for a preview that looks like the real page.

    Returns [(page width, page height, [(x, y from the top, font, size, text),
    ...]), ...] in points, with y measured downwards from the top-left corner
    the way a screen canvas measures it. The JPG holds these same pages, one
    under the next, so the preview is the same either way.
    """
    names = dict(FONTS)
    pages = []
    for height, lines in _paginate(_layout(doc)):
        pages.append((PAGE_WIDTH, height,
                      [(x, height - y, names[font], size, text)
                       for x, y, font, size, text in lines]))
    return pages


def page_shape():
    """(width, height, margin) of a page, in points, for a to-scale preview."""
    return PAGE_WIDTH, PAGE_HEIGHT, MARGIN


def xlsx_blocks(doc):
    blocks = [("title", doc["title"]), ("subtitle", doc["subtitle"])]
    for index, table in enumerate(doc["tables"], start=1):
        blocks.append(("gap", ""))
        blocks.append(("heading", "Sheet %d:  %s" % (index, _sheet_name(table["name"], index))))
        blocks.extend(table_lines(table))
    return blocks


def csv_blocks(doc):
    blocks = [("title", "CSV file contents"),
              ("subtitle", "Exactly the text that will be saved.")]
    blocks.extend(("row", line) for line in csv_text(doc).splitlines())
    return blocks


def preview_text(doc):
    """Exactly the text the PDF and JPG will carry."""
    lines = [doc["title"], doc["subtitle"]]
    for table in doc["tables"]:
        lines.append("")
        lines.append(table["name"])
        lines.extend(text for _kind, text in table_lines(table))
    return "\n".join(lines)


def xlsx_preview(doc):
    """Sheet by sheet, the way Excel will open it."""
    lines = ["%s\n%s" % (doc["title"], doc["subtitle"]), ""]
    for index, table in enumerate(doc["tables"], start=1):
        lines.append("+-- Sheet %d: %s" % (index, _sheet_name(table["name"], index)))
        lines.extend("|  " + text for _kind, text in table_lines(table))
        lines.append("")
    return "\n".join(lines)


def describe(kind, doc):
    """One line about the file that is about to be written."""
    if kind == "pdf":
        pages = page_count(doc)
        return "%d page%s, A4 portrait (210 x 297 mm), text you can select and search." % (
            pages, "" if pages == 1 else "s")
    if kind == "jpg":
        pages = page_count(doc)
        held = ("the whole A4 page" if pages == 1
                else "all %d A4 pages, one under the next" % pages)
        return "One image %d pixels wide (A4 width at 200 dpi), holding %s." % (
            A4_JPG_WIDTH, held)
    if kind == "xlsx":
        names = ", ".join(table["name"] for table in doc["tables"])
        return "%d sheets (%s). Scores are real numbers, so Excel can calculate with them." % (
            len(doc["tables"]), names)
    return ("One CSV file to import into Google Sheets (File -> Import -> Upload). "
            "Scores stay as numbers.")


# preview_image() used to rasterise the report for the preview window. It was
# removed: Tk draws a PhotoImage one image pixel per point, so on a Retina
# display every preview was scaled up by the compositor and looked blurry no
# matter how large it was rendered. The preview now uses real text, which the
# system draws at full device resolution.
