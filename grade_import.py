"""
Import helpers for the KISJ Grade Calculator.

Reads a report that this program saved earlier - PDF, JPG/PNG, Excel (.xlsx)
or CSV - and recovers the classes, the number of quarters, and every score
that was typed in, so the calculator can be filled in again and corrected or
continued rather than typed from scratch. No third-party libraries:

  * CSV   - the csv module.
  * XLSX  - a zip of XML parts, read with zipfile and xml.etree.
  * PDF   - on macOS, PDFKit through osascript; elsewhere a small reader for
            the PDFs this program writes (plain or Flate-compressed streams).
  * JPG   - text recognition. On macOS the Vision framework, again through
            osascript, so nothing has to be installed; elsewhere the
            `tesseract` command if it is on the PATH.

Every format ends up as plain text lines, and the same parser reads those.
The "Scores entered" table at the foot of a report lists each score exactly
as typed; older reports (saved before that table existed) only carry quarter
averages, which are imported as one rounded score each and flagged.

read_report() returns:

    {"source": file name,
     "classes": [class name, ...],              in the order found
     "quarter_count": 1..4,
     "scores": {class name: {quarter: ([formatives], [summatives])}},
     "approximate": bool,                       True when built from averages
     "recognized": bool,                        True when read from a picture
     "warnings": [str, ...]}
"""

import csv
import difflib
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import zlib
from xml.etree import ElementTree


class ReadError(Exception):
    """Raised when a file cannot be imported, with a message for the user."""


IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def read_report(path, known_classes=()):
    """Read one saved report. known_classes fixes small spelling differences
    (a picture read by text recognition is rarely letter-perfect)."""
    if not os.path.exists(path):
        raise ReadError("The file could not be found:\n%s" % path)
    extension = os.path.splitext(path)[1].lower()

    recognized = False
    if extension == ".csv":
        lines = _grid_lines(_csv_rows(path))
    elif extension == ".xlsx":
        lines = _grid_lines(_xlsx_rows(path))
    elif extension == ".pdf":
        lines, recognized = _pdf_lines(path)
    elif extension in IMAGE_EXTENSIONS:
        lines, recognized = _image_lines(path), True
    else:
        raise ReadError("Choose a PDF, JPG, Excel (.xlsx) or CSV file - the same "
                        "kinds this calculator saves.")

    result = parse_lines(lines, known_classes)
    if not result["classes"]:
        raise ReadError(
            "No KISJ grade report was found in %s.\n\nImport a file that this "
            "calculator saved (PDF, JPG, Excel or CSV)%s."
            % (os.path.basename(path),
               " - and if it is a picture, make sure the text is sharp and level"
               if recognized else ""))
    result["source"] = os.path.basename(path)
    result["recognized"] = recognized
    if recognized:
        result["warnings"].insert(
            0, "The numbers were read from a picture, so please double-check "
               "every one of them.")
    return result


# ---------------------------------------------------------------------------
# Parsing the report text
# ---------------------------------------------------------------------------

# One row of the "Scores entered" table:  <class>  Quarter 1  Formative  90, 85
ENTERED_ROW = re.compile(
    r"^(?P<name>.+?)\s+Q(?:uarter)?\s*(?P<quarter>[1-4])\s+"
    r"(?P<category>Formative|Summative)\b\s*(?P<scores>.*)$", re.IGNORECASE)

# A row of a class table in a report saved before scores were listed:
#   Quarter 1  87.50  90.00  89.50  A-  3.70
AVERAGE_ROW = re.compile(r"^Q(?:uarter)?\s*(?P<quarter>[1-4])\s+(?P<rest>.*)$",
                         re.IGNORECASE)
TABLE_HEADER = re.compile(r"^Period\b", re.IGNORECASE)
NUMBER = re.compile(r"^\d+(?:[.,]\d+)?$")


def parse_lines(lines, known_classes=()):
    """Turn the lines of a report into the read_report() structure."""
    lines = [_clean(line) for line in lines]
    lines = [line for line in lines if line]

    scores, order, warnings = _entered_scores(lines)
    approximate = False
    if not order:
        scores, order, warnings = _average_scores(lines)
        approximate = bool(order)
        if approximate:
            warnings.insert(
                0, "This report only lists quarter averages, not the individual "
                   "scores, so each average was entered as one (rounded) score. "
                   "Retype the real scores if you have them.")

    classes, scores = _match_names(order, scores, known_classes, warnings)

    quarter_count = 1
    for per_quarter in scores.values():
        quarter_count = max([quarter_count] + list(per_quarter))
    for name in classes:
        for number in range(1, quarter_count + 1):
            scores[name].setdefault(number, ([], []))

    return {"classes": classes, "quarter_count": quarter_count, "scores": scores,
            "approximate": approximate, "warnings": warnings}


def _clean(line):
    """Recognised text is not letter-perfect: unify the dashes and quotes it
    produces and squeeze the spacing."""
    line = (line or "").replace("–", "-").replace("—", "-")
    line = line.replace("‘", "'").replace("’", "'")
    return re.sub(r"\s+", " ", line).strip()


def _entered_scores(lines):
    """Read the "Scores entered" table, wherever its rows are on the page."""
    scores, order, warnings = {}, [], []
    for line in lines:
        match = ENTERED_ROW.match(line)
        if not match:
            continue
        name = match.group("name").strip(" -:|")
        if not name or name.lower() == "class":
            continue
        quarter = int(match.group("quarter"))
        column = 0 if match.group("category").lower() == "formative" else 1
        values, doubtful = _score_list(match.group("scores"))
        if doubtful:
            warnings.append("%s, Quarter %d, %s: part of the score list could not "
                            "be read (\"%s\")." % (name, quarter, match.group("category"),
                                                   match.group("scores").strip()))
        if name not in scores:
            scores[name] = {}
            order.append(name)
        # A long list is written over several rows; they simply run on.
        pair = scores[name].setdefault(quarter, ([], []))
        pair[column].extend(values)
    return scores, order, warnings


def _score_list(text):
    """"90, 85, 100" -> ([90, 85, 100], doubtful). A dash means no scores."""
    text = text.strip()
    if not text or text in ("-", ":", ".", "•"):
        return [], False
    values = [int(part) for part in re.findall(r"\d+", text)]
    values = [value for value in values if 0 <= value <= 100]
    leftovers = re.sub(r"[\d,\s.;/|-]", "", text)
    return values, bool(leftovers) or not values


def _average_scores(lines):
    """Fallback for reports that predate the "Scores entered" table: take the
    formative and summative averages from each class table instead."""
    scores, order, warnings = {}, [], []
    current = None
    for index, line in enumerate(lines):
        if TABLE_HEADER.match(line):
            # The class name is the line above the table header.
            name = lines[index - 1] if index else ""
            name = _strip_sheet_prefix(name)
            if not name or name.lower() == "summary":
                current = None
                continue
            current = name
            if name not in scores:
                scores[name] = {}
                order.append(name)
            continue
        if current is None:
            continue
        match = AVERAGE_ROW.match(line)
        if not match:
            if re.match(r"^(Semester|Total|Year)\b", line, re.IGNORECASE):
                continue
            if not re.match(r"^-+$", line):
                current = None      # the next class heading, or anything else
            continue
        quarter = int(match.group("quarter"))
        tokens = match.group("rest").split()
        formative = _as_whole(tokens[0] if tokens else "")
        summative = _as_whole(tokens[1] if len(tokens) > 1 else "")
        scores[current][quarter] = ([formative] if formative is not None else [],
                                    [summative] if summative is not None else [])
    return scores, order, warnings


def _strip_sheet_prefix(name):
    """Excel sheet names carry a number: "2 AP Calculus AB" -> "AP Calculus AB"."""
    return re.sub(r"^\d+\s+", "", name).strip()


def _as_whole(token):
    """"87.50" -> 88 (rounded half up); a dash or a word -> None."""
    token = token.replace(",", ".") if token.count(",") == 1 and "." not in token else token
    if not NUMBER.match(token):
        return None
    value = int(float(token) + 0.5)
    return value if 0 <= value <= 100 else None


def _match_names(order, scores, known_classes, warnings):
    """Snap each class name to the course list where the spelling is close."""
    known = list(known_classes)
    lookup = {name.lower(): name for name in known}
    classes, matched = [], {}
    for raw in order:
        name = raw
        if raw in known:
            pass
        elif raw.lower() in lookup:
            name = lookup[raw.lower()]
        else:
            close = difflib.get_close_matches(raw, known, n=1, cutoff=0.8)
            if close:
                name = close[0]
                warnings.append("\"%s\" was read as %s." % (raw, name))
            elif known:
                warnings.append("\"%s\" is not on the course list - it was kept as "
                                "written." % raw)
        if name in matched:
            # The same class twice (e.g. a garbled duplicate): merge the scores.
            for quarter, (formatives, summatives) in scores[raw].items():
                pair = matched[name].setdefault(quarter, ([], []))
                pair[0].extend(formatives)
                pair[1].extend(summatives)
            continue
        matched[name] = scores[raw]
        classes.append(name)
    return classes, matched


# ---------------------------------------------------------------------------
# Describing what was found
# ---------------------------------------------------------------------------

def summary_text(result):
    """A short, plain description of the import for a confirmation dialog."""
    lines = ["Found in %s:" % result["source"], ""]
    for name in result["classes"]:
        parts = []
        for quarter in range(1, result["quarter_count"] + 1):
            formatives, summatives = result["scores"][name].get(quarter, ([], []))
            if not formatives and not summatives:
                parts.append("Q%d: nothing" % quarter)
            else:
                parts.append("Q%d: %d formative, %d summative"
                             % (quarter, len(formatives), len(summatives)))
        lines.append("• %s  -  %s" % (name, " · ".join(parts)))
    count = result["quarter_count"]
    lines.append("")
    lines.append("Quarters completed: %d (%s)" % (
        count, {1: "Q1", 2: "Q1-Q2 = Semester 1", 3: "Q1-Q3", 4: "full year"}[count]))
    if result["warnings"]:
        lines.append("")
        lines.append("Please note:")
        lines.extend("• " + warning for warning in result["warnings"])
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CSV and XLSX  ->  rows of cells  ->  lines
# ---------------------------------------------------------------------------

def _grid_lines(rows):
    """Join each row's cells into one line, with a dash for an empty cell so
    the columns still line up the way they do in the PDF."""
    lines = []
    for row in rows:
        cells = [str(cell).strip() or "-" for cell in row]
        while cells and cells[-1] == "-":
            cells.pop()
        lines.append("  ".join(cells))
    return lines


def _csv_rows(path):
    try:
        with io.open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as handle:
            text = handle.read()
    except OSError as error:
        raise ReadError("The file could not be opened:\n%s" % error)
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return list(csv.reader(io.StringIO(text), dialect))


XML_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XML_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _xlsx_rows(path):
    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        raise ReadError("This is not an Excel workbook that can be opened.")
    with archive:
        names = set(archive.namelist())
        if "xl/workbook.xml" not in names:
            raise ReadError("This is not an Excel workbook that can be opened.")

        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.iter(XML_MAIN + "si"):
                shared.append("".join(t.text or "" for t in item.iter(XML_MAIN + "t")))

        targets = {}
        if "xl/_rels/workbook.xml.rels" in names:
            for rel in ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels")):
                targets[rel.get("Id")] = rel.get("Target", "")

        rows = []
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        for sheet in workbook.iter(XML_MAIN + "sheet"):
            target = targets.get(sheet.get(XML_REL_ID), "")
            part = target.lstrip("/") if target.startswith("/") else "xl/" + target
            if part not in names:
                continue
            root = ElementTree.fromstring(archive.read(part))
            for row in root.iter(XML_MAIN + "row"):
                cells, last = {}, -1
                for cell in row.iter(XML_MAIN + "c"):
                    reference = cell.get("r")
                    column = _column_index(reference) if reference else last + 1
                    cells[column] = _cell_value(cell, shared)
                    last = column
                if cells:
                    rows.append([cells.get(i, "") for i in range(max(cells) + 1)])
            rows.append([])
        return rows


def _column_index(reference):
    index = 0
    for character in reference:
        if not character.isalpha():
            break
        index = index * 26 + (ord(character.upper()) - 64)
    return index - 1


def _cell_value(cell, shared):
    kind = cell.get("t", "n")
    if kind == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(XML_MAIN + "t"))
    value = cell.find(XML_MAIN + "v")
    text = (value.text or "") if value is not None else ""
    if kind == "s":
        try:
            return shared[int(text)]
        except (ValueError, IndexError):
            return ""
    if kind in ("str", "b", "e"):
        return text
    try:
        number = float(text)
    except ValueError:
        return text
    return str(int(number)) if number == int(number) else repr(number)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

# Extracts the text layer of a PDF with PDFKit, one line per line of the page.
PDF_TEXT_SCRIPT = r"""
ObjC.import('Foundation');
ObjC.import('Quartz');
function run(argv) {
  var url = $.NSURL.fileURLWithPath(argv[0]);
  var doc = $.PDFDocument.alloc.initWithURL(url);
  if (doc.isNil()) return "ERROR: not a PDF";
  var out = [];
  for (var i = 0; i < doc.pageCount; i++) {
    var text = doc.pageAtIndex(i).string;
    out.push(text.isNil() ? "" : ObjC.unwrap(text));
  }
  return out.join("\n");
}
"""


def _pdf_lines(path):
    """(lines, recognized). A PDF with no text layer - a scan, or a picture
    saved as PDF - is rasterised and read like a JPG."""
    text = None
    if sys.platform == "darwin" and shutil.which("osascript"):
        try:
            text = _run_jxa(PDF_TEXT_SCRIPT, path)
        except ReadError:
            text = None
    if text is None:
        text = _pdf_text_fallback(path)
    lines = text.splitlines()
    if any(_clean(line) for line in lines):
        return lines, False

    image = _rasterize_pdf(path)
    if image is None:
        raise ReadError("This PDF has no readable text in it (it may be a scan). "
                        "Import the JPG, Excel or CSV version instead.")
    try:
        return _image_lines(image), True
    finally:
        shutil.rmtree(os.path.dirname(image), ignore_errors=True)


def _rasterize_pdf(path):
    """First page of the PDF as a PNG, via macOS Quick Look. None elsewhere."""
    if sys.platform != "darwin" or not shutil.which("qlmanage"):
        return None
    folder = tempfile.mkdtemp(prefix="kisj-import-")
    result = subprocess.run(["qlmanage", "-t", "-s", "2400", "-o", folder, path],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    produced = [name for name in os.listdir(folder) if name.lower().endswith(".png")]
    if result.returncode != 0 or not produced:
        shutil.rmtree(folder, ignore_errors=True)
        return None
    return os.path.join(folder, produced[0])


def _pdf_text_fallback(path):
    """Read the text runs out of a PDF's content streams.

    Enough for the PDFs this program writes and for most that other programs
    re-save: plain or Flate-compressed streams, Tj/TJ text operators, and a
    new line whenever the text position moves down the page.
    """
    try:
        with io.open(path, "rb") as handle:
            data = handle.read()
    except OSError as error:
        raise ReadError("The file could not be opened:\n%s" % error)
    if not data.startswith(b"%PDF"):
        raise ReadError("This is not a PDF file that can be opened.")

    lines = []
    for match in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", data, re.DOTALL):
        body = match.group(1)
        head = data[max(0, match.start() - 300):match.start()]
        if b"FlateDecode" in head:
            try:
                body = zlib.decompress(body)
            except zlib.error:
                continue
        if b"BT" not in body:
            continue
        lines.extend(_content_lines(body.decode("latin-1")))
    return "\n".join(lines)


def _content_lines(content):
    """Text runs of one content stream, grouped into lines by y position."""
    lines, current, current_y = [], [], None
    y = 0.0
    tokens = re.findall(r"\((?:\\.|[^\\)])*\)|\[[^\]]*\]|[^\s\[\]()]+", content)
    operands = []
    for token in tokens:
        if token.startswith("(") or token.startswith("["):
            operands.append(token)
            continue
        try:
            operands.append(float(token))
            continue
        except ValueError:
            pass
        operator = token
        if operator == "Td" or operator == "TD":
            if len(operands) >= 2 and isinstance(operands[-1], float):
                y += operands[-1]
        elif operator == "Tm":
            if len(operands) >= 6 and isinstance(operands[-1], float):
                y = operands[-1]
        elif operator in ("T*", "'", '"'):
            y -= 1000.0     # any move down starts a new line
        if operator in ("Tj", "'", '"', "TJ"):
            text = "".join(_pdf_string(part) for part in operands
                           if isinstance(part, str))
            if current_y is None or abs(y - current_y) > 1.0:
                if current:
                    lines.append(" ".join(current))
                current, current_y = [], y
            if text.strip():
                current.append(text.strip())
        operands = []
    if current:
        lines.append(" ".join(current))
    return lines


def _pdf_string(token):
    """The characters inside a (string) or a [kerned array] of them."""
    if token.startswith("["):
        return "".join(_pdf_string(part)
                       for part in re.findall(r"\((?:\\.|[^\\)])*\)", token))
    inner = token[1:-1]
    return re.sub(r"\\([nrtbf()\\]|[0-7]{1,3})", _pdf_unescape, inner)


def _pdf_unescape(match):
    code = match.group(1)
    if code.isdigit():
        return chr(int(code, 8))
    return {"n": "\n", "r": "\r", "t": "\t", "b": "\b", "f": "\f"}.get(code, code)


# ---------------------------------------------------------------------------
# Pictures
# ---------------------------------------------------------------------------

# Recognises the text in a picture with the Vision framework and prints one
# fragment per line as: x, y, width, height (fractions of the picture, y up
# from the bottom) and the text.
#
# Vision shrinks a big picture before reading it, and the JPG this program
# saves is several A4 pages tall, so the text became too small to read. The
# picture is therefore read one page-sized strip at a time. Strips overlap by
# more than a line of text and each fragment is taken from the strip whose
# middle part it lies in, so nothing is cut in half or counted twice.
OCR_SCRIPT = r"""
ObjC.import('Foundation');
ObjC.import('AppKit');
ObjC.import('Vision');
function recognise(image) {
  var request = $.VNRecognizeTextRequest.alloc.init;
  request.recognitionLevel = $.VNRequestTextRecognitionLevelAccurate;
  request.usesLanguageCorrection = false;
  request.recognitionLanguages = $.NSArray.arrayWithObject($('en-US'));
  var handler = $.VNImageRequestHandler.alloc.initWithCGImageOptions(image, $.NSDictionary.dictionary);
  if (!handler.performRequestsError($.NSArray.arrayWithObject(request), Ref())) return null;
  return request.results;
}
function run(argv) {
  var url = $.NSURL.fileURLWithPath(argv[0]);
  var source = $.CGImageSourceCreateWithURL(url, $());
  if (!source) return "ERROR: cannot read picture";
  var image = $.CGImageSourceCreateImageAtIndex(source, 0, $());
  if (!image) return "ERROR: cannot read picture";
  var width = $.CGImageGetWidth(image), height = $.CGImageGetHeight(image);
  var strip = Math.round(width * 842 / 595);      // an A4 page at this width
  var overlap = Math.round(width * 0.05);
  var out = [];
  for (var top = 0; top < height; top += strip - overlap) {
    var tall = Math.min(strip, height - top);
    var piece = $.CGImageCreateWithImageInRect(image, $.CGRectMake(0, top, width, tall));
    var results = recognise(piece);
    if (!results) return "ERROR: recognition failed";
    var ownTop = top === 0 ? 0 : top + overlap / 2;
    var ownBottom = top + tall >= height ? height : top + tall - overlap / 2;
    for (var i = 0; i < results.count; i++) {
      var found = results.objectAtIndex(i), box = found.boundingBox;
      var pixelTop = top + (1 - box.origin.y - box.size.height) * tall;
      var pixelHeight = box.size.height * tall;
      var middle = pixelTop + pixelHeight / 2;
      if (middle < ownTop || middle >= ownBottom) continue;
      var text = found.topCandidates(1).objectAtIndex(0).string;
      out.push([box.origin.x.toFixed(5), (1 - (pixelTop + pixelHeight) / height).toFixed(5),
                box.size.width.toFixed(5), (pixelHeight / height).toFixed(5),
                ObjC.unwrap(text)].join("\t"));
    }
    if (top + tall >= height) break;
  }
  return out.join("\n");
}
"""


def ocr_available():
    if sys.platform == "darwin" and shutil.which("osascript"):
        return True
    return shutil.which("tesseract") is not None


def _image_lines(path):
    if sys.platform == "darwin" and shutil.which("osascript"):
        return _assemble_lines(_run_jxa(OCR_SCRIPT, path))
    if shutil.which("tesseract"):
        result = subprocess.run(["tesseract", path, "stdout", "--psm", "6"],
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode != 0:
            raise ReadError("The picture could not be read:\n%s"
                            % result.stderr.decode("utf-8", "replace").strip())
        return result.stdout.decode("utf-8", "replace").splitlines()
    raise ReadError("Reading a picture needs macOS text recognition (or the "
                    "`tesseract` tool), and neither is available here. Import the "
                    "PDF, Excel or CSV version of the report instead.")


def _assemble_lines(fragments_text):
    """Put recognised fragments back into lines of the page.

    Vision returns each piece of text with its box. The boxes are sloppy
    about height - one may take in a descender or a stray mark - but their
    bottoms sit on the baseline, so pieces whose bottoms lie within a
    fraction of a text height of each other belong to one line, ordered
    left to right.
    """
    fragments = []
    for line in fragments_text.splitlines():
        parts = line.split("\t", 4)
        if len(parts) < 5:
            continue
        try:
            x, y, _width, height = (float(part) for part in parts[:4])
        except ValueError:
            continue
        fragments.append((y, x, height, parts[4]))
    if not fragments:
        return []
    heights = sorted(height for _y, _x, height, _text in fragments)
    tolerance = 0.4 * heights[len(heights) // 2]
    fragments.sort(key=lambda item: (-item[0], item[1]))

    lines, current, bottoms = [], [], []
    for bottom, x, _height, text in fragments:
        if bottoms and abs(bottom - sum(bottoms) / len(bottoms)) > tolerance:
            lines.append("  ".join(text for _x, text in sorted(current)))
            current, bottoms = [], []
        current.append((x, text))
        bottoms.append(bottom)
    if current:
        lines.append("  ".join(text for _x, text in sorted(current)))
    return lines


def _run_jxa(script, argument):
    """Run a JavaScript-for-Automation script and return what it printed."""
    handle, script_path = tempfile.mkstemp(suffix=".js", prefix="kisj-import-")
    os.close(handle)
    try:
        with io.open(script_path, "w", encoding="utf-8") as out:
            out.write(script)
        result = subprocess.run(["osascript", "-l", "JavaScript", script_path, argument],
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
    except (OSError, subprocess.TimeoutExpired) as error:
        raise ReadError("The file could not be read:\n%s" % error)
    finally:
        if os.path.exists(script_path):
            os.remove(script_path)
    output = result.stdout.decode("utf-8", "replace")
    if result.returncode != 0:
        raise ReadError("The file could not be read:\n%s"
                        % result.stderr.decode("utf-8", "replace").strip())
    if output.startswith("ERROR:"):
        raise ReadError(output[len("ERROR:"):].strip().capitalize() + ".")
    return output.rstrip("\n")
