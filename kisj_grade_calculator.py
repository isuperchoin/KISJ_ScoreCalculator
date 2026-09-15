"""
KISJ Grade Calculator
=====================

A step-by-step desktop app that turns quarter-by-quarter formative/summative
scores into percentage scores, letter grades, and GPA -- for one class or for
several at once.

    Page 1  Pick your classes (as many as you like) - or import a saved
            report and have everything below filled in
    Page 2  Say how many quarters you have completed
    Page 3  Enter scores, one page per class
    Page 4  All results together

Weighting:  Formative = 20%   |   Summative = 80%
Course list: KISJ High School Course Guide 2026-2027.

Run with a Python that has tkinter:   /usr/bin/python3 kisj_grade_calculator.py
"""

import os
import sys
import webbrowser

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

import grade_export
import grade_import

# ---------------------------------------------------------------------------
# Course catalog (KISJ High School Course Guide 2026-2027)
# ---------------------------------------------------------------------------

COURSES = {
    "English": [
        "English 9",
        "Writing 9",
        "English 10",
        "English 11",
        "English 12",
        "Creative Writing",
        "Film as Literature",
        "Journalism",
        "AP English Language and Composition",
        "AP English Literature and Composition",
        "AP Seminar",
        "AP Research",
    ],
    "Mathematics": [
        "Geometry",
        "Algebra II",
        "Pre-Calculus",
        "Calculus",
        "AP Calculus AB",
        "AP Calculus BC",
        "AP Statistics",
        "Linear Algebra",
        "Multivariable Calculus (Online)",
    ],
    "Science": [
        "Biology",
        "Chemistry",
        "Physics",
        "Earth Science",
        "AP Biology",
        "AP Chemistry",
        "AP Environmental Science",
        "AP Physics 1",
        "AP Physics C (Mechanics and Electricity & Magnetism)",
        "AP Physics C: Mechanics only (Semester 1)",
    ],
    "Social Studies": [
        "Global Studies 9",
        "Global Studies 10",
        "US History",
        "Ethics",
        "Economics",
        "Psychology",
        "Sociology",
        "AP Economics (Micro and Macro)",
        "AP Economics: Macro only",
        "AP Comparative Government and Politics",
        "AP US Government and Politics",
        "AP Human Geography",
        "AP Psychology",
        "AP US History",
        "AP World History: Modern",
    ],
    "Korean Studies": [
        "Korean Language 9",
        "Korean Social Studies 9",
        "Korean Language 10",
        "Korean Social Studies 10",
    ],
    "World Languages": [
        "Chinese I",
        "Chinese II",
        "Chinese III",
        "Chinese IV",
        "AP Chinese Language and Culture",
        "Heritage Chinese",
        "Spanish I",
        "Spanish II",
        "Spanish III",
        "Spanish IV",
        "AP Spanish Language and Culture",
    ],
    "Health and Physical Education": [
        "Health and Physical Education 9",
        "Individual/Dual Activities",
        "Wellness",
        "Movement & Expression",
        "Personal Fitness",
        "Recreational & Lifetime Sports",
    ],
    "Multimedia and Technology": [
        "Design and Technology",
        "Advanced Design and Technology",
        "Digital Photography",
        "Engineering",
        "Advanced Engineering",
        "Graphic Design",
        "Programming I",
        "Programming II",
        "Robotics",
        "Advanced Robotics",
        "Videography",
        "Yearbook",
        "AP Computer Science Principles",
        "AP Computer Science A",
    ],
    "Speech and Debate": [
        "Debate",
        "Public Speaking",
        "Theater I",
    ],
    "Visual and Performing Arts": [
        "Visual Art I",
        "Visual Art II - 2D",
        "Visual Art II - 3D",
        "Choir",
        "Chamber Choir",
        "Solo Vocal Technique",
        "Concert Band",
        "Wind Ensemble",
        "String Orchestra",
        "Advanced String Orchestra",
        "Modern Band",
        "Theater I",
        "Theater II",
        "Advanced Theater",
        "AP Art and Design: 2D",
        "AP Art and Design: 3D",
        "AP Drawing",
        "AP Music Theory",
    ],
}

FORMATIVE_WEIGHT = 0.20
SUMMATIVE_WEIGHT = 0.80

# (minimum rounded percent, letter, gpa points) -- unweighted 4.0 scale
GRADE_SCALE = [
    (97, "A+", 4.0),
    (93, "A", 4.0),
    (90, "A-", 3.7),
    (87, "B+", 3.3),
    (83, "B", 3.0),
    (80, "B-", 2.7),
    (77, "C+", 2.3),
    (73, "C", 2.0),
    (70, "C-", 1.7),
    (67, "D+", 1.3),
    (63, "D", 1.0),
    (60, "D-", 0.7),
    (0, "F", 0.0),
]

QUARTER_CHOICES = [
    "1 quarter (Q1)",
    "2 quarters (Q1-Q2 = Semester 1)",
    "3 quarters (Q1-Q3)",
    "4 quarters (full year)",
]

# Colors
RED = "#c0392b"
GREEN = "#2e7d5b"
INK = "#1f2d3d"
MUTED = "#5a6b7b"
ACCENT = "#1b3a6b"
BG = "#f4f6f9"
CARD = "#ffffff"
BAND = "#eaf0f8"
LINE = "#dde3ea"
IMPORTED_BG = "#eaf6ee"      # a score box filled from an imported report
NOTE_BG = "#e6f2ea"          # green band: scores came from a saved report
WARN_BG = "#fdf3d8"          # amber band: they were only approximated
WARN_FG = "#7a5c12"


# ---------------------------------------------------------------------------
# Grade math (pure functions, easy to test)
# ---------------------------------------------------------------------------

def letter_and_gpa(score):
    """Return (letter, gpa) for a percentage score, rounded to the nearest whole."""
    rounded = int(round(score))
    for minimum, letter, gpa in GRADE_SCALE:
        if rounded >= minimum:
            return letter, gpa
    return "F", 0.0


def average(values):
    return sum(values) / len(values) if values else None


def weighted_score(formatives, summatives):
    """Combine one quarter's scores. Returns (score, note) or (None, note)."""
    f_avg = average(formatives)
    s_avg = average(summatives)

    if f_avg is not None and s_avg is not None:
        return f_avg * FORMATIVE_WEIGHT + s_avg * SUMMATIVE_WEIGHT, ""
    if s_avg is not None:
        return s_avg, "summative only (no formative entered)"
    if f_avg is not None:
        return f_avg, "formative only (no summative entered)"
    return None, "no scores entered"


def validate_score(text):
    """Validate one typed score.

    Returns (value, error_message). value is None when the input is unusable;
    error_message is "" when the input is fine.
    """
    text = text.strip()
    if text == "":
        return None, ""  # blank rows are simply ignored
    try:
        value = float(text)
    except ValueError:
        return None, "Numbers only - please type a whole number from 0 to 100."
    if value != int(value):
        return None, "Must be a whole number (no decimals)."
    value = int(value)
    if value < 0:
        return None, "Cannot be negative - the lowest score is 0."
    if value > 100:
        return None, "Too high - the highest score is 100."
    return value, ""


def summarize(quarter_data):
    """Build one class's report.

    quarter_data maps quarter number -> (formative scores, summative scores).
    Returns a dict with per-quarter rows, semester rows, the overall total, and
    the scores as entered (which the saved report lists so it can be imported).
    """
    quarters = {}
    rows = []
    for number in sorted(quarter_data):
        formatives, summatives = quarter_data[number]
        score, note = weighted_score(formatives, summatives)
        if score is None:
            rows.append({"period": "Quarter %d" % number, "formative": None,
                         "summative": None, "score": None, "note": note})
            continue
        quarters[number] = score
        rows.append({"period": "Quarter %d" % number,
                     "formative": average(formatives),
                     "summative": average(summatives),
                     "score": score, "note": note})

    semesters = {}
    for semester, pair in ((1, (1, 2)), (2, (3, 4))):
        parts = [quarters[q] for q in pair if q in quarters]
        if not parts:
            continue
        semesters[semester] = average(parts)
        note = "Q%d and Q%d" % pair if len(parts) == 2 else \
               "so far - only Q%d counted" % (pair[0] if pair[0] in quarters else pair[1])
        rows.append({"period": "Semester %d" % semester, "formative": None,
                     "summative": None, "score": semesters[semester], "note": note})

    total = average(list(quarters.values()))
    if total is not None:
        label = "Year total" if len(quarters) == 4 else \
                "Total so far (%d quarter%s)" % (len(quarters),
                                                 "" if len(quarters) == 1 else "s")
        rows.append({"period": label, "formative": None, "summative": None,
                     "score": total, "note": "average of all quarters entered"})

    entered = {number: (list(formatives), list(summatives))
               for number, (formatives, summatives) in quarter_data.items()}
    return {"rows": rows, "quarters": quarters, "semesters": semesters, "total": total,
            "entered": entered}


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------

class AccentButton:
    """A filled, colored button.

    tkinter's own Button ignores background colors on macOS, which makes an
    important action look like plain grey chrome. This is a Label styled and
    bound to behave like a button, so the color shows on every platform.
    """

    def __init__(self, parent, text, command, base=ACCENT, hover="#26559b",
                 press="#12294d", fg="white", font=("Helvetica", 12, "bold"),
                 padx=18, pady=8):
        self.command = command
        self.base, self.hover, self.press = base, hover, press
        self.fg = fg
        self.inside = False
        self.enabled = True

        self.frame = tk.Frame(parent, bg=base, cursor="hand2", highlightthickness=1,
                              highlightbackground="#12294d", highlightcolor="#12294d")
        self.label = tk.Label(self.frame, text=text, bg=base, fg=fg, font=font,
                              padx=padx, pady=pady, cursor="hand2")
        self.label.pack()

        for widget in (self.frame, self.label):
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
            widget.bind("<ButtonPress-1>", self._on_press)
            widget.bind("<ButtonRelease-1>", self._on_release)

    def _paint(self, color, fg=None):
        self.frame.config(bg=color)
        self.label.config(bg=color, fg=fg or self.fg)

    def _on_enter(self, _event=None):
        self.inside = True
        if self.enabled:
            self._paint(self.hover)

    def _on_leave(self, _event=None):
        self.inside = False
        if self.enabled:
            self._paint(self.base)

    def _on_press(self, _event=None):
        if self.enabled:
            self._paint(self.press)

    def _on_release(self, _event=None):
        if not self.enabled:
            return
        self._paint(self.hover if self.inside else self.base)
        if self.inside:
            self.command()

    def set_text(self, text):
        self.label.config(text=text)

    def set_enabled(self, enabled):
        self.enabled = enabled
        if enabled:
            self._paint(self.base)
            self.frame.config(cursor="hand2")
            self.label.config(cursor="hand2")
        else:
            self._paint("#cfd6de", fg="#8b97a4")
            self.frame.config(cursor="arrow")
            self.label.config(cursor="arrow")

    def flash(self):
        """Brief highlight, used when the keyboard shortcut fires the action."""
        if not self.enabled:
            return
        self._paint(self.press)
        self.frame.after(140, lambda: self._paint(self.hover if self.inside else self.base))

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)
        return self


class TextButton:
    """A small clickable label.

    Measured on macOS Tk 8.5: every tk.Button created and destroyed leaks about
    90 KB that is never returned (the native button behind it is not released),
    while labels leak nothing. The controls below are created and destroyed
    constantly - one per score box, one per chosen class - so none of them may
    be a tk.Button.
    """

    def __init__(self, parent, text, command, bg=CARD, fg=MUTED, hover_fg=ACCENT,
                 font=None, padx=6, pady=2):
        self.command = command
        self.fg, self.hover_fg = fg, hover_fg
        self.inside = False
        self.label = tk.Label(parent, text=text, bg=bg, fg=fg, cursor="hand2",
                              padx=padx, pady=pady)
        if font:
            self.label.config(font=font)
        self.label.bind("<Enter>", self._on_enter)
        self.label.bind("<Leave>", self._on_leave)
        self.label.bind("<ButtonRelease-1>", self._on_release)

    def _on_enter(self, _event=None):
        self.inside = True
        self.label.config(fg=self.hover_fg)

    def _on_leave(self, _event=None):
        self.inside = False
        self.label.config(fg=self.fg)

    def _on_release(self, _event=None):
        if self.inside:
            self.command()

    def pack(self, **kwargs):
        self.label.pack(**kwargs)
        return self


class SmoothScroll:
    """Pixel-precise, eased vertical scrolling for a canvas.

    A canvas scrolls in "units" - a tenth of the window at a time - and that
    is what every wheel notch, arrow key and trackpad event used to do here,
    so the page moved in visible jumps. Instead:

      * A trackpad (or Magic Mouse) reports how many pixels the fingers moved,
        and the page follows them exactly, every event, the way Tk's own text
        widget does. The system already smooths and adds momentum to that
        stream, so nothing is animated on top of it.
      * A wheel notch or a key press names a distance, and the page glides
        there over a few frames, easing out, rather than landing at once.
        Notches arriving during a glide extend it.

    Scrolling is done with yview_moveto rather than a 1-pixel scroll
    increment, so scrollbar arrows keep their usual step where a platform
    still draws them.
    """

    FRAME_MS = 15         # ~66 frames a second
    EASE = 0.3            # share of the remaining distance covered per frame
    NOTCH = 72            # pixels per wheel notch or arrow key
    MAC_WHEEL = 15        # Tk 8.6 on macOS: pixels per unit of wheel delta

    def __init__(self, canvas):
        self.canvas = canvas
        self.remaining = 0.0
        self.job = None

    # -- geometry ------------------------------------------------------------

    def _extent(self):
        """Height of the scrollable area in pixels, or 0 if nothing scrolls."""
        try:
            region = str(self.canvas.cget("scrollregion")).split()
            height = float(region[3]) - float(region[1])
        except (IndexError, ValueError, tk.TclError):
            return 0.0
        first, last = self.canvas.yview()
        if first <= 0.0 and last >= 1.0:
            return 0.0
        return height

    def _offset(self):
        return self.canvas.yview()[0] * self._extent()

    def page_height(self):
        return max(1, self.canvas.winfo_height() - 40)

    # -- moving --------------------------------------------------------------

    def move(self, pixels):
        """Scroll by this many pixels right now (positive = down)."""
        self.stop()
        self._apply(pixels)

    def _apply(self, pixels):
        extent = self._extent()
        if extent <= 0.0 or not pixels:
            return
        fraction = self.canvas.yview()[0] + pixels / extent
        self.canvas.yview_moveto(max(0.0, min(1.0, fraction)))

    def glide(self, pixels):
        """Scroll by this many pixels smoothly over the next few frames."""
        if self._extent() <= 0.0:
            return
        self.remaining += pixels
        if self.job is None:
            self._step()

    def glide_to(self, offset):
        """Glide so that `offset` pixels of the content sit above the top."""
        self.glide(offset - self._offset() - self.remaining)

    def _step(self):
        self.job = None
        if not self.canvas.winfo_exists():
            return
        step = self.remaining * self.EASE
        if abs(self.remaining) <= 1.5:      # last frame: land exactly
            step = self.remaining
        elif abs(step) < 1.0:               # never crawl below a pixel a frame
            step = 1.0 if self.remaining > 0 else -1.0
        self._apply(step)
        self.remaining -= step
        if self.remaining:
            self.job = self.canvas.after(self.FRAME_MS, self._step)

    def stop(self):
        if self.job is not None:
            try:
                self.canvas.after_cancel(self.job)
            except tk.TclError:
                pass
            self.job = None
        self.remaining = 0.0

    # -- events --------------------------------------------------------------

    def wheel(self, event):
        """<MouseWheel>. Tk 8.6 on macOS sends the trackpad's own small
        deltas at high speed, which are followed directly; everywhere else
        (and on Tk 9 for a real wheel) a notch is a multiple of 120."""
        delta = event.delta
        if delta == 0:
            return
        if sys.platform == "darwin" and float(tk.TkVersion) < 8.7:
            self.move(-delta * self.MAC_WHEEL)
        else:
            notches = delta / 120.0 if abs(delta) >= 120 else (1 if delta > 0 else -1)
            self.glide(-notches * self.NOTCH)

    def buttons(self, direction):
        """X11 reports the wheel as <Button-4> (up) and <Button-5> (down)."""
        self.glide(direction * self.NOTCH)

    def touchpad(self, event):
        """<TouchpadScroll> on Tk 9: the exact distance the fingers moved."""
        try:
            _across, down = self.canvas.tk.call("tk::PreciseScrollDeltas", event.delta)
        except tk.TclError:
            return
        if down:
            self.move(-float(down))

    def bind_local(self, widget, bind_touchpad):
        """Bind the wheel and trackpad on one widget, stopping the events
        there so the window underneath does not scroll as well."""
        widget.bind("<MouseWheel>", lambda e: (self.wheel(e), "break")[1])
        widget.bind("<Shift-MouseWheel>", lambda e: (self.wheel(e), "break")[1])
        widget.bind("<Button-4>", lambda _e: (self.buttons(-1), "break")[1])
        widget.bind("<Button-5>", lambda _e: (self.buttons(1), "break")[1])
        bind_touchpad(widget, lambda e: (self.touchpad(e), "break")[1])


class AssessmentRow:
    """One score entry: a numbered box, an error message, and a remove button."""

    def __init__(self, parent, category, on_remove):
        self.category = category
        self.on_remove = on_remove
        # True while the box still holds a value read from a saved report; the
        # tint tells the user which numbers they have not yet looked at.
        self.imported = False

        self.frame = tk.Frame(parent, bg=CARD)
        self.frame.pack(fill="x", pady=1)

        self.label = tk.Label(self.frame, text="", width=14, anchor="w",
                              bg=CARD, fg=MUTED)
        self.label.pack(side="left")

        self.var = tk.StringVar()
        self.entry = tk.Entry(self.frame, textvariable=self.var, width=7,
                              justify="center", relief="solid", bd=1,
                              highlightthickness=1, highlightbackground="#c9d3de",
                              highlightcolor=ACCENT)
        self.entry.pack(side="left")

        tk.Label(self.frame, text="/100", bg=CARD, fg=MUTED).pack(side="left", padx=(4, 8))

        self.remove_btn = TextButton(self.frame, "✕", self._remove, bg=CARD,
                                     fg=MUTED, hover_fg=RED, padx=4)
        self.remove_btn.pack(side="left")

        self.error = tk.Label(self.frame, text="", bg=CARD, fg=RED, anchor="w",
                              wraplength=340, justify="left")
        self.error.pack(side="left", padx=(8, 0), fill="x", expand=True)

        # Live validation. A variable trace registers a command inside the Tcl
        # interpreter that keeps this row alive even after the widget is gone,
        # so dispose() must remove it - see the dispose chain below.
        self._trace = self.var.trace_add("write", lambda *_: self.check())
        self.entry.bind("<KeyPress>", self._touched)
        self.entry.bind("<KeyRelease>", lambda _e: self.check())
        self.entry.bind("<FocusOut>", lambda _e: self.check())

    def set_value(self, text, imported=False):
        """Fill the box programmatically (from a saved report, or when the
        page is rebuilt); the trace repaints it."""
        self.imported = imported
        self.var.set(text)

    def _touched(self, event):
        """The user typed here: whatever was imported is now theirs."""
        edits = event.char.isprintable() or event.keysym in ("BackSpace", "Delete")
        if self.imported and edits:
            self.imported = False
            self.check()

    def _remove(self):
        self.dispose()
        self.on_remove(self)

    def dispose(self):
        """Drop the Tcl-side trace, then the widgets. Every path that gets rid
        of a row must come through here or the trace outlives the row."""
        if self._trace is not None:
            try:
                self.var.trace_remove("write", self._trace)
            except tk.TclError:
                pass
            self._trace = None
        if self.frame.winfo_exists():
            self.frame.destroy()

    def set_number(self, number):
        self.label.config(text="%s #%d" % (self.category, number))

    def check(self):
        """Show/clear the red message. Returns (value, error)."""
        value, error = validate_score(self.var.get())
        if error:
            self.error.config(text=error)
            self.entry.config(highlightbackground=RED, highlightcolor=RED, bg="#fff5f4")
        else:
            self.error.config(text="")
            self.entry.config(highlightbackground="#c9d3de", highlightcolor=ACCENT,
                              bg=IMPORTED_BG if self.imported else "white")
        return value, error

    def focus(self):
        self.entry.focus_set()


class CategoryBlock:
    """A titled group of assessment rows (Formative or Summative)."""

    def __init__(self, parent, category, weight_text):
        self.category = category
        self.rows = []

        self.frame = tk.LabelFrame(parent, text="  %s  -  %s  " % (category, weight_text),
                                   bg=CARD, fg=ACCENT, bd=1, relief="solid",
                                   font=("Helvetica", 11, "bold"), padx=10, pady=8)
        self.frame.pack(side="left", fill="both", expand=True, padx=6, pady=4)

        self.rows_holder = tk.Frame(self.frame, bg=CARD)
        self.rows_holder.pack(fill="x")

        TextButton(self.frame, "+ Add %s assessment" % category.lower(), self.add_row,
                   bg="#e8eef7", fg=ACCENT, hover_fg="#12294d", padx=8, pady=3
                   ).pack(anchor="w", pady=(6, 0))

        self.add_row(focus=False)

    def add_row(self, focus=True):
        row = AssessmentRow(self.rows_holder, self.category, self._on_remove)
        self.rows.append(row)
        self._renumber()
        if focus:
            row.focus()

    def _on_remove(self, row):
        if row in self.rows:
            self.rows.remove(row)
        if not self.rows:            # always keep at least one box
            self.add_row(focus=False)
        else:
            self._renumber()

    def _renumber(self):
        for index, row in enumerate(self.rows, start=1):
            row.set_number(index)

    def dispose(self):
        for row in self.rows:
            row.dispose()
        self.rows = []

    def set_scores(self, entries):
        """Replace every box with one per entry: (text, imported) pairs."""
        self.dispose()
        for text, imported in entries:
            row = AssessmentRow(self.rows_holder, self.category, self._on_remove)
            self.rows.append(row)
            row.set_value(text, imported)
        if not self.rows:            # always keep at least one box
            self.add_row(focus=False)
        else:
            self._renumber()

    def snapshot(self):
        """(text, imported) for every box that has something in it."""
        return [(row.var.get(), row.imported) for row in self.rows
                if row.var.get().strip()]

    def collect(self):
        """Returns (scores, first_bad_row). first_bad_row is None when all valid."""
        scores = []
        first_bad = None
        for row in self.rows:
            value, error = row.check()
            if error and first_bad is None:
                first_bad = row
            elif value is not None:
                scores.append(value)
        return scores, first_bad


class QuarterBlock:
    """One quarter: a formative column and a summative column."""

    def __init__(self, parent, number):
        self.number = number

        self.frame = tk.LabelFrame(parent, text="  Quarter %d  " % number,
                                   bg=CARD, fg=INK, bd=2, relief="groove",
                                   font=("Helvetica", 13, "bold"), padx=10, pady=8)
        self.frame.pack(fill="x", pady=7, padx=2)

        columns = tk.Frame(self.frame, bg=CARD)
        columns.pack(fill="x")

        self.formative = CategoryBlock(columns, "Formative", "20% of the quarter")
        self.summative = CategoryBlock(columns, "Summative", "80% of the quarter")

    def collect(self):
        formatives, bad_f = self.formative.collect()
        summatives, bad_s = self.summative.collect()
        return formatives, summatives, (bad_f or bad_s)

    def fill(self, formatives, summatives):
        self.formative.set_scores(formatives)
        self.summative.set_scores(summatives)

    def snapshot(self):
        return self.formative.snapshot(), self.summative.snapshot()

    def dispose(self):
        self.formative.dispose()
        self.summative.dispose()


# ---------------------------------------------------------------------------
# Wizard pages
# ---------------------------------------------------------------------------

class Page:
    """Base page: a frame that the wizard shows one at a time."""

    title = ""
    subtitle = ""

    def __init__(self, app, parent):
        self.app = app
        self.frame = tk.Frame(parent, bg=BG)

        header = tk.Frame(self.frame, bg=BG)
        header.pack(fill="x", padx=18, pady=(16, 0))
        tk.Label(header, text=self.title, bg=BG, fg=ACCENT,
                 font=("Helvetica", 17, "bold")).pack(anchor="w")
        if self.subtitle:
            tk.Label(header, text=self.subtitle, bg=BG, fg=MUTED, justify="left",
                     anchor="w", wraplength=880).pack(anchor="w", pady=(3, 0))

        self.body = tk.Frame(self.frame, bg=BG)
        self.body.pack(fill="both", expand=True, padx=18, pady=(10, 18))

    def card(self, padding=14):
        card = tk.Frame(self.body, bg=CARD, bd=1, relief="solid",
                        highlightbackground=LINE)
        card.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="x", padx=padding, pady=padding)
        return inner

    def show(self):
        self.frame.pack(fill="both", expand=True)

    def hide(self):
        self.frame.pack_forget()

    def on_enter(self):
        """Called every time the page becomes visible."""

    def validate(self):
        """Return an error message, or "" when it is fine to move on."""
        return ""


class ClassesPage(Page):
    """Step 1 -- pick one or more classes."""

    HEADER_PREFIX = "─── "
    title = "Step 1  ·  Which classes are you taking?"
    subtitle = ("Pick a department to shorten the list, choose a class, then press "
                "Add class. Repeat for every class you want to calculate - or "
                "import a report you saved earlier and skip the typing.")

    def __init__(self, app, parent):
        Page.__init__(self, app, parent)

        self._build_import_card()

        picker = self.card()

        tk.Label(picker, text="Department", bg=CARD, fg=INK,
                 font=("Helvetica", 11, "bold")).grid(row=0, column=0, sticky="w")
        self.dept_var = tk.StringVar(value="All departments")
        dept_box = ttk.Combobox(picker, textvariable=self.dept_var, state="readonly",
                                width=32, values=["All departments"] + list(COURSES))
        dept_box.grid(row=1, column=0, sticky="w", pady=(2, 0))
        dept_box.bind("<<ComboboxSelected>>", lambda _e: self._refresh_course_list())

        tk.Label(picker, text="Class", bg=CARD, fg=INK,
                 font=("Helvetica", 11, "bold")).grid(row=0, column=1, sticky="w", padx=(20, 0))
        self.course_var = tk.StringVar()
        self.course_box = ttk.Combobox(picker, textvariable=self.course_var,
                                       state="readonly", width=46)
        self.course_box.grid(row=1, column=1, sticky="w", padx=(20, 0), pady=(2, 0))
        self.course_box.bind("<<ComboboxSelected>>", lambda _e: self._reject_header())

        AccentButton(picker, "＋  Add class", self.add_class,
                     font=("Helvetica", 11, "bold"), padx=14, pady=6
                     ).frame.grid(row=1, column=2, sticky="w", padx=(20, 0), pady=(2, 0))

        self.picker_error = tk.Label(picker, text="", bg=CARD, fg=RED, anchor="w")
        self.picker_error.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        chosen = self.card()
        tk.Label(chosen, text="Your classes", bg=CARD, fg=ACCENT,
                 font=("Helvetica", 13, "bold")).pack(anchor="w")
        self.list_holder = tk.Frame(chosen, bg=CARD)
        self.list_holder.pack(fill="x", pady=(8, 0))

        self.empty_label = tk.Label(self.list_holder,
                                    text="No classes yet - add at least one to continue.",
                                    bg=CARD, fg=MUTED)

        self._refresh_course_list()
        self._render_list()

    # -- importing a saved report -------------------------------------------

    def _build_import_card(self):
        card = self.card()
        tk.Label(card, text="Continue from a saved report", bg=CARD, fg=ACCENT,
                 font=("Helvetica", 13, "bold")).pack(anchor="w")
        tk.Label(card, text="Saved your results before? Import that file and every "
                            "class, quarter and score is filled in for you. Then fix "
                            "what was wrong or add what is new.",
                 bg=CARD, fg=MUTED, anchor="w", justify="left",
                 wraplength=860).pack(anchor="w", pady=(2, 8))
        row = tk.Frame(card, bg=CARD)
        row.pack(anchor="w")
        AccentButton(row, "Import previous result...", self.app.import_report,
                     font=("Helvetica", 11, "bold"), padx=14, pady=6).pack(side="left")
        shortcut = "\u2318O" if sys.platform == "darwin" else "Ctrl+O"
        tk.Label(row, text="PDF  \u00b7  JPG  \u00b7  Excel (.xlsx)  \u00b7  CSV / Google "
                           "Sheets       %s" % shortcut,
                 bg=CARD, fg=MUTED).pack(side="left", padx=(14, 0))
        self.import_status = tk.Label(card, text="", bg=CARD, fg=GREEN, anchor="w",
                                      justify="left", wraplength=860)
        self.import_status.pack(anchor="w", pady=(6, 0))

    def set_import_status(self, text):
        self.import_status.config(text=text)

    # -- course dropdown ----------------------------------------------------

    def _refresh_course_list(self):
        dept = self.dept_var.get()
        values = []
        if dept == "All departments":
            for name, courses in COURSES.items():
                values.append("%s%s" % (self.HEADER_PREFIX, name.upper()))
                values.extend("    " + course for course in courses)
        else:
            values = list(COURSES[dept])
        self.course_box.config(values=values)
        first = next((v for v in values if not v.startswith(self.HEADER_PREFIX)), "")
        self.course_var.set(first)

    def _reject_header(self):
        """Department headings in the list are labels, not selectable classes."""
        if self.course_var.get().startswith(self.HEADER_PREFIX):
            values = list(self.course_box.cget("values"))
            index = values.index(self.course_var.get())
            for candidate in values[index + 1:]:
                if not candidate.startswith(self.HEADER_PREFIX):
                    self.course_var.set(candidate)
                    return

    # -- chosen class list --------------------------------------------------

    def add_class(self):
        name = self.course_var.get().strip()
        if not name or name.startswith(self.HEADER_PREFIX.strip()):
            self.picker_error.config(text="Choose a class from the list first.")
            return
        if name in self.app.classes:
            self.picker_error.config(text="%s is already on your list." % name)
            return
        self.picker_error.config(text="")
        self.app.classes.append(name)
        self._render_list()
        self.app.refresh_navigation()

    def remove_class(self, name):
        if name in self.app.classes:
            self.app.classes.remove(name)
        self.picker_error.config(text="")
        self._render_list()
        self.app.refresh_navigation()

    def _render_list(self):
        # destroy, not pack_forget: forgotten frames stay alive and pile up
        # every time a class is added, removed, or the page is revisited.
        for child in self.list_holder.winfo_children():
            if child is not self.empty_label:
                child.destroy()
        self.empty_label.pack_forget()

        if not self.app.classes:
            self.empty_label.pack(anchor="w")
            return

        for index, name in enumerate(self.app.classes, start=1):
            row = tk.Frame(self.list_holder, bg="#f2f6fc", bd=1, relief="solid",
                           highlightbackground=LINE)
            row.pack(fill="x", pady=3)
            tk.Label(row, text="%d." % index, bg="#f2f6fc", fg=MUTED,
                     width=3, anchor="w").pack(side="left", padx=(10, 0), pady=6)
            tk.Label(row, text=name, bg="#f2f6fc", fg=INK, anchor="w",
                     font=("Helvetica", 12)).pack(side="left", pady=6)
            TextButton(row, "Remove", lambda n=name: self.remove_class(n),
                       bg="#f2f6fc", fg=RED, hover_fg="#7d1f14", padx=8
                       ).pack(side="right", padx=10)

    def on_enter(self):
        self._render_list()

    def validate(self):
        if not self.app.classes:
            return "Add at least one class before moving on."
        return ""


class QuartersPage(Page):
    """Step 2 -- how many quarters have been completed."""

    title = "Step 2  ·  How many quarters have you finished?"
    subtitle = ("There are four quarters in a year. Two quarters make one semester "
                "(Q1+Q2 = Semester 1, Q3+Q4 = Semester 2).")

    def __init__(self, app, parent):
        Page.__init__(self, app, parent)

        card = self.card()
        tk.Label(card, text="Quarters completed so far", bg=CARD, fg=INK,
                 font=("Helvetica", 11, "bold")).pack(anchor="w")
        box = ttk.Combobox(card, textvariable=app.quarter_var, state="readonly",
                           width=36, values=QUARTER_CHOICES)
        box.pack(anchor="w", pady=(3, 0))
        box.bind("<<ComboboxSelected>>", lambda _e: self._describe())

        self.detail = tk.Label(card, text="", bg=CARD, fg=MUTED, anchor="w",
                               justify="left")
        self.detail.pack(anchor="w", pady=(10, 0))

        self.class_summary = tk.Label(self.card(), text="", bg=CARD, fg=INK,
                                      anchor="w", justify="left", wraplength=860)
        self.class_summary.pack(anchor="w")

    def _describe(self):
        count = self.app.quarter_count()
        if count == 1:
            text = "You will enter scores for Quarter 1 only. No semester average yet."
        elif count == 2:
            text = "You will enter Q1 and Q2, which together give Semester 1."
        elif count == 3:
            text = "You will enter Q1-Q3. Semester 2 will show Q3 on its own for now."
        else:
            text = "You will enter all four quarters, both semesters, and the year total."
        self.detail.config(text=text)

    def on_enter(self):
        self._describe()
        names = ", ".join(self.app.classes)
        self.class_summary.config(
            text="Classes chosen (%d): %s" % (len(self.app.classes), names))


class ScoresPage(Page):
    """Step 3 -- one page per class."""

    def __init__(self, app, parent, class_name, index, total, quarter_count):
        self.class_name = class_name
        self.title = "Step 3  ·  Scores for %s" % class_name
        self.subtitle = ("Class %d of %d.   Add as many assessments as you took. "
                         "Leave a box blank to skip it. Whole numbers from 0 to 100."
                         % (index, total))
        Page.__init__(self, app, parent)

        # Which class this page is for, and every other class one click away:
        # the name is repeated here, large, so that after an import (or a jump
        # between classes) there is never any doubt whose scores these are.
        self._build_class_strip()

        # Shown only when the scores below came from an imported report.
        self.import_note = None
        self.note_frame = tk.Frame(self.body, bg=NOTE_BG, bd=1, relief="solid",
                                   highlightbackground=LINE)
        self.note_label = tk.Label(self.note_frame, text="", bg=NOTE_BG, fg=GREEN,
                                   anchor="w", justify="left", wraplength=860)
        self.note_label.pack(anchor="w", padx=12, pady=8)

        holder = self.card()
        self.scores_card = holder.master
        self.quarter_blocks = [QuarterBlock(holder, n + 1) for n in range(quarter_count)]

    # -- which class is this --------------------------------------------------

    def _build_class_strip(self):
        strip = tk.Frame(self.body, bg=ACCENT)
        strip.pack(fill="x", pady=(0, 12))
        top = tk.Frame(strip, bg=ACCENT)
        top.pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(top, text="NOW ENTERING", bg=ACCENT, fg="#c8d8f0",
                 font=("Helvetica", 9, "bold")).pack(anchor="w")
        tk.Label(top, text=self.class_name, bg=ACCENT, fg="white", anchor="w",
                 justify="left", wraplength=860,
                 font=("Helvetica", 20, "bold")).pack(anchor="w")
        self.chip_holder = tk.Frame(strip, bg=ACCENT)
        self.chip_holder.pack(fill="x", padx=10, pady=(2, 10))

    CHIP_ROW_WIDTH = 900

    def refresh_class_strip(self):
        """One chip per class: this one lit, the others clickable. Chips flow
        on to further rows rather than run off the right-hand edge."""
        for child in self.chip_holder.winfo_children():
            child.destroy()
        row, used = None, self.CHIP_ROW_WIDTH
        for position, page in enumerate(self.app.score_pages):
            current = page is self
            done = page.has_scores()
            text = "%d  %s" % (position + 1, page.class_name)
            if current:
                text = "\u25b6  " + text
            elif done:
                text = "\u2713  " + text
            font = ("Helvetica", 11, "bold" if current else "normal")
            width = tkfont.Font(font=font).measure(text) + 28
            if row is None or used + width > self.CHIP_ROW_WIDTH:
                row = tk.Frame(self.chip_holder, bg=ACCENT)
                row.pack(fill="x")
                used = 0
            used += width
            chip = tk.Label(row, text=text, padx=10, pady=3, font=font,
                            bg="white" if current else "#2f5187",
                            fg=ACCENT if current else ("#dfe9f7" if done else "#aebdd3"),
                            cursor="arrow" if current else "hand2")
            chip.pack(side="left", padx=4, pady=2)
            if not current:
                chip.bind("<Enter>", lambda _e, c=chip: c.config(bg="#3d64a3", fg="white"))
                chip.bind("<Leave>", lambda _e, c=chip, f=chip.cget("fg"):
                          c.config(bg="#2f5187", fg=f))
                chip.bind("<ButtonRelease-1>",
                          lambda _e, i=position: self.app.jump_to_class(i))

    def on_enter(self):
        self.refresh_class_strip()

    # -- filling in from a saved report ---------------------------------------

    def fill(self, quarter_data, imported=False):
        """quarter_data maps quarter number -> (formatives, summatives), each a
        list of scores or of (text, imported) pairs."""
        for block in self.quarter_blocks:
            formatives, summatives = quarter_data.get(block.number, ([], []))
            block.fill(self._entries(formatives, imported),
                       self._entries(summatives, imported))

    @staticmethod
    def _entries(values, imported):
        entries = []
        for value in values:
            if isinstance(value, tuple):
                entries.append((str(value[0]), bool(value[1])))
            else:
                entries.append((str(value), imported))
        return entries

    def snapshot(self):
        """Everything typed so far, so the page can be rebuilt without loss."""
        return {block.number: block.snapshot() for block in self.quarter_blocks}

    def has_scores(self):
        return any(block.formative.snapshot() or block.summative.snapshot()
                   for block in self.quarter_blocks)

    def set_import_note(self, text, warning=False):
        self.import_note = (text, warning)
        bg, fg = (WARN_BG, WARN_FG) if warning else (NOTE_BG, GREEN)
        self.note_frame.config(bg=bg)
        self.note_label.config(text=text, bg=bg, fg=fg)
        self.note_frame.pack(fill="x", pady=(0, 12), before=self.scores_card)

    def clear_import_note(self):
        self.import_note = None
        self.note_frame.pack_forget()

    def collect(self):
        """Returns (quarter_data, first_bad_quarter_number, first_bad_row)."""
        data = {}
        bad_number = None
        bad_row = None
        for block in self.quarter_blocks:
            formatives, summatives, bad = block.collect()
            data[block.number] = (formatives, summatives)
            if bad is not None and bad_row is None:
                bad_number, bad_row = block.number, bad
        return data, bad_number, bad_row

    def dispose(self):
        for block in self.quarter_blocks:
            block.dispose()
        self.quarter_blocks = []
        self.frame.destroy()

    def validate(self):
        _data, bad_number, bad_row = self.collect()
        if bad_row is not None:
            bad_row.focus()
            return ("Quarter %d of %s has a score that needs fixing - see the red "
                    "message next to it." % (bad_number, self.class_name))
        return ""


EXPORT_LABELS = {
    "pdf": ("PDF", "Save as PDF"),
    "jpg": ("JPG image", "Save as JPG"),
    "xlsx": ("Excel workbook", "Save as .xlsx"),
    "gsheet": ("Google Sheets (CSV)", "Save the CSV"),
}


class PreviewDialog:
    """Shows what a chosen format will actually produce, before saving.

    The window is built once and reused. Measured on macOS Tk 8.5, every
    Toplevel that is created and destroyed leaks about 440 KB that never comes
    back (the native window behind it is not released), so opening a fresh
    preview window each time would grow memory with every look. Widgets inside
    the reused window are rebuilt freely - those cost nothing.
    """

    def __init__(self, app):
        self.app = app
        self.photo = None
        self.temp_image = None
        self.fit_page = True
        self._preview_size = (0, 0)
        self._pages = None
        self.on_confirm = None

        top = self.top = tk.Toplevel(app.root)
        top.withdraw()
        top.configure(bg=BG)
        top.transient(app.root)

        width = self.width = min(900, app.root.winfo_screenwidth() - 120)
        height = min(880, app.root.winfo_screenheight() - 140)
        top.geometry("%dx%d" % (width, height))

        header = tk.Frame(top, bg=ACCENT)
        header.pack(fill="x")
        self.title_label = tk.Label(header, text="", bg=ACCENT, fg="white",
                                    font=("Helvetica", 15, "bold"))
        self.title_label.pack(anchor="w", padx=16, pady=(12, 0))
        self.description = tk.Label(header, text="", bg=ACCENT, fg="#c8d8f0",
                                    anchor="w", justify="left", wraplength=width - 60)
        self.description.pack(anchor="w", padx=16, pady=(2, 12))

        footer = tk.Frame(top, bg=BAND)
        footer.pack(side="bottom", fill="x")
        tk.Frame(top, bg=LINE, height=1).pack(side="bottom", fill="x")
        buttons = tk.Frame(footer, bg=BAND)
        buttons.pack(side="right", padx=16, pady=12)
        AccentButton(buttons, "Cancel", self.cancel, base="#6b7a8c",
                     hover="#7f8d9e", press="#54606f").pack(side="left", padx=(0, 10))
        self.save_button = AccentButton(buttons, "Save...", self.confirm)
        self.save_button.pack(side="left")
        tk.Label(footer, text="This is exactly what will be written to the file.",
                 bg=BAND, fg=MUTED).pack(side="left", padx=16, pady=12)
        # A whole A4 sheet shrunk into the window has type too small to read,
        # so the preview can also be shown at a readable size and scrolled.
        self.zoom_button = TextButton(footer, "", self.toggle_zoom, bg=BAND,
                                      font=("Helvetica", 12, "bold"))

        self.body = tk.Frame(top, bg=CARD)
        self.body.pack(fill="both", expand=True, padx=14, pady=14)

        top.protocol("WM_DELETE_WINDOW", self.cancel)
        top.bind("<Escape>", lambda _e: self.cancel())
        top.bind("<Return>", lambda _e: self.confirm())

    # -- showing ------------------------------------------------------------

    def show(self, kind, document, on_confirm):
        """Display the preview. on_confirm() runs if the user chooses to save.

        This does not block: the answer arrives through the callback rather
        than through a nested Tcl event loop, which keeps the save dialog and
        the preview from fighting over the grab. Each preview still costs
        roughly 70 KB that macOS Tk does not give back; that is fine for
        something opened a handful of times, unlike the per-widget leaks that
        used to grow with every score box in the main window."""
        title, save_label = EXPORT_LABELS[kind]
        self.top.title("Preview - %s" % title)
        self.title_label.config(text="Preview - %s" % title)
        self.description.config(text=grade_export.describe(kind, document))
        self.save_button.set_text("%s..." % save_label)

        for child in self.body.winfo_children():
            child.destroy()
        if kind in ("pdf", "jpg"):
            self._show_zoom_button()
        else:
            self.zoom_button.label.pack_forget()
        if kind == "xlsx":
            self._render(grade_export.xlsx_blocks(document))
        elif kind == "gsheet":
            self._render(grade_export.csv_blocks(document))
        else:
            self._render_pages(kind, document)

        self.on_confirm = on_confirm
        self.top.deiconify()
        self.top.lift()
        self.top.update_idletasks()
        try:
            self.top.grab_set()     # fails only if the parent is not on screen
        except tk.TclError:
            pass
        self.top.focus_set()

    # The PDF's four base fonts, and how to ask Tk for the same thing.
    PREVIEW_FONTS = {
        "Helvetica": ("Helvetica", "normal"),
        "Helvetica-Bold": ("Helvetica", "bold"),
        "Courier": ("Menlo" if sys.platform == "darwin" else "Courier", "normal"),
        "Courier-Bold": ("Menlo" if sys.platform == "darwin" else "Courier", "bold"),
    }

    def _show_zoom_button(self):
        self.zoom_button.label.config(
            text="Show whole page" if not self.fit_page else "Zoom in to read")
        self.zoom_button.label.pack(side="left", padx=(0, 16), pady=12)

    def toggle_zoom(self):
        """Switch between a whole sheet and one big enough to read."""
        self.fit_page = not self.fit_page
        self._show_zoom_button()
        for child in self.body.winfo_children():
            if isinstance(child, tk.Canvas) and self._pages:
                self._preview_size = (0, 0)         # force a redraw at the new size
                self._draw_pages(child, *self._pages)

    def _render_pages(self, kind, document):
        """Draw the report on to-scale pages, paged exactly as the file will be.

        Both formats are A4 portrait - the JPG holds the same sheets as the
        PDF, one under the next - so the preview shows whole A4 sheets with the
        same proportions, margins, page breaks and type sizes, sized so that a
        full sheet is visible at once rather than cropped by the window.
        """
        pages = grade_export.placed_pages(document)
        self._pages = (pages, kind)

        vertical = ttk.Scrollbar(self.body, orient="vertical")
        canvas = tk.Canvas(self.body, bg=BAND, bd=0, highlightthickness=0,
                           yscrollcommand=vertical.set)
        vertical.config(command=canvas.yview)
        vertical.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # The sheets are sized against the canvas, which does not know how big
        # it is until the window is on screen - hence drawing from <Configure>.
        self._preview_size = (0, 0)
        canvas.bind("<Configure>", lambda _e: self._draw_pages(canvas, pages, kind))
        canvas.bind("<Double-Button-1>", lambda _e: self.toggle_zoom())
        self._bind_preview_wheel(canvas)

    # Room left around a sheet: side margins, and the caption under it. A page
    # is never drawn narrower than PAGE_MIN_WIDTH, or the type on a portrait
    # sheet becomes too small to read in a short window; the preview scrolls
    # instead.
    PAGE_SURROUND, PAGE_CAPTION, PAGE_MIN_WIDTH = 56.0, 76.0, 380.0

    def _draw_pages(self, canvas, pages, kind):
        """Draw every sheet at the largest size the window can show whole."""
        width, height = canvas.winfo_width(), canvas.winfo_height()
        if width <= 1 or height <= 1 or (width, height) == self._preview_size:
            return
        self._preview_size = (width, height)
        canvas.delete("all")

        width_points, height_points = pages[0][0], pages[0][1]
        page_width = width - self.PAGE_SURROUND
        if self.fit_page:
            # A whole page, not a cropped band of one: the sheet has to fit the
            # window's height as well as its width, or it does not read as A4.
            fitted = (height - self.PAGE_CAPTION) * width_points / height_points
            page_width = min(page_width, max(fitted, self.PAGE_MIN_WIDTH))
        page_width = max(280.0, page_width)
        scale = page_width / width_points

        # A point of Courier in the PDF and a pixel of Menlo on screen are not
        # the same width, so the type is measured and stepped down until a
        # full-width line really fits between the margins. Without this the
        # widest rows would run off the edge of the drawn page.
        shrink = self._font_shrink(pages, scale)

        left = max(0.0, (width - page_width) / 2.0)
        y = 18.0
        for index, (_width, page_points, lines) in enumerate(pages, start=1):
            page_height = page_points * scale
            canvas.create_rectangle(left + 4, y + 4, left + page_width + 4,
                                    y + page_height + 4, fill="#dde3ea", width=0)
            canvas.create_rectangle(left, y, left + page_width, y + page_height,
                                    fill="white", outline=LINE)
            for x, top, base_font, size, text in lines:
                family, weight = self.PREVIEW_FONTS[base_font]
                pixels = max(5, int(round(size * scale * shrink)))
                canvas.create_text(left + x * scale, y + top * scale, text=text,
                                   anchor="sw", fill=INK, font=(family, -pixels, weight))
            y += page_height + 6
            canvas.create_text(left + page_width / 2.0, y, anchor="n", fill=MUTED,
                               text=self._page_caption(kind, index, len(pages)),
                               font=("Helvetica", 11))
            y += 30

        canvas.configure(scrollregion=(0, 0, width, y))

    def _font_shrink(self, pages, scale):
        """How much to shrink the screen type so the widest line still fits.

        The report is set at whatever size fills the sheet, so the longest line
        of each font is measured as Tk will actually draw it and the type is
        stepped down by however much the screen font runs wider than the PDF's.
        """
        right = (grade_export.PAGE_WIDTH - grade_export.MARGIN) * scale
        longest = {}
        for _width, _height, lines in pages:
            for x, _top, base_font, size, text in lines:
                key = (base_font, max(5, int(round(size * scale))))
                if len(text) > len(longest.get(key, ("", 0.0))[0]):
                    longest[key] = (text, x * scale)

        shrink = 1.0
        for (base_font, pixels), (text, x) in longest.items():
            family, weight = self.PREVIEW_FONTS[base_font]
            drawn = tkfont.Font(family=family, size=-pixels, weight=weight).measure(text)
            room = right - x
            if drawn > room > 0:
                shrink = min(shrink, room / drawn)
        return shrink

    @staticmethod
    def _page_caption(kind, index, total):
        sheet = "Page %d of %d  -  A4 portrait (210 x 297 mm)" % (index, total)
        if kind == "jpg" and total > 1:
            return sheet + ", all in the one image"
        return sheet

    def _bind_preview_wheel(self, canvas):
        """Scroll the preview smoothly, and only the preview: the events stop
        here rather than reaching the main window's handlers as well."""
        self.scroller = SmoothScroll(canvas)
        self.scroller.bind_local(canvas, self.app._bind_touchpad)

    def _render(self, blocks):
        """Draw the preview as text.

        The preview used to be a rendered picture of the page. Tk draws a
        PhotoImage at one image pixel per point, so on a Retina screen every
        such preview was scaled up by the compositor and looked blurry however
        large it was rendered. Text is drawn by the system at the display's
        real resolution, so it stays sharp.
        """
        vertical = ttk.Scrollbar(self.body, orient="vertical")
        horizontal = ttk.Scrollbar(self.body, orient="horizontal")
        text = tk.Text(self.body, wrap="none", bg="white", fg=INK, bd=0,
                       padx=22, pady=18, highlightthickness=0,
                       yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        vertical.config(command=text.yview)
        horizontal.config(command=text.xview)
        vertical.pack(side="right", fill="y")
        horizontal.pack(side="bottom", fill="x")
        text.pack(side="left", fill="both", expand=True)

        mono = "Menlo" if sys.platform == "darwin" else "Courier"
        text.tag_configure("title", font=("Helvetica", 18, "bold"), spacing3=4)
        text.tag_configure("subtitle", font=("Helvetica", 11), foreground=MUTED,
                           spacing3=10)
        text.tag_configure("heading", font=("Helvetica", 14, "bold"),
                           foreground=ACCENT, spacing1=16, spacing3=4)
        text.tag_configure("row", font=(mono, 12))
        text.tag_configure("row_bold", font=(mono, 12, "bold"))
        text.tag_configure("rule", font=(mono, 12), foreground="#9aa5b1")
        text.tag_configure("gap", font=("Helvetica", 5))

        for kind, line in blocks:
            text.insert("end", (line or "") + "\n", kind)
        text.config(state="disabled")

    def _release_image(self):
        if self.temp_image and os.path.exists(self.temp_image):
            os.remove(self.temp_image)
        self.temp_image = None
        self.photo = None

    def _hide(self):
        try:
            self.top.grab_release()
        except tk.TclError:
            pass
        self.top.withdraw()
        self._release_image()

    def confirm(self):
        callback = self.on_confirm
        self.on_confirm = None
        self._hide()
        # Let the preview actually disappear before the save dialog appears.
        if callback:
            self.top.after(1, callback)

    def cancel(self):
        self.on_confirm = None
        self._hide()


class ResultsPage(Page):
    """Step 4 -- every class, side by side, plus the combined picture."""

    title = "Step 4  ·  Your results"
    subtitle = "Everything you entered, per class and all classes together."

    def __init__(self, app, parent):
        Page.__init__(self, app, parent)
        self.headline = tk.Label(self.body, text="", bg=BG, fg=INK, anchor="w",
                                 justify="left", font=("Helvetica", 16, "bold"))
        self.headline.pack(anchor="w", pady=(0, 10))
        self.content = tk.Frame(self.body, bg=BG)
        self.content.pack(fill="both", expand=True)
        # Packed after the content so it sits underneath the numbers: read your
        # results first, then decide whether to save them.
        self.export_bar = tk.Frame(self.body, bg=BG)
        self.export_bar.pack(fill="x", pady=(14, 0))
        self.reports = []

    def _build_export_bar(self):
        for child in self.export_bar.winfo_children():
            child.destroy()

        card = tk.Frame(self.export_bar, bg=CARD, bd=1, relief="solid",
                        highlightbackground=LINE)
        card.pack(fill="x")
        tk.Label(card, text="Save these results", bg=CARD, fg=ACCENT,
                 font=("Helvetica", 13, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(card, text="Happy with the numbers above? Pick a format - you will see "
                            "a preview before anything is saved. Any of these files can "
                            "be imported again later (Step 1) to pick up where you left off.",
                 bg=CARD, fg=MUTED, anchor="w", justify="left",
                 wraplength=880).pack(anchor="w", padx=14)

        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(anchor="w", padx=14, pady=12)
        options = [
            ("PDF", lambda: self.export("pdf")),
            ("JPG image", lambda: self.export("jpg")),
            ("Excel (.xlsx)", lambda: self.export("xlsx")),
            ("Google Sheets", lambda: self.export("gsheet")),
        ]
        for label, command in options:
            AccentButton(buttons, label, command, font=("Helvetica", 11, "bold"),
                         padx=14, pady=6).pack(side="left", padx=(0, 10))

        self.export_note = tk.Label(card, text="", bg=CARD, fg=MUTED, anchor="w",
                                    justify="left", wraplength=880)
        self.export_note.pack(anchor="w", padx=14, pady=(0, 12))
        if not grade_export.jpg_available():
            self.export_note.config(
                text="JPG needs the macOS `sips` tool, which was not found here - "
                     "PDF, Excel and Google Sheets still work.")

    # -- saving -------------------------------------------------------------

    def _document(self):
        return grade_export.build_document(self.reports, letter_and_gpa, average,
                                           self.app.quarter_count())

    def export(self, kind):
        if not any(report["total"] is not None for _name, report in self.reports):
            messagebox.showwarning("Nothing to save",
                                   "Enter at least one score before saving.")
            return

        specs = {
            "pdf": (".pdf", [("PDF document", "*.pdf")], "KISJ grade report.pdf"),
            "jpg": (".jpg", [("JPEG image", "*.jpg")], "KISJ grade report.jpg"),
            "xlsx": (".xlsx", [("Excel workbook", "*.xlsx")], "KISJ grade report.xlsx"),
            "gsheet": (".csv", [("CSV for Google Sheets", "*.csv")],
                       "KISJ grade report.csv"),
        }
        document = self._document()
        self.app.preview_dialog().show(
            kind, document, lambda: self._save(kind, document, specs[kind]))

    def _save(self, kind, document, spec):
        extension, types, initial = spec
        path = filedialog.asksaveasfilename(
            title="Save results", defaultextension=extension,
            filetypes=types + [("All files", "*.*")], initialfile=initial)
        if not path:
            return

        try:
            if kind == "pdf":
                grade_export.write_pdf(path, document)
            elif kind == "jpg":
                grade_export.write_jpg(path, document)
            elif kind == "xlsx":
                grade_export.write_xlsx(path, document)
            else:
                grade_export.write_csv(path, document)
        except grade_export.ExportError as error:
            messagebox.showerror("Could not save", str(error))
            return
        except Exception as error:                     # keep the app alive
            messagebox.showerror("Could not save",
                                 "Something went wrong while saving:\n\n%s" % error)
            return

        if kind == "gsheet":
            opened = messagebox.askyesno(
                "Saved for Google Sheets",
                "Saved to:\n%s\n\nGoogle Sheets cannot be written to directly, so the "
                "results were saved as a CSV file.\n\nOpen a new Google Sheet now? Then "
                "choose File -> Import -> Upload and pick this file." % path)
            if opened:
                webbrowser.open("https://sheets.new")
        else:
            messagebox.showinfo("Saved", "Saved to:\n%s" % path)
        self.export_note.config(text="Last saved: %s" % os.path.basename(path))

    def on_enter(self):
        for child in self.content.winfo_children():
            child.destroy()

        reports = self.app.build_reports()
        self.reports = reports
        self._build_export_bar()
        if not reports:
            self.headline.config(text="No scores were entered yet.")
            self.export_bar.pack_forget()
            tk.Label(self.content, text="Go back and type at least one score.",
                     bg=BG, fg=MUTED).pack(anchor="w")
            return

        self.export_bar.pack(fill="x", pady=(14, 0))
        self._overall_card(reports)
        for name, report in reports:
            self._class_card(name, report)

        tk.Label(self.content,
                 text="Letter grade and GPA use the unweighted 4.0 scale (A+ 97, A 93, "
                      "A- 90, B+ 87, B 83, B- 80, C+ 77, C 73, C- 70, D+ 67, D 63, "
                      "D- 60, F below 60), applied to the score rounded to the nearest "
                      "whole number. The combined GPA is the average of the classes "
                      "listed, treating every class as equal weight.",
                 bg=BG, fg=MUTED, anchor="w", justify="left",
                 wraplength=900).pack(anchor="w", pady=(4, 0))

    # -- combined table -----------------------------------------------------

    def _overall_card(self, reports):
        card = tk.Frame(self.content, bg=CARD, bd=1, relief="solid",
                        highlightbackground=LINE)
        card.pack(fill="x", pady=(0, 14))
        tk.Label(card, text="All classes together", bg=CARD, fg=ACCENT,
                 font=("Helvetica", 14, "bold")).pack(anchor="w", padx=14, pady=(12, 6))

        columns = ("class", "q1", "q2", "q3", "q4", "s1", "s2", "total", "letter", "gpa")
        headings = {
            "class": ("Class", 250), "q1": ("Q1", 72), "q2": ("Q2", 72),
            "q3": ("Q3", 72), "q4": ("Q4", 72), "s1": ("Sem 1", 82),
            "s2": ("Sem 2", 82), "total": ("Total", 90),
            "letter": ("Letter", 66), "gpa": ("GPA", 60),
        }
        tree = ttk.Treeview(card, columns=columns, show="headings",
                            height=len(reports) + 1)
        for key, (text, width) in headings.items():
            tree.heading(key, text=text)
            tree.column(key, width=width, anchor="w" if key == "class" else "center")
        tree.tag_configure("summary", font=("Helvetica", 11, "bold"))
        tree.pack(fill="x", padx=14, pady=(0, 12))
        self.app.forward_wheel(tree)

        totals, gpas = [], []
        for name, report in reports:
            quarters = report["quarters"]
            semesters = report["semesters"]
            letter, gpa = ("-", None)
            if report["total"] is not None:
                letter, gpa = letter_and_gpa(report["total"])
                totals.append(report["total"])
                gpas.append(gpa)
            tree.insert("", "end", values=(
                name,
                pct(quarters.get(1)), pct(quarters.get(2)),
                pct(quarters.get(3)), pct(quarters.get(4)),
                pct(semesters.get(1)), pct(semesters.get(2)),
                pct(report["total"]), letter,
                "-" if gpa is None else "%.1f" % gpa,
            ))

        if totals:
            combined = average(totals)
            combined_gpa = average(gpas)
            tree.insert("", "end", tags=("summary",), values=(
                "ALL CLASSES (%d)" % len(totals), "", "", "", "", "", "",
                pct(combined), letter_and_gpa(combined)[0], "%.2f" % combined_gpa,
            ))
            self.headline.config(
                text="Overall GPA %.2f   ·   %.2f%% average across %d class%s"
                     % (combined_gpa, combined, len(totals),
                        "" if len(totals) == 1 else "es"))
        else:
            self.headline.config(text="No scores were entered yet.")

    # -- per class detail ---------------------------------------------------

    def _class_card(self, name, report):
        card = tk.Frame(self.content, bg=CARD, bd=1, relief="solid",
                        highlightbackground=LINE)
        card.pack(fill="x", pady=(0, 12))

        header = tk.Frame(card, bg=CARD)
        header.pack(fill="x", padx=14, pady=(12, 6))
        tk.Label(header, text=name, bg=CARD, fg=ACCENT,
                 font=("Helvetica", 13, "bold")).pack(side="left")
        if report["total"] is not None:
            letter, gpa = letter_and_gpa(report["total"])
            tk.Label(header, text="%.2f%%   %s   GPA %.1f" % (report["total"], letter, gpa),
                     bg=CARD, fg=INK, font=("Helvetica", 12, "bold")).pack(side="right")
        else:
            tk.Label(header, text="no scores entered", bg=CARD, fg=RED).pack(side="right")

        columns = ("period", "formative", "summative", "score", "letter", "gpa", "note")
        headings = {
            "period": ("Period", 200), "formative": ("Formative avg", 120),
            "summative": ("Summative avg", 120), "score": ("Weighted score", 130),
            "letter": ("Letter", 70), "gpa": ("GPA", 60), "note": ("Note", 260),
        }
        tree = ttk.Treeview(card, columns=columns, show="headings",
                            height=max(1, len(report["rows"])))
        for key, (text, width) in headings.items():
            tree.heading(key, text=text)
            tree.column(key, width=width,
                        anchor="w" if key in ("period", "note") else "center")
        tree.tag_configure("summary", font=("Helvetica", 11, "bold"))
        tree.pack(fill="x", padx=14, pady=(0, 12))
        self.app.forward_wheel(tree)

        for row in report["rows"]:
            score = row["score"]
            letter, gpa = ("-", "-")
            if score is not None:
                letter, gpa_value = letter_and_gpa(score)
                gpa = "%.1f" % gpa_value
            is_summary = not row["period"].startswith("Quarter")
            tree.insert("", "end", tags=("summary",) if is_summary else (), values=(
                row["period"], num(row["formative"]), num(row["summative"]),
                pct(score), letter, gpa, row["note"],
            ))


def pct(value):
    return "-" if value is None else "%.2f%%" % value


def num(value):
    return "-" if value is None else "%.2f" % value


# ---------------------------------------------------------------------------
# The wizard
# ---------------------------------------------------------------------------

class GradeCalculatorApp:

    STEP_LABELS = ["1  Classes", "2  Quarters", "3  Scores", "4  Results"]

    def __init__(self, root):
        self.root = root
        root.title("KISJ Score Calculator")
        self._apply_window_icon(root)
        # Never open taller or wider than the screen, or the footer buttons end
        # up off-screen where no amount of scrolling can reach them.
        width = min(1040, root.winfo_screenwidth() - 80)
        height = min(780, root.winfo_screenheight() - 120)
        root.geometry("%dx%d" % (max(width, 760), max(height, 520)))
        root.minsize(760, 520)
        root.configure(bg=BG)

        # Guards for the scroll plumbing (see _build_scroll_area).
        self._canvas_width = None
        self._scroll_region = None
        self._scroll_pending = False

        self.classes = []
        self.quarter_var = tk.StringVar(value=QUARTER_CHOICES[1])
        self.score_pages = []
        self.score_signature = None
        self.index = 0
        self._preview = None
        self.import_source = None

        self._build_header()
        self._build_footer()
        self._build_scroll_area()

        self.classes_page = ClassesPage(self, self.page_holder)
        self.quarters_page = QuartersPage(self, self.page_holder)
        self.results_page = ResultsPage(self, self.page_holder)

        self.show_page(0)
        root.bind("<Return>", self._on_return_key)
        root.bind("<KP_Enter>", self._on_return_key)
        root.bind("<Command-o>", lambda _e: self.import_report())
        root.bind("<Control-o>", lambda _e: self.import_report())

    # -- chrome -------------------------------------------------------------

    @staticmethod
    def _apply_window_icon(root):
        """Set the window icon where the toolkit supports it.

        On Windows and Linux this puts the KIScore mark on the window and in
        the taskbar. macOS ignores it: there the Dock icon and the menu-bar
        name come from the application bundle, which build_macos_app.py
        creates - run the calculator from "KISJ score calculator.app" to get
        the icon instead of the Python rocket.
        """
        # sys._MEIPASS is where PyInstaller unpacks bundled data files; when
        # the program runs as a plain script it is not set and the icon sits
        # next to the source.
        base = getattr(sys, "_MEIPASS",
                       os.path.dirname(os.path.abspath(__file__)))
        icon = os.path.join(base, "kiscore_icon.gif")
        if not os.path.exists(icon):
            return
        try:
            root._window_icon = tk.PhotoImage(file=icon)   # kept alive on purpose
            root.iconphoto(True, root._window_icon)
        except tk.TclError:
            pass

    def _build_header(self):
        bar = tk.Frame(self.root, bg=ACCENT)
        bar.pack(fill="x")
        tk.Label(bar, text="KISJ Score Calculator", bg=ACCENT, fg="white",
                 font=("Helvetica", 18, "bold")).pack(anchor="w", padx=18, pady=(12, 0))
        tk.Label(bar, text="Formative 20%  ·  Summative 80%  ·  quarter, semester and "
                           "year results for as many classes as you like",
                 bg=ACCENT, fg="#c8d8f0").pack(anchor="w", padx=18, pady=(2, 10))

        strip = tk.Frame(self.root, bg=BAND)
        strip.pack(fill="x")
        inner = tk.Frame(strip, bg=BAND)
        inner.pack(anchor="w", padx=18, pady=8)
        self.step_chips = []
        for position, text in enumerate(self.STEP_LABELS):
            if position:
                tk.Label(inner, text="  ▸  ", bg=BAND, fg="#9aa8b8").pack(side="left")
            chip = tk.Label(inner, text=text, bg="#dfe5ec", fg=MUTED, padx=12, pady=4,
                            font=("Helvetica", 11, "bold"))
            chip.pack(side="left")
            self.step_chips.append(chip)
        # Apple's system Python still ships Tk 8.5, which predates macOS
        # precise trackpad scrolling: two-finger scrolling sends this program
        # nothing at all. Say so rather than let it look broken.
        if float(tk.TkVersion) < 8.6:
            warning = tk.Frame(self.root, bg="#fdf3d8")
            warning.pack(fill="x")
            tk.Label(warning,
                     text="Two-finger scrolling does not work on this old Tk "
                          "(%s). Use the scrollbar or the arrow keys - or open "
                          "\u201cKISJ Score Calculator.app\u201d, which uses a newer one."
                          % tk.TkVersion,
                     bg="#fdf3d8", fg="#7a5c12", anchor="w", justify="left",
                     wraplength=980).pack(anchor="w", padx=18, pady=6)

        tk.Frame(self.root, bg=LINE, height=1).pack(fill="x")

    def _build_footer(self):
        tk.Frame(self.root, bg=LINE, height=1).pack(side="bottom", fill="x")
        footer = tk.Frame(self.root, bg=BAND)
        footer.pack(side="bottom", fill="x")

        buttons = tk.Frame(footer, bg=BAND)
        buttons.pack(side="right", padx=18, pady=12)
        self.back_button = AccentButton(buttons, "◀  Back", self.go_back,
                                        base="#6b7a8c", hover="#7f8d9e", press="#54606f")
        self.back_button.pack(side="left", padx=(0, 10))
        self.next_button = AccentButton(buttons, "Next  ▶", self.go_next)
        self.next_button.pack(side="left")

        self.status = tk.Label(footer, text="", bg=BAND, fg=RED, anchor="w",
                               justify="left", wraplength=560)
        self.status.pack(side="left", padx=18, pady=12, fill="x", expand=True)

    def _build_scroll_area(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.page_holder = tk.Frame(self.canvas, bg=BG)
        self.holder_id = self.canvas.create_window((0, 0), window=self.page_holder,
                                                   anchor="nw")

        # These two handlers used to call each other forever: resizing the canvas
        # set the inner frame's width, the frame's <Configure> reset the
        # scrollregion, which resized the canvas again. Both are now guarded so
        # they only act on a real change, and the scrollregion update is
        # coalesced into a single idle callback.
        self.page_holder.bind("<Configure>", self._on_holder_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Wheel/trackpad. These are bound once, for the whole application, and
        # never unbound: the canvas is completely covered by the page content,
        # and Tk sends the canvas a <Leave> the moment the pointer moves onto
        # any child widget - scoping the bindings to Enter/Leave tore them down
        # immediately and two-finger scrolling did nothing at all.
        scroller = self.scroller = SmoothScroll(self.canvas)
        self._bind_wheel()

        # Keyboard scrolling always works, whatever the trackpad does.
        self.root.bind("<Up>", lambda _e: scroller.glide(-scroller.NOTCH))
        self.root.bind("<Down>", lambda _e: scroller.glide(scroller.NOTCH))
        self.root.bind("<Prior>", lambda _e: scroller.glide(-scroller.page_height()))
        self.root.bind("<Next>", lambda _e: scroller.glide(scroller.page_height()))
        self.root.bind("<Home>", lambda _e: scroller.glide_to(0))
        self.root.bind("<End>", lambda _e: scroller.glide_to(scroller._extent()))

    # -- scrolling ----------------------------------------------------------

    def _on_canvas_configure(self, event):
        if event.width != self._canvas_width:          # guard breaks the loop
            self._canvas_width = event.width
            self.canvas.itemconfig(self.holder_id, width=event.width)
        self.schedule_scrollregion()

    def _on_holder_configure(self, _event=None):
        self.schedule_scrollregion()

    def schedule_scrollregion(self):
        """Recompute the scrollregion once, after the layout settles."""
        if self._scroll_pending:
            return
        self._scroll_pending = True
        self.canvas.after_idle(self._apply_scrollregion)

    def _apply_scrollregion(self):
        self._scroll_pending = False
        if not self.canvas.winfo_exists():
            return
        bbox = self.canvas.bbox("all")
        if bbox is None:
            return
        region = (0, 0, bbox[2], bbox[3])
        if region != self._scroll_region:              # guard breaks the loop
            self._scroll_region = region
            self.canvas.configure(scrollregion=region)

    def _bind_wheel(self, _event=None):
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Shift-MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>", lambda e: self._in_main(e) and self.scroller.buttons(-1))
        self.root.bind_all("<Button-5>", lambda e: self._in_main(e) and self.scroller.buttons(1))
        self._bind_touchpad(self.root, self._on_touchpad, bind_all=True)

    def _in_main(self, event):
        """bind_all also hears the preview window; only the main page scrolls
        for events that happen over the main window."""
        widget = event.widget
        try:
            return widget.winfo_toplevel() is self.root
        except (AttributeError, tk.TclError):
            return False

    @staticmethod
    def _bind_touchpad(widget, handler, bind_all=False):
        """Bind the precise-scrolling event, where the toolkit has one.

        Tk 9 stopped reporting precise scrolling devices - every Mac trackpad,
        and the Magic Mouse - as <MouseWheel>, and gave them <TouchpadScroll>
        instead. Miss it and two-finger scrolling reaches nothing at all, which
        is exactly how the bundled app behaved. Tk 8.6 has no such event and
        raises rather than binding it; there the wheel bindings above are
        already what a trackpad sends, so there is nothing to fall back to.
        """
        try:
            bind = widget.bind_all if bind_all else widget.bind
            bind("<TouchpadScroll>", handler)
        except tk.TclError:
            pass

    def _on_touchpad(self, event):
        if self._in_main(event):
            self.scroller.touchpad(event)

    def _on_mousewheel(self, event):
        if self._in_main(event):
            self.scroller.wheel(event)

    def forward_wheel(self, widget):
        """Let the page scroll even when the pointer sits on a Treeview,
        which otherwise swallows the wheel event."""
        self.scroller.bind_local(widget, self._bind_touchpad)

    # -- page bookkeeping ---------------------------------------------------

    def quarter_count(self):
        return int(self.quarter_var.get()[0])

    def pages(self):
        return [self.classes_page, self.quarters_page] + self.score_pages + [self.results_page]

    def rebuild_score_pages(self):
        """One score page per class.

        Nothing typed is lost when the setup changes: a class that stays on
        the list gets its scores back, so an imported report can have a class
        added to it - or a quarter - without retyping the rest.
        """
        signature = (tuple(self.classes), self.quarter_count())
        if signature == self.score_signature:
            return
        kept = {page.class_name: (page.snapshot(), page.import_note)
                for page in self.score_pages}
        for page in self.score_pages:
            page.dispose()
        total = len(self.classes)
        self.score_pages = [
            ScoresPage(self, self.page_holder, name, position, total, self.quarter_count())
            for position, name in enumerate(self.classes, start=1)
        ]
        for page in self.score_pages:
            if page.class_name in kept:
                data, note = kept[page.class_name]
                page.fill(data)
                if note:
                    page.set_import_note(*note)
        self.score_signature = signature

    def show_page(self, index):
        pages = self.pages()
        index = max(0, min(index, len(pages) - 1))
        for page in pages:
            page.hide()
        self.index = index
        page = pages[index]
        page.on_enter()
        page.show()
        self.scroller.stop()
        self.canvas.yview_moveto(0)
        self.schedule_scrollregion()
        self.status.config(text="")
        self.refresh_navigation()

    def refresh_navigation(self):
        pages = self.pages()
        page = pages[self.index]

        # step chips
        if isinstance(page, ClassesPage):
            active = 0
        elif isinstance(page, QuartersPage):
            active = 1
        elif isinstance(page, ScoresPage):
            active = 2
        else:
            active = 3
        for position, chip in enumerate(self.step_chips):
            if position == active:
                chip.config(bg=ACCENT, fg="white")
            elif position < active:
                chip.config(bg=GREEN, fg="white")
            else:
                chip.config(bg="#dfe5ec", fg=MUTED)
        if isinstance(page, ScoresPage) and len(self.score_pages) > 1:
            position = self.score_pages.index(page) + 1
            self.step_chips[2].config(text="3  Scores (%d/%d)"
                                           % (position, len(self.score_pages)))
        else:
            self.step_chips[2].config(text=self.STEP_LABELS[2])

        # buttons
        self.back_button.set_enabled(self.index > 0)
        if isinstance(page, ResultsPage):
            self.next_button.set_text("↺  Start over")
        elif isinstance(page, ScoresPage) and page is self.score_pages[-1]:
            self.next_button.set_text("See my results  ▶")
        elif isinstance(page, QuartersPage):
            self.next_button.set_text("Enter scores  ▶")
        else:
            self.next_button.set_text("Next  ▶")
        self.next_button.set_enabled(
            not (isinstance(page, ClassesPage) and not self.classes))

    # -- navigation ---------------------------------------------------------

    def go_next(self):
        pages = self.pages()
        page = pages[self.index]

        if isinstance(page, ResultsPage):
            self.start_over()
            return

        error = page.validate()
        if error:
            self.status.config(text=error)
            return
        self.status.config(text="")

        if isinstance(page, QuartersPage):
            self.rebuild_score_pages()

        if isinstance(page, ScoresPage) and page is self.score_pages[-1]:
            # The class chips let you skip ahead, so check every class before
            # the results, not just the one on screen.
            for other in self.score_pages:
                if other.validate():
                    self.show_page(pages.index(other))
                    self.status.config(text=other.validate())
                    return

        self.show_page(self.index + 1)

    def go_back(self):
        if self.index > 0:
            self.show_page(self.index - 1)

    def jump_to_class(self, position):
        """Switch straight to another class's score page (from the chips)."""
        page = self.pages()[self.index]
        if isinstance(page, ScoresPage):
            error = page.validate()
            if error:
                self.status.config(text=error)
                return
        if 0 <= position < len(self.score_pages):
            self.show_page(2 + position)

    def _on_return_key(self, _event=None):
        self.next_button.flash()
        self.go_next()

    def start_over(self):
        for page in self.score_pages:
            page.dispose()
        self.score_pages = []
        self.score_signature = None
        self.classes = []
        self.quarter_var.set(QUARTER_CHOICES[1])
        self.import_source = None
        self.classes_page.set_import_status("")
        self.show_page(0)

    # -- importing a saved report -------------------------------------------

    @staticmethod
    def known_courses():
        names = []
        for courses in COURSES.values():
            names.extend(course for course in courses if course not in names)
        return names

    def import_report(self):
        """Read a report saved earlier and fill the whole calculator from it."""
        path = filedialog.askopenfilename(
            title="Import a saved report",
            filetypes=[("Saved reports", "*.pdf *.jpg *.jpeg *.png *.xlsx *.csv"),
                       ("PDF document", "*.pdf"), ("Picture", "*.jpg *.jpeg *.png"),
                       ("Excel workbook", "*.xlsx"), ("CSV (Google Sheets)", "*.csv"),
                       ("All files", "*.*")])
        if not path:
            return

        self.root.config(cursor="watch")
        self.status.config(text="")
        self.root.update_idletasks()
        try:
            result = grade_import.read_report(path, self.known_courses())
        except grade_import.ReadError as error:
            messagebox.showerror("Could not import", str(error))
            return
        except Exception as error:                     # keep the app alive
            messagebox.showerror("Could not import",
                                 "Something went wrong while reading the file:\n\n%s"
                                 % error)
            return
        finally:
            self.root.config(cursor="")

        text = grade_import.summary_text(result)
        if any(page.has_scores() for page in self.score_pages):
            text += ("\n\nThis replaces the classes and scores you have entered "
                     "so far.")
        text += "\n\nFill in the calculator with these? You can fix or add to them next."
        if not messagebox.askokcancel("Import previous result", text):
            return
        self.apply_import(result)

    def apply_import(self, result):
        for page in self.score_pages:
            page.dispose()
        self.score_pages = []
        self.score_signature = None
        self.classes = list(result["classes"])
        self.quarter_var.set(QUARTER_CHOICES[result["quarter_count"] - 1])
        self.rebuild_score_pages()

        source = result["source"]
        if result["approximate"]:
            note = ("Imported from %s, which only listed quarter averages - each "
                    "average was entered as one rounded score. Retype the real "
                    "scores if you have them, then add anything new." % source)
        else:
            note = ("Imported from %s - the tinted boxes are the scores you saved. "
                    "Check them, fix anything that changed, or add new assessments."
                    % source)
        if result["recognized"]:
            note += " The numbers were read from a picture, so look over every one."
        for page in self.score_pages:
            page.fill(result["scores"].get(page.class_name, {}), imported=True)
            page.set_import_note(note, warning=result["approximate"] or
                                 result["recognized"])

        self.import_source = source
        self.classes_page.set_import_status(
            "Imported %d class%s and %d quarter%s from %s." % (
                len(self.classes), "" if len(self.classes) == 1 else "es",
                result["quarter_count"], "" if result["quarter_count"] == 1 else "s",
                source))
        self.show_page(2 if self.score_pages else 0)

    def preview_dialog(self):
        """One reused preview window for the whole session."""
        if self._preview is None:
            self._preview = PreviewDialog(self)
        return self._preview

    # -- results ------------------------------------------------------------

    def build_reports(self):
        """[(class name, report dict)] for every class that has a score page."""
        reports = []
        for page in self.score_pages:
            data, _bad_number, _bad_row = page.collect()
            reports.append((page.class_name, summarize(data)))
        return reports


def main():
    root = tk.Tk()
    GradeCalculatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
