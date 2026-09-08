"""
Build the KIScore icon and a macOS application bundle.

Why this exists: on macOS the Dock icon and the menu-bar name come from the
application bundle that launches the process, not from Tkinter. Running
`python3 kisj_grade_calculator.py` therefore shows the Python launcher's rocket
and the name "Python", no matter what the program does. Wrapping the same
script in a tiny .app fixes both.

    /usr/bin/python3 build_macos_app.py

Produces, next to this file:
    KISJ score calculator.app   double-click to run, correct icon and name
    kiscore_icon.gif            window icon used on Windows/Linux
"""

import io
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

# Interpreters to consider, best first. Tk 8.5 (Apple's system Python) predates
# macOS precise trackpad scrolling and is not Retina aware, so two-finger
# scrolling does nothing and bitmaps look soft; any Tk 8.6+ build fixes both.
CANDIDATE_PYTHONS = [
    "/Library/Frameworks/Python.framework/Versions/*/bin/python3",
    "/opt/homebrew/bin/python3",
    "/usr/local/bin/python3",
    "/usr/bin/python3",
]

PROBE = ("import tkinter, sys;"
         "print(tkinter.TkVersion, sys.base_prefix)")


def find_interpreter():
    """Pick the installed interpreter with the newest Tk."""
    import glob
    best = None
    seen = set()
    for pattern in CANDIDATE_PYTHONS:
        for path in sorted(glob.glob(pattern)) or [pattern]:
            real = os.path.realpath(path)
            if real in seen or not os.path.exists(path):
                continue
            seen.add(real)
            probe = subprocess.run([path, "-c", PROBE],
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if probe.returncode != 0:
                continue
            parts = probe.stdout.decode("utf-8", "replace").split()
            try:
                version = float(parts[0])
            except (IndexError, ValueError):
                continue
            base = " ".join(parts[1:])
            if best is None or version > best[1]:
                best = (path, version, base)
    return best
APP_NAME = "KISJ Score Calculator"        # Dock label, Finder name, menu bar
ICON_TEXT = "KIScore"
SCRIPT = os.path.join(HERE, "kisj_grade_calculator.py")

# KIS signature navy, as RGB fractions
BLUE = (0x1B / 255.0, 0x3A / 255.0, 0x6B / 255.0)

# Helvetica-Bold advance widths (per 1000 units) for the letters we draw
WIDTHS = {"K": 722, "I": 278, "S": 667, "c": 556, "o": 611, "r": 389, "e": 556}


def icon_pdf(path, size=1024):
    """A rounded blue square with the wordmark centred on it."""
    radius = size * 0.22
    k = radius * 0.5523                      # circle-ish bezier handle
    x0 = y0 = 0.0
    x1 = y1 = float(size)

    text_width_units = sum(WIDTHS[ch] for ch in ICON_TEXT)
    font_size = (size * 0.78) / (text_width_units / 1000.0)
    text_x = (size - text_width_units / 1000.0 * font_size) / 2.0
    text_y = (size - 0.717 * font_size) / 2.0    # 0.717 = Helvetica cap height

    content = (
        "%.4f %.4f %.4f rg\n"
        "%.2f %.2f m\n"
        "%.2f %.2f l\n"
        "%.2f %.2f %.2f %.2f %.2f %.2f c\n"
        "%.2f %.2f l\n"
        "%.2f %.2f %.2f %.2f %.2f %.2f c\n"
        "%.2f %.2f l\n"
        "%.2f %.2f %.2f %.2f %.2f %.2f c\n"
        "%.2f %.2f l\n"
        "%.2f %.2f %.2f %.2f %.2f %.2f c\n"
        "f\n"
        "1 1 1 rg\n"
        "BT /F1 %.2f Tf %.2f %.2f Td (%s) Tj ET\n"
        % (BLUE[0], BLUE[1], BLUE[2],
           x0 + radius, y0,
           x1 - radius, y0,
           x1 - radius + k, y0, x1, y0 + radius - k, x1, y0 + radius,
           x1, y1 - radius,
           x1, y1 - radius + k, x1 - radius + k, y1, x1 - radius, y1,
           x0 + radius, y1,
           x0 + radius - k, y1, x0, y1 - radius + k, x0, y1 - radius,
           x0, y0 + radius,
           x0, y0 + radius - k, x0 + radius - k, y0, x0 + radius, y0,
           font_size, text_x, text_y, ICON_TEXT))

    objects = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Resources "
        "<< /Font << /F1 5 0 R >> >> /Contents 4 0 R >>" % (size, size),
        "<< /Length %d >>\nstream\n%s\nendstream" % (len(content) + 1, content),
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold "
        "/Encoding /WinAnsiEncoding >>",
    ]

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
    with io.open(path, "wb") as handle:
        handle.write(bytes(out))
    return path


def sips(*arguments):
    result = subprocess.run(["sips"] + list(arguments),
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        raise SystemExit("sips failed: %s" % result.stderr.decode("utf-8", "replace"))


def build_icns(work, icns_path):
    pdf = icon_pdf(os.path.join(work, "icon.pdf"))
    master = os.path.join(work, "icon_1024.png")
    sips("-s", "format", "png", "--resampleHeightWidth", "1024", "1024", pdf, "--out", master)

    iconset = os.path.join(work, "KIScore.iconset")
    os.makedirs(iconset, exist_ok=True)
    for base in (16, 32, 128, 256, 512):
        for scale, suffix in ((1, ""), (2, "@2x")):
            pixels = base * scale
            target = os.path.join(iconset, "icon_%dx%d%s.png" % (base, base, suffix))
            sips("-s", "format", "png", "--resampleHeightWidth", str(pixels), str(pixels),
                 master, "--out", target)

    result = subprocess.run(["iconutil", "-c", "icns", iconset, "-o", icns_path],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        raise SystemExit("iconutil failed: %s" % result.stderr.decode("utf-8", "replace"))
    return master


INFO_PLIST = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
 "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key><string>%(name)s</string>
    <key>CFBundleDisplayName</key><string>%(name)s</string>
    <key>CFBundleExecutable</key><string>launcher</string>
    <key>CFBundleIconFile</key><string>KIScore</string>
    <key>CFBundleIdentifier</key><string>com.kisj.kiscore</string>
    <key>CFBundleInfoDictionaryVersion</key><string>6.0</string>
    <key>CFBundlePackageType</key><string>APPL</string>
    <key>CFBundleShortVersionString</key><string>1.0</string>
    <key>CFBundleVersion</key><string>1</string>
    <key>NSHighResolutionCapable</key><true/>
    <key>LSMinimumSystemVersion</key><string>10.13</string>
</dict>
</plist>
"""

LAUNCHER = """#!/bin/bash
# Runs the calculator with the interpreter copy that lives inside this bundle.
# It has to be the in-bundle copy: macOS gives the Dock icon and the menu-bar
# name of whichever bundle owns the executable that ends up running, so
# exec-ing /usr/bin/python3 here would hand the app back to Python.app and the
# rocket would return.
here="$(cd "$(dirname "$0")" && pwd)"
exec "$here/python" "%s"
"""

FALLBACK_LAUNCHER = """#!/bin/bash
# Fallback: no in-bundle interpreter, so macOS will show the Python icon.
exec "%s" "%s"
"""


def embed_interpreter(macos_dir, framework):
    """Put a working copy of the GUI interpreter inside the bundle.

    Returns True when it worked. The copy is linked against the framework by a
    relative path that only resolves inside Python.app, so the load command is
    repointed at the absolute framework and the binary is re-signed ad-hoc.
    """
    source = os.path.join(framework, "Resources", "Python.app", "Contents",
                          "MacOS", "Python")
    dylib = None
    for name in ("Python3", "Python"):            # 3.9 ships Python3, 3.14 ships Python
        candidate = os.path.join(framework, name)
        if os.path.exists(candidate):
            dylib = candidate
            break
    if not (os.path.exists(source) and dylib):
        return False

    target = os.path.join(macos_dir, "python")
    shutil.copy2(source, target)
    os.chmod(target, 0o755)

    old_path = subprocess.run(["otool", "-L", target], stdout=subprocess.PIPE,
                              stderr=subprocess.PIPE).stdout.decode("utf-8", "replace")
    # Some builds link the framework relative to the app that contains them;
    # those need repointing at the real framework and re-signing. Builds that
    # already use an absolute path work as a plain copy.
    reference = None
    for line in old_path.splitlines():
        stripped = line.strip()
        if stripped.startswith("@executable_path"):
            reference = stripped.split()[0]
            break
    if not shutil.which("codesign"):
        os.remove(target)
        return False

    commands = []
    if reference is not None:
        if not shutil.which("install_name_tool"):
            os.remove(target)
            return False
        commands.append(["install_name_tool", "-change", reference, dylib, target])
    # Always re-sign. These stubs ship with the hardened runtime, and a copy
    # taken out of its signed bundle is killed on sight (SIGKILL) until it
    # carries a signature of its own.
    commands.append(["codesign", "-s", "-", "--force", target])

    for command in commands:
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode != 0:
            os.remove(target)
            return False

    check = subprocess.run([target, "-c", "import tkinter"],
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if check.returncode != 0:
        os.remove(target)
        return False
    return True


def build_app():
    app = os.path.join(HERE, APP_NAME + ".app")
    if os.path.exists(app):
        shutil.rmtree(app)
    macos = os.path.join(app, "Contents", "MacOS")
    resources = os.path.join(app, "Contents", "Resources")
    os.makedirs(macos)
    os.makedirs(resources)

    work = os.path.join(HERE, ".icon_build")
    if os.path.exists(work):
        shutil.rmtree(work)
    os.makedirs(work)
    try:
        master_png = build_icns(work, os.path.join(resources, "KIScore.icns"))
        # a GIF copy for Tk's own window icon on Windows and Linux
        sips("-s", "format", "gif", "--resampleHeightWidth", "128", "128",
             master_png, "--out", os.path.join(HERE, "kiscore_icon.gif"))
    finally:
        shutil.rmtree(work, ignore_errors=True)

    with io.open(os.path.join(app, "Contents", "Info.plist"), "w", encoding="utf-8") as handle:
        handle.write(INFO_PLIST % {"name": APP_NAME})
    chosen = find_interpreter()
    if chosen is None:
        raise SystemExit("No Python with tkinter was found.")
    interpreter, tk_version, framework = chosen
    print("Using %s (Tk %.1f)" % (interpreter, tk_version))
    if tk_version < 8.6:
        print("Warning: Tk %.1f does not support trackpad scrolling or Retina "
              "rendering. Install a newer Python (python.org or Homebrew) and "
              "run this builder again." % tk_version)
    embedded = embed_interpreter(macos, framework)
    launcher = os.path.join(macos, "launcher")
    with io.open(launcher, "w", encoding="utf-8") as handle:
        handle.write((LAUNCHER % SCRIPT) if embedded
                     else (FALLBACK_LAUNCHER % (interpreter, SCRIPT)))
    os.chmod(launcher, 0o755)
    if not embedded:
        print("Note: could not embed the interpreter, so macOS will still show "
              "the Python icon for this app.")

    # make Finder notice the new icon straight away
    subprocess.run(["touch", app], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return app


if __name__ == "__main__":
    if sys.platform != "darwin":
        raise SystemExit("This builder is for macOS. On Windows and Linux the window "
                         "icon is set from kiscore_icon.gif by the app itself.")
    built = build_app()
    print("Built: %s" % built)
    print("Icon:  %s" % os.path.join(HERE, "kiscore_icon.gif"))
