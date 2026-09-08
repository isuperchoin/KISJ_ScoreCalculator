"""
Build a self-contained "KISJ Score Calculator.app" that runs on any Mac.

Why this exists alongside build_macos_app.py: that builder makes an app for
*this* computer. Its launcher points at kisj_grade_calculator.py by absolute
path and borrows the Python framework installed here, so copying it to a
friend's Mac gives them a bundle that cannot find either. This builder uses
PyInstaller to copy the interpreter, the standard library and Tk *into* the
bundle, so the result needs nothing installed on the machine it lands on.

    /Library/Frameworks/Python.framework/Versions/3.14/bin/python3 build_shareable_app.py

Produces, next to this file:
    KISJ Score Calculator.app   self-contained, universal (Intel + Apple Silicon)
    KISJ Score Calculator.zip   the same app, packed for sending to someone

The build has to run under a python.org framework Python, not pyenv or the
system one: pyenv builds here have no tkinter at all, and Apple's /usr/bin
Python is stuck on Tk 8.5, which has no trackpad scrolling and no Retina
rendering. The framework build is also universal2, which is what lets the
finished app run on Intel Macs as well as Apple Silicon.
"""

import os
import shutil
import subprocess
import sys

import build_macos_app as icons

HERE = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "KISJ Score Calculator"
WORK = os.path.join(HERE, ".build")


def check_interpreter():
    if sys.platform != "darwin":
        raise SystemExit("This builder is for macOS.")
    try:
        import tkinter
    except ImportError:
        raise SystemExit(
            "This Python has no tkinter. Run the builder with a python.org "
            "framework build, e.g.\n"
            "  /Library/Frameworks/Python.framework/Versions/3.14/bin/python3 "
            "build_shareable_app.py")
    if tkinter.TkVersion < 8.6:
        raise SystemExit("Tk %.1f is too old; install Python from python.org."
                         % tkinter.TkVersion)
    try:
        import PyInstaller            # noqa: F401
    except ImportError:
        raise SystemExit("PyInstaller is missing. Install it with:\n"
                         "  %s -m pip install pyinstaller" % sys.executable)


def universal():
    """True when this interpreter can produce an Intel + Apple Silicon build."""
    framework = os.path.join(sys.base_prefix, "Python")
    if not os.path.exists(framework):
        return False
    archs = subprocess.run(["lipo", "-archs", framework], stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE).stdout.decode("utf-8", "replace")
    return "x86_64" in archs and "arm64" in archs


def main():
    check_interpreter()
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK)

    # Reuse the icon drawing from the other builder: the .icns for the bundle,
    # and the .gif that Tk shows as the window icon on Windows and Linux.
    icns = os.path.join(WORK, "KIScore.icns")
    master = icons.build_icns(WORK, icns)
    icons.sips("-s", "format", "gif", "--resampleHeightWidth", "128", "128",
               master, "--out", os.path.join(HERE, "kiscore_icon.gif"))

    command = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--windowed",                 # .app bundle, no terminal window
        "--noconfirm", "--clean",
        "--icon", icns,
        "--osx-bundle-identifier", "com.kisj.kiscore",
        "--add-data", os.path.join(HERE, "kiscore_icon.gif") + ":.",
        "--paths", HERE,
        "--hidden-import", "grade_export",
        "--workpath", os.path.join(WORK, "build"),
        "--distpath", os.path.join(WORK, "dist"),
        "--specpath", WORK,
        os.path.join(HERE, "kisj_grade_calculator.py"),
    ]
    if universal():
        command[3:3] = ["--target-arch", "universal2"]
    else:
        print("Note: this Python is not universal, so the app will only run on "
              "Macs with the same chip as this one.")

    if subprocess.run(command).returncode != 0:
        raise SystemExit("PyInstaller failed.")

    built = os.path.join(WORK, "dist", APP_NAME + ".app")
    final = os.path.join(HERE, APP_NAME + ".app")
    shutil.rmtree(final, ignore_errors=True)
    subprocess.run(["ditto", built, final], check=True)

    # Strip any quarantine flags picked up during the build, then ad-hoc sign
    # the whole bundle. Without a signature macOS refuses to start it at all on
    # Apple Silicon; the signature is not a Developer ID one, so recipients
    # still have to right-click > Open the first time.
    subprocess.run(["xattr", "-cr", final])
    subprocess.run(["codesign", "-s", "-", "--force", "--deep", final],
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # ditto, not zip: it keeps the symlinks and resource forks inside the
    # bundle intact, which a plain zip mangles.
    archive = os.path.join(HERE, APP_NAME + ".zip")
    if os.path.exists(archive):
        os.remove(archive)
    subprocess.run(["ditto", "-c", "-k", "--keepParent", "--sequesterRsrc",
                    final, archive], check=True)

    shutil.rmtree(WORK, ignore_errors=True)
    print("\nBuilt: %s" % final)
    print("Send:  %s" % archive)


if __name__ == "__main__":
    main()
