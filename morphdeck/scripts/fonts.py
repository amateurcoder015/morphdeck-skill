#!/usr/bin/env python3
"""Install the bundled Unbounded font for the current user if it is missing.

Called automatically by build_deck.py (skip with --no-font-install), or run on
its own:  python3 fonts.py
Unbounded is licensed under the SIL Open Font License (see ../fonts/OFL.txt).
"""
import glob
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLED = os.path.join(HERE, "..", "fonts", "Unbounded[wght].ttf")
FAMILY = "Unbounded"


def _font_dirs():
    if sys.platform == "darwin":
        return [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts"]
    if sys.platform.startswith("win"):
        return [os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
                os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")]
    return [os.path.expanduser("~/.local/share/fonts"), os.path.expanduser("~/.fonts"), "/usr/share/fonts"]


def installed():
    for d in _font_dirs():
        if glob.glob(os.path.join(d, "**", FAMILY + "*"), recursive=True):
            return True
    return False


def install():
    """Copy the bundled font into the per-user font folder. Returns a status line."""
    if installed():
        return None
    if not os.path.exists(BUNDLED):
        return f"warning: {FAMILY} font not installed and bundled copy not found"
    target_dir = _font_dirs()[0]
    os.makedirs(target_dir, exist_ok=True)
    target = os.path.join(target_dir, os.path.basename(BUNDLED))
    shutil.copy(BUNDLED, target)
    if sys.platform.startswith("win"):  # per-user fonts must also be registered
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows NT\CurrentVersion\Fonts", 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, f"{FAMILY} (TrueType)", 0, winreg.REG_SZ, target)
        winreg.CloseKey(key)
    elif not sys.platform == "darwin" and shutil.which("fc-cache"):
        os.system("fc-cache -f >/dev/null 2>&1")
    return f"installed {FAMILY} font to {target} (restart PowerPoint if it is open)"


if __name__ == "__main__":
    print(install() or f"{FAMILY} is already installed")
