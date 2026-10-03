#!/usr/bin/env python3
"""Render a .pptx to PNG previews + a contact sheet using Microsoft PowerPoint (macOS).

Static frames only: they show the final resting state of every slide, which is
what you check for layout, overflow and contrast. Morph and animations can only
be judged by playing the deck in PowerPoint.

Usage:
    python3 preview.py deck.pptx OUT_DIR [--restart]

Needs: Microsoft PowerPoint for Mac, `pip install pymupdf pillow`.
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import time

# PowerPoint is sandboxed: it can only write inside its own container.
BOX = os.path.expanduser("~/Library/Containers/com.microsoft.Powerpoint/Data/morphdeck")


def export_pdf(pptx, restart=False):
    os.makedirs(BOX, exist_ok=True)
    src = os.path.join(BOX, "preview.pptx")
    pdf = os.path.join(BOX, "preview.pdf")
    shutil.copy(pptx, src)
    if os.path.exists(pdf):
        os.remove(pdf)
    if restart:  # PowerPoint only picks up newly installed fonts after a restart
        subprocess.run(["osascript", "-e", 'tell application "Microsoft PowerPoint" to quit'])
        time.sleep(3)
    script = f'''
    with timeout of 180 seconds
      tell application "Microsoft PowerPoint"
        open POSIX file "{src}"
        delay 3
        set p to active presentation
        save p in POSIX file "{pdf}" as save as PDF
        close p saving no
      end tell
    end timeout'''
    subprocess.run(["osascript", "-e", script], check=True)
    if not os.path.exists(pdf):
        sys.exit("PowerPoint did not produce a PDF (did it show a repair dialog?)")
    return pdf


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("out_dir")
    ap.add_argument("--restart", action="store_true", help="quit PowerPoint first (reloads fonts)")
    ap.add_argument("--dpi", type=int, default=60)
    a = ap.parse_args()

    import pymupdf
    from PIL import Image

    pdf = export_pdf(os.path.abspath(a.pptx), a.restart)
    os.makedirs(a.out_dir, exist_ok=True)
    for f in glob.glob(os.path.join(a.out_dir, "slide-*.png")):
        os.remove(f)
    doc = pymupdf.open(pdf)
    paths = []
    for i, page in enumerate(doc):
        p = os.path.join(a.out_dir, f"slide-{i + 1:02d}.png")
        page.get_pixmap(dpi=a.dpi).save(p)
        paths.append(p)
    ims = [Image.open(p) for p in paths]
    w, h = ims[0].size
    rows = (len(ims) + 1) // 2
    sheet = Image.new("RGB", (w * 2 + 12, (h + 12) * rows), "white")
    for i, im in enumerate(ims):
        sheet.paste(im, ((i % 2) * (w + 12), (i // 2) * (h + 12)))
    out = os.path.join(a.out_dir, "contact-sheet.png")
    sheet.save(out)
    print(f"{len(paths)} slides -> {out}")


if __name__ == "__main__":
    main()
