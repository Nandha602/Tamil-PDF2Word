"""
PDF2Word - OCR PDF to Word converter with Tamil support.

Turns PDFs into editable Word (.docx) files with real Unicode text using Tesseract OCR.
Removes faint watermarks behind the text, keeps pictures and diagrams, and downloads
the OCR engine and language data by itself on first use.

SPDX-License-Identifier: AGPL-3.0-only
Copyright (c) 2026 Nandha

Run with no arguments for the window, or:
    pdf2word.py input.pdf [more.pdf] -o OutputFolder
"""
import argparse
import csv
import io
import os
import re
import subprocess
import sys
import math
import difflib
import shutil
import tempfile
import threading
import urllib.request
import queue
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

APP_NAME = "PDF2Word"
VERSION = "1.2.0"

# ----------------------------------------------------------------------------
# Locating Tesseract (bundled next to the program, otherwise on PATH)
# ----------------------------------------------------------------------------
def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def data_dir() -> Path:
    """Folder for the downloaded OCR engine + language data (next to the program if writable)."""
    base = app_dir()
    if os.access(base, os.W_OK):
        return base / "PDF2Word-data"
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "PDF2Word"


def tesseract_candidates():
    exe = "tesseract.exe" if os.name == "nt" else "tesseract"
    yield data_dir() / "tesseract" / exe
    yield app_dir() / "tesseract" / exe
    yield app_dir() / exe
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        if os.environ.get(env):
            yield Path(os.environ[env]) / "Tesseract-OCR" / exe
    if os.environ.get("LOCALAPPDATA"):
        yield Path(os.environ["LOCALAPPDATA"]) / "Programs" / "Tesseract-OCR" / exe


def find_tesseract():
    """Returns (command, tessdata_dir_or_None). Uses our own language data when present."""
    own = data_dir() / "tessdata"
    for c in tesseract_candidates():
        if c.exists():
            if (own / "eng.traineddata").exists():
                return str(c), str(own)
            sib = c.parent / "tessdata"
            return str(c), (str(sib) if sib.exists() else None)
    exe = shutil.which("tesseract")
    if exe:
        return exe, (str(own) if (own / "eng.traineddata").exists() else None)
    return None, None


TESSDATA_URL = "https://github.com/tesseract-ocr/tessdata_best/raw/main/{}.traineddata"
WINGET_ID = "UB-Mannheim.TesseractOCR"
FALLBACK_INSTALLERS = [
    "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.0.20221222.exe",
    "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.2.0.20220712.exe",
]


def download(url, dest: Path, progress=None):
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".part")
    req = urllib.request.Request(url, headers={"User-Agent": "PDF2Word"})
    with urllib.request.urlopen(req, timeout=60) as r, open(tmp, "wb") as f:
        total = int(r.headers.get("Content-Length") or 0)
        done = 0
        while True:
            chunk = r.read(1 << 16)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)
    os.replace(tmp, dest)


def install_tesseract(log, progress=None):
    """Install the Tesseract engine (Windows). Tries winget first, then a direct download."""
    if os.name != "nt":
        return "Please install Tesseract with your package manager (for example: sudo apt install tesseract-ocr)."
    target = data_dir() / "tesseract"
    winget = shutil.which("winget")
    if winget:
        log("Installing the OCR engine with winget (a Windows permission prompt may appear)...")
        try:
            subprocess.run([winget, "install", "-e", "--id", WINGET_ID, "--silent",
                            "--accept-package-agreements", "--accept-source-agreements",
                            "--location", str(target)],
                           capture_output=True, creationflags=NO_WINDOW, timeout=900)
        except Exception as e:
            log(f"  winget did not finish: {e}")
        if find_tesseract()[0]:
            return None
    for url in FALLBACK_INSTALLERS:
        try:
            log("Downloading the OCR engine installer...")
            inst = Path(tempfile.gettempdir()) / "tesseract-setup.exe"
            download(url, inst, progress)
        except Exception as e:
            log(f"  could not download: {e}")
            continue
        log("Running the installer (a Windows permission prompt may appear)...")
        try:
            subprocess.run([str(inst), "/S", f"/D={target}"], creationflags=NO_WINDOW, timeout=900)
        except OSError:
            # installer needs administrator rights: let Windows show its permission prompt
            import ctypes, time
            ctypes.windll.shell32.ShellExecuteW(None, "runas", str(inst), f"/S /D={target}", None, 0)
            for _ in range(180):
                if find_tesseract()[0]:
                    break
                time.sleep(1)
        if find_tesseract()[0]:
            return None
    return ("The OCR engine could not be installed automatically. Install Tesseract from "
            "https://github.com/UB-Mannheim/tesseract/wiki and start the program again.")


def ensure_engine(lang, log=print, progress=None):
    """Download whatever is missing (engine + language data). Returns an error text or None."""
    try:
        for code in sorted(set(lang.split("+")) | {"eng"}):
            f = data_dir() / "tessdata" / f"{code}.traineddata"
            if not f.exists():
                log(f"Downloading language data: {code} ...")
                download(TESSDATA_URL.format(code), f, progress)
        if not find_tesseract()[0]:
            return install_tesseract(log, progress)
    except Exception as e:
        return f"Download failed: {e}. Please check your internet connection and try again."
    return None


NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


# ----------------------------------------------------------------------------
# Watermark / illustration detection
# ----------------------------------------------------------------------------
def pixmap_to_pil(doc, xref):
    pix = fitz.Pixmap(doc, xref)
    try:
        smask = doc.extract_image(xref).get("smask", 0)
    except Exception:
        smask = 0
    if pix.colorspace is not None and pix.colorspace.n > 3:
        pix = fitz.Pixmap(fitz.csRGB, pix)
    if smask and not pix.alpha:
        pix = fitz.Pixmap(pix, fitz.Pixmap(doc, smask))
    mode = "RGBA" if pix.alpha else ("RGB" if pix.n >= 3 else "L")
    im = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    if mode == "RGBA":
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[3])
        im = bg
    return im.convert("RGB")


def is_pale(im: Image.Image) -> bool:
    """True if the picture has (almost) nothing dark in it -> typical of a watermark."""
    small = im.copy()
    small.thumbnail((200, 200))
    hist = small.convert("L").histogram()
    total = sum(hist) or 1
    return sum(hist[:150]) / total < 0.003


def analyse_document(doc, log):
    """Work out which embedded images are the watermark (repeated on most pages / big and pale)."""
    counts, info = {}, {}
    n = len(doc)
    for p in doc:
        seen = set()
        for it in p.get_image_info(xrefs=True):
            x = it.get("xref", 0)
            if x and x not in seen:
                counts[x] = counts.get(x, 0) + 1
                seen.add(x)
    watermark = set()
    for x, c in counts.items():
        if n >= 3 and c >= max(2, 0.4 * n):
            watermark.add(x)
    return watermark


def page_illustrations(doc, page, watermark_xrefs, cache):
    """Return list of (bbox_in_pt, PIL image) for real pictures on the page."""
    out = []
    page_area = page.rect.width * page.rect.height
    for it in page.get_image_info(xrefs=True):
        x = it.get("xref", 0)
        bbox = fitz.Rect(it["bbox"])
        if bbox.width < 30 or bbox.height < 30:
            continue
        if x in watermark_xrefs:
            continue
        if bbox.width * bbox.height > 0.70 * page_area:
            continue          # a full-page picture is a scanned page, not an illustration
        try:
            if x:
                if x not in cache:
                    cache[x] = pixmap_to_pil(doc, x)
                im = cache[x]
            else:
                im = None
        except Exception:
            im = None
        if im is None:
            continue
        big = bbox.width * bbox.height > 0.20 * page_area
        if is_pale(im) and (big or x in watermark_xrefs):
            continue          # a big, pale picture is the watermark
        if is_pale(im) and bbox.width * bbox.height > 0.08 * page_area:
            continue
        out.append((bbox, im))

    # vector drawings (diagrams drawn with lines, not pictures)
    try:
        for r in page.cluster_drawings():
            r = fitz.Rect(r)
            if r.width < 45 or r.height < 45:
                continue                       # lines, underlines, borders
            if r.width * r.height > 0.5 * page_area:
                continue                       # page-sized frames / big logo
            if any(r.intersects(b) and (r & b).get_area() > 0.5 * r.get_area() for b, _ in out):
                continue
            pix = page.get_pixmap(clip=r, dpi=200, alpha=False)
            im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            if is_pale(im):
                continue
            # wash out faint watermark pixels that sit behind the drawing
            px = im.load()
            for yy in range(im.height):
                for xx in range(im.width):
                    r_, g_, b_ = px[xx, yy]
                    if min(r_, g_, b_) > 200 and max(r_, g_, b_) - min(r_, g_, b_) < 60:
                        px[xx, yy] = (255, 255, 255)
            out.append((r, im))
    except Exception:
        pass
    return out


# ----------------------------------------------------------------------------
# OCR
# ----------------------------------------------------------------------------
DOTS_RE = re.compile(r"[.\u2026\u00b7_\-\u2022\s]{3,}")


def leader_rects(page):
    """Rectangles of dotted leaders (.........) from the PDF text layer, e.g. in contents pages."""
    rects = []
    try:
        for w in page.get_text("words"):
            if DOTS_RE.fullmatch(w[4]):
                rects.append(fitz.Rect(w[:4]))
    except Exception:
        pass
    return rects


def prepare_page_image(page, dpi, threshold, illustrations, top_pct=0.0, bottom_pct=0.0, extra_rects=()):
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L")
    # kill everything lighter than the threshold (the faint watermark), keep dark text
    lut = [255 if v > threshold else v for v in range(256)]
    im = im.point(lut)
    white = Image.new("L", (10, 10), 255)

    def blank(rect, pad=3):
        x0, y0 = max(0, int((rect.x0 - pad) * zoom)), max(0, int((rect.y0 - pad) * zoom))
        x1, y1 = min(im.width, int((rect.x1 + pad) * zoom)), min(im.height, int((rect.y1 + pad) * zoom))
        if x1 > x0 and y1 > y0:
            im.paste(white.resize((x1 - x0, y1 - y0)), (x0, y0))

    for bbox, _ in illustrations:
        blank(bbox)
    for r in leader_rects(page):
        blank(r, pad=1)
    ph, pw = page.rect.height, page.rect.width
    for r in extra_rects:
        blank(r, pad=2)
    if top_pct > 0:
        blank(fitz.Rect(0, 0, pw, ph * top_pct / 100.0), pad=0)
    if bottom_pct > 0:
        blank(fitz.Rect(0, ph * (1 - bottom_pct / 100.0), pw, ph), pad=0)
    return im


def run_tesseract(img_path, lang, psm, tess_cmd, tessdata):
    """Run Tesseract writing its table to a file (more reliable on Windows than reading stdout)."""
    base = os.path.splitext(img_path)[0] + "_ocr"
    cmd = [tess_cmd, img_path, base, "-l", lang, "--psm", str(psm), "--dpi", "300"]
    if tessdata:
        cmd += ["--tessdata-dir", tessdata]
    cmd += ["tsv"]
    env = dict(os.environ, OMP_THREAD_LIMIT="1")
    res = subprocess.run(cmd, capture_output=True, creationflags=NO_WINDOW, env=env)
    out_file = base + ".tsv"
    if res.returncode != 0 or not os.path.exists(out_file):
        raise RuntimeError("Tesseract failed: " + (res.stderr + res.stdout).decode("utf-8", "replace")[:500])
    with open(out_file, "rb") as f:
        data = f.read().decode("utf-8-sig", "replace")
    try:
        os.remove(out_file)
    except OSError:
        pass
    return data


def parse_words(tsv):
    # Windows builds of Tesseract can add a BOM and doubled carriage returns: normalise first
    text = tsv.replace("\ufeff", "").replace("\r", "")
    lines = text.split("\n")
    start = next((i for i, ln in enumerate(lines) if ln.startswith("level\t")), None)
    if start is None:
        if not text.strip():
            return []              # blank page
        raise RuntimeError("Unexpected OCR output: " + repr(text[:200]))
    rows = csv.reader(lines[start:], delimiter="\t", quoting=csv.QUOTE_NONE)
    header = next(rows)
    idx = {h.strip(): i for i, h in enumerate(header)}
    words = []
    for r in rows:
        if len(r) < len(header) or r[idx["level"]] != "5":
            continue
        txt = r[idx["text"]].replace("\u200c", "").replace("\u200d", "").strip()
        if not txt or float(r[idx["conf"]]) < 0:
            continue
        words.append({
            "text": txt, "conf": float(r[idx["conf"]]),
            "l": int(r[idx["left"]]), "t": int(r[idx["top"]]),
            "w": int(r[idx["width"]]), "h": int(r[idx["height"]]),
            "key": (r[idx["block_num"]], r[idx["par_num"]], int(r[idx["line_num"]])),
        })
    return words


def ocr_page(img_path, lang, tess_cmd, tessdata):
    """Worker: OCR one prepared page image and measure the ink density of every word."""
    words = parse_words(run_tesseract(img_path, lang, 3, tess_cmd, tessdata))
    try:
        im = Image.open(img_path)
        for w in words:
            box = (w["l"], w["t"], w["l"] + w["w"], w["t"] + w["h"])
            hist = im.crop(box).histogram()
            tot = sum(hist) or 1
            w["ink"] = sum(hist[:128]) / tot
    except Exception:
        for w in words:
            w["ink"] = 0.0
    return words


# ----------------------------------------------------------------------------
# Layout information taken from the PDF's own (hidden) text layer
# ----------------------------------------------------------------------------
BULLET_CHARS = set("\u2022\u00b7\u25aa\u25a0\u25e6\u25cf\u25cb\u27a2\u27a4\u2756\u2713\u2714\u25ba\u25b6\uf0b7\uf0a7\uf076\uf0d8\uf0fc\uf0a8")
OCR_BULLET_SURE = {"\u2022", "\u00b7", "\u25cf", "\u25a0", "\u25aa", "\u25e6", "\u25cb", "*", "\u00b0"}
OCR_BULLET_CONFUSIONS = {"e", "o", "O", "\u00ab", "\u00ae", "\u00a9", "\u00a2", "\u00a5", "+", "~", "\u00bb",
                         "-", "\u2014", "\u2013", "=", "#", "@"}
BOLD_RE = re.compile(r"bold|black|heavy|-bd\b|bld", re.I)
ITALIC_RE = re.compile(r"italic|oblique", re.I)


def pdf_layout(page):
    """Spans (size, bold, italic, colour, box) and bullet glyph boxes from the PDF text layer."""
    spans, bullets = [], []
    try:
        d = page.get_text("dict")
    except Exception:
        return spans, bullets
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            for sp in ln.get("spans", []):
                t = sp.get("text", "")
                if not t.strip():
                    continue
                box = fitz.Rect(sp["bbox"])
                font, flags, size = sp.get("font", ""), sp.get("flags", 0), float(sp.get("size", 0) or 0)
                st = t.strip()
                rec = {"box": box, "size": size,
                       "bold": bool(flags & 16) or bool(BOLD_RE.search(font)),
                       "italic": bool(flags & 2) or bool(ITALIC_RE.search(font)),
                       "color": int(sp.get("color", 0))}
                if len(st) <= 2 and (st[0] in BULLET_CHARS or re.search(r"symbol|wingdings|dingbat", font, re.I)):
                    bullets.append(box)
                    continue
                if len(st) > 1 and st[0] in BULLET_CHARS:
                    bullets.append(fitz.Rect(box.x0, box.y0, box.x0 + max(size * 0.7, 4), box.y1))
                spans.append(rec)
    return spans, bullets


def span_at(spans, cx, cy):
    best, area = None, None
    for sp in spans:
        b = sp["box"]
        if b.x0 - 1.5 <= cx <= b.x1 + 1.5 and b.y0 - 1.5 <= cy <= b.y1 + 1.5:
            a = b.width * b.height
            if area is None or a < area:
                best, area = sp, a
    if best is None:
        d = 14.0
        for sp in spans:
            b = sp["box"]
            dx = max(b.x0 - cx, 0, cx - b.x1)
            dy = max(b.y0 - cy, 0, cy - b.y1)
            if (dx * dx + dy * dy) ** 0.5 < d:
                best, d = sp, (dx * dx + dy * dy) ** 0.5
    return best


def pct(vals, p):
    vals = sorted(vals)
    return vals[min(len(vals) - 1, max(0, int(round(p * (len(vals) - 1)))))] if vals else 0.0


def median(vals, default=0.0):
    return pct(vals, 0.5) if vals else default


def hex_color(c):
    return "%02X%02X%02X" % ((c >> 16) & 255, (c >> 8) & 255, c & 255)


# ----------------------------------------------------------------------------
# Building lines (with formatting) for one page
# ----------------------------------------------------------------------------
JUNK_RUN = re.compile(r"([^\W\d])\1{2,}")


def is_junk(tok):
    if JUNK_RUN.search(tok):
        return True
    return bool(re.fullmatch(r"[A-Za-z]{12,}", tok)) and len(set(tok.lower())) <= 4


def cut_leader_junk(words):
    for i, w in enumerate(words):
        if is_junk(w["text"]):
            keep = words[:i]
            last = words[-1]
            if i < len(words) - 1 and re.fullmatch(r"\d{1,4}", last["text"]):
                keep.append(last)
            return keep
    return words


def build_lines(words, spans, bullets, zoom, prect, opts):
    pw, ph = prect.width, prect.height
    groups = {}
    for w in words:
        groups.setdefault(w["key"], []).append(w)
    have_layer = bool(spans)
    lines = []
    # Tesseract sometimes splits one visual line into separate blocks (numbers, bullets, labels):
    # glue fragments that sit on the same row and next to each other back together
    frags = []
    for ws in groups.values():
        ws.sort(key=lambda w: w["l"])
        frags.append({"ws": ws,
                      "mid": median([w["t"] + w["h"] / 2 for w in ws]),
                      "h": median([w["h"] for w in ws]),
                      "l": ws[0]["l"], "r": max(w["l"] + w["w"] for w in ws)})
    frags.sort(key=lambda f: f["mid"])
    rows = []
    for f in frags:
        row = rows[-1] if rows else None
        if row and abs(f["mid"] - row["mid"]) < 0.45 * min(f["h"], row["h"]):
            gap_h = 2.5 * max(f["h"], row["h"])
            if (-4 <= f["l"] - row["r"] < gap_h) or (-4 <= row["l"] - f["r"] < gap_h):
                row["ws"] += f["ws"]
                row["l"], row["r"] = min(row["l"], f["l"]), max(row["r"], f["r"])
                continue
        rows.append(f)
    for row in rows:
        ws = sorted(row["ws"], key=lambda w: w["l"])
        ws = cut_leader_junk(ws)
        if not ws:
            continue
        items = []
        for w in ws:
            l, t = w["l"] / zoom, w["t"] / zoom
            r, b = l + w["w"] / zoom, t + w["h"] / zoom
            sp = span_at(spans, (l + r) / 2, (t + b) / 2) if have_layer else None
            items.append({"text": w["text"], "l": l, "r": r, "t": t, "b": b, "ink": w.get("ink", 0.0),
                          "size": sp["size"] if sp else None,
                          "bold": sp["bold"] if sp else False,
                          "italic": sp["italic"] if sp else False,
                          "color": sp["color"] if sp else 0})
        mid = median([(i["t"] + i["b"]) / 2 for i in items])
        height = median([i["b"] - i["t"] for i in items])
        # bullet?
        bullet_x = None
        for bb in bullets:
            if bb.y0 - 2 <= mid <= bb.y1 + 2 and bb.x0 <= items[0]["l"] + 2 and bb.x0 >= items[0]["l"] - 40:
                bullet_x = bb.x0
                # drop OCR's reading of the bullet glyph
                while items and (items[0]["text"] in OCR_BULLET_SURE or items[0]["text"] in OCR_BULLET_CONFUSIONS
                                 or (items[0]["r"] <= bb.x1 + 3 and items[0]["l"] >= bb.x0 - 3)):
                    items.pop(0)
                break
        if bullet_x is None and items and items[0]["text"] in OCR_BULLET_SURE:
            bullet_x = items[0]["l"]
            items.pop(0)
        if not items:
            continue
        text = " ".join(i["text"] for i in items)
        if re.fullmatch(r"[\d\s\-\u2013|]{1,6}", text) and opts.get("drop_page_numbers") and mid > 0.85 * ph:
            continue
        low = text.lower()
        if any(ph_ in low for ph_ in opts.get("phrases", ())):
            continue
        if opts.get("top_pct", 0) > 0 and mid < ph * opts["top_pct"] / 100.0:
            continue
        if opts.get("bottom_pct", 0) > 0 and mid > ph * (1 - opts["bottom_pct"] / 100.0):
            continue
        sizes = [i["size"] for i in items if i["size"]]
        size = median(sizes) if sizes else max(6.0, height / 1.25)
        for i in items:
            if i["size"] is None or abs(i["size"] - size) < 0.6:
                i["size"] = size
        lines.append({"items": items, "l": items[0]["l"], "r": items[-1]["r"], "mid": mid,
                      "h": height, "size": size, "bullet_x": bullet_x,
                      "text": text})
    lines.sort(key=lambda ln: ln["mid"])

    # bold from ink density when the PDF has no usable information (scans)
    if not have_layer:
        dens = [i["ink"] for ln in lines for i in ln["items"] if len(i["text"]) >= 2 and i["ink"] > 0]
        base = median(dens)
        for ln in lines:
            ds = [i["ink"] for i in ln["items"] if len(i["text"]) >= 2]
            if base and len(ds) >= 1 and median(ds) > 1.35 * base:
                for i in ln["items"]:
                    i["bold"] = True
    return lines


def norm_key(text):
    t = re.sub(r"\d+", "#", text.lower())
    return re.sub(r"[^\w#]+", " ", t).strip()


def repeated_header_footer_ids(pages, band=0.12, share=0.5):
    """ids of lines in the top/bottom band whose (fuzzy) text appears on most pages."""
    n = len(pages)
    if n < 3:
        return set()
    groups = []
    for pi, p in enumerate(pages):
        h = p["rect"].height
        for ln in p["lines"]:
            side = "top" if ln["mid"] < band * h else ("bottom" if ln["mid"] > (1 - band) * h else None)
            key = norm_key(ln["text"])
            if side is None or not key:
                continue
            for g in groups:
                if g["side"] == side and (g["key"] == key or
                        difflib.SequenceMatcher(None, key, g["key"]).ratio() >= 0.8):
                    g["pages"].add(pi)
                    g["lines"].append(ln)
                    break
            else:
                groups.append({"side": side, "key": key, "pages": {pi}, "lines": [ln]})
    need = max(3, math.ceil(share * n))
    return {id(ln) for g in groups if len(g["pages"]) >= need for ln in g["lines"]}


# ----------------------------------------------------------------------------
# Lines -> paragraphs
# ----------------------------------------------------------------------------
def page_pitch(lines, size):
    """Typical distance between consecutive lines of this text size on the page."""
    dys = []
    for a, b in zip(lines, lines[1:]):
        if abs(a["size"] - size) < 0.7 and abs(b["size"] - size) < 0.7:
            dy = b["mid"] - a["mid"]
            if dy > size * 0.8:
                dys.append(dy)
    if not dys:
        return size * 1.5
    dys.sort()
    return median(dys[: max(1, (len(dys) + 1) // 2)])


def group_paragraphs(lines, ML, MR, pw):
    tol = 0.025 * pw
    paras, cur = [], None
    pitch_cache = {}

    def pitch_for(sz):
        k = round(sz * 2)
        if k not in pitch_cache:
            pitch_cache[k] = page_pitch(lines, sz)
        return pitch_cache[k]

    for ln in lines:
        new = cur is None
        if not new:
            prev = cur["lines"][-1]
            p = pitch_for(prev["size"])
            dy = ln["mid"] - prev["mid"]
            prev_full = prev["r"] >= MR - max(0.10 * pw, 3.5 * prev["size"])
            centered_prev = abs((prev["l"] + prev["r"]) / 2 - pw / 2) < tol and prev["l"] > ML + 0.06 * pw
            if (ln["bullet_x"] is not None or dy > 1.35 * p or abs(ln["size"] - prev["size"]) > 1.0
                    or not prev_full or centered_prev
                    or (ln["l"] > ML + 1.2 * ln["size"] and cur["lines"][0]["l"] <= ln["l"] - 1.0
                        and len(cur["lines"]) >= 1 and ln["l"] - ML > 1.2 * ln["size"]
                        and prev["l"] <= ln["l"] + 1.0 and cur["indented_first"] is False)):
                new = True
        if new:
            cur = {"lines": [ln], "indented_first": ln["l"] > ML + 1.2 * ln["size"]}
            paras.append(cur)
        else:
            cur["lines"].append(ln)

    out = []
    for pr in paras:
        ls = pr["lines"]
        first = ls[0]
        size = median([l["size"] for l in ls])
        pitch = pitch_for(size)
        rest_l = median([l["l"] for l in ls[1:]]) if len(ls) > 1 else first["l"]
        bullet = first["bullet_x"] is not None
        # alignment
        centered = all(abs((l["l"] + l["r"]) / 2 - pw / 2) < tol for l in ls) and first["l"] > ML + 0.05 * pw \
            and (pw - first["r"]) > (pw - MR) + 0.05 * pw
        if centered:
            align = WD_ALIGN_PARAGRAPH.CENTER
        elif len(ls) >= 2 and all(l["r"] >= MR - tol for l in ls[:-1]):
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif len(ls) == 1 and first["r"] >= MR - 0.01 * pw and first["l"] > ML + 0.3 * pw:
            align = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            align = WD_ALIGN_PARAGRAPH.LEFT
        if bullet:
            text_left = first["l"]
            left_indent = max(0.0, (rest_l if len(ls) > 1 else text_left) - ML)
            first_line = first["bullet_x"] - (rest_l if len(ls) > 1 else text_left)
            if first_line > -size * 0.5:
                first_line = -size * 0.9
        else:
            left_indent = max(0.0, rest_l - ML)
            first_line = first["l"] - rest_l
        if align == WD_ALIGN_PARAGRAPH.CENTER:
            left_indent, first_line = 0.0, 0.0
        out.append({"kind": "p", "lines": ls, "align": align, "bullet": bullet,
                    "left_indent": left_indent, "first_line": first_line, "pitch": pitch, "size": size,
                    "top": ls[0]["mid"] - pitch / 2, "bottom": ls[-1]["mid"] + pitch / 2})
    return out


# ----------------------------------------------------------------------------
# Font size calibration: make Word's Tamil font as wide as the original text
# ----------------------------------------------------------------------------
FONT_FILES = {
    "Nirmala UI": ["Nirmala.ttf", "NirmalaUI.ttf", "Nirmala.ttc"],
    "Latha": ["latha.ttf"],
    "Vijaya": ["vijaya.ttf"],
}
DEFAULT_SCALE = {"Nirmala UI": 0.92, "Latha": 1.0, "Vijaya": 1.0}


def find_font_file(font, bold=False):
    override = os.environ.get("PDF2WORD_FONT_FILE")
    if override and Path(override).exists():
        return override
    folder = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    for name in FONT_FILES.get(font, []):
        if (folder / name).exists():
            return str(folder / name)
    return None


def calibrate_sizes(samples, font, log):
    """samples: list of (size_bucket, text, width_pt). Returns {bucket: scale}, default."""
    path = find_font_file(font)
    default = DEFAULT_SCALE.get(font, 1.0)
    if not path or not samples:
        log(f"  (font file for {font} not found; using default size scale {default})")
        return {}, default
    try:
        import uharfbuzz as hb
        blob = hb.Blob.from_file_path(path)
        face = hb.Face(blob)
        hfont = hb.Font(face)
        upem = face.upem
    except Exception as e:
        log(f"  (size calibration unavailable: {e}; using default size scale {default})")
        return {}, default

    def width(text, size):
        buf = hb.Buffer()
        buf.add_str(text)
        buf.guess_segment_properties()
        hb.shape(hfont, buf, {})
        return sum(p.x_advance for p in buf.glyph_positions) * size / upem

    by_bucket = {}
    for bucket, text, w in samples:
        ref = width(text, bucket)
        if ref > 0 and w > 0:
            by_bucket.setdefault(bucket, []).append(min(1.4, max(0.5, w / ref)))
    allr = [r for v in by_bucket.values() for r in v]
    glob = median(allr, default) if allr else default
    res = {b: median(v, glob) for b, v in by_bucket.items() if len(v) >= 8}
    log(f"  font size scale for {font}: {glob:.2f}")
    return res, glob


# ----------------------------------------------------------------------------
# Word output
# ----------------------------------------------------------------------------
def set_fonts(rPr_owner_rpr, font):
    rFonts = rPr_owner_rpr.get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(a), font)


def style_run(run, font, size, bold=False, italic=False, color=0):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold or None
    run.font.italic = italic or None
    rPr = run._element.get_or_add_rPr()
    set_fonts(rPr, font)
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(int(round(size * 2))))
    rPr.append(szcs)
    if bold:
        rPr.append(OxmlElement("w:bCs"))
    if italic:
        rPr.append(OxmlElement("w:iCs"))
    if color and color != 0:
        run.font.color.rgb = RGBColor.from_string(hex_color(color))
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "ta-IN")
    rPr.append(lang)


def new_document(font, page_w, page_h, ml, mr, mt, mb):
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Pt(page_w), Pt(page_h)
    sec.left_margin, sec.right_margin = Pt(ml), Pt(mr)
    sec.top_margin, sec.bottom_margin = Pt(mt), Pt(mb)
    sec.header_distance = sec.footer_distance = Pt(8)
    st = d.styles["Normal"]
    st.font.name = font
    st.font.size = Pt(11)
    set_fonts(st.element.get_or_add_rPr(), font)
    pf = st.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    return d


def no_contextual_spacing(para):
    pPr = para._element.get_or_add_pPr()
    cs = OxmlElement("w:contextualSpacing")
    cs.set(qn("w:val"), "0")
    pPr.append(cs)


def runs_for(par, scale, glob_scale):
    """Merge consecutive words with identical formatting into runs."""
    out = []
    for ln in par["lines"]:
        for i in ln["items"]:
            bucket = round(i["size"] * 2) / 2
            sc = scale.get(bucket, glob_scale)
            size = max(5.0, round(i["size"] * sc * 2) / 2)
            fmt = (size, i["bold"], i["italic"], i["color"])
            if out and out[-1][1] == fmt:
                out[-1][0].append(i["text"])
            else:
                out.append(([i["text"]], fmt))
    return out


def add_paragraph(d, par, prev_bottom, first_on_page, font, scale, glob_scale):
    para = d.add_paragraph(style="List Bullet" if par["bullet"] else None)
    pf = para.paragraph_format
    pf.alignment = par["align"]
    pf.space_before = Pt(max(0.0, par["top"] - prev_bottom))
    pf.space_after = Pt(0)
    pf.left_indent = Pt(par["left_indent"])
    pf.first_line_indent = Pt(par["first_line"])
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(max(par["pitch"], par["size"] * 1.1))
    pf.widow_control = False
    pf.page_break_before = bool(first_on_page and first_on_page == "break")
    if par["bullet"]:
        no_contextual_spacing(para)
    runs = runs_for(par, scale, glob_scale)
    for k, (words, (size, bold, italic, color)) in enumerate(runs):
        text = " ".join(words) + (" " if k < len(runs) - 1 else "")
        style_run(para.add_run(text), font, size, bold, italic, color)
    return para


def add_picture(d, item, ml, text_w, prev_bottom, first_on_page):
    b = item["bbox"]
    w = min(b.width, text_w)
    h = b.height * (w / b.width)
    buf = io.BytesIO()
    item["im"].save(buf, "PNG")
    buf.seek(0)
    para = d.add_paragraph()
    pf = para.paragraph_format
    pf.space_before = Pt(max(0.0, b.y0 - prev_bottom))
    pf.space_after = Pt(0)
    pf.page_break_before = bool(first_on_page and first_on_page == "break")
    cx = (b.x0 + b.x1) / 2
    if abs(cx - (ml + text_w / 2)) < 0.03 * (text_w):
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.left_indent = Pt(max(0.0, b.x0 - ml))
    para.add_run().add_picture(buf, width=Pt(w), height=Pt(h))


# ----------------------------------------------------------------------------
# Main conversion
# ----------------------------------------------------------------------------
def convert_pdf(pdf_path, out_path, lang="tam+eng", dpi=300, threshold=165,
                drop_page_numbers=False, page_breaks=True, font="Nirmala UI",
                top_pct=0.0, bottom_pct=0.0, remove_repeated=False, remove_text="",
                progress=lambda done, total: None, log=print, cancel=lambda: False):
    tess_cmd, tessdata = find_tesseract()
    if not tess_cmd:
        raise RuntimeError("OCR engine not found. Start the program again and allow the download.")
    doc = fitz.open(pdf_path)
    total = len(doc)
    watermark = analyse_document(doc, log)
    cache = {}
    tmp = tempfile.mkdtemp(prefix="pdf2word_")
    zoom = dpi / 72.0
    workers = max(1, min(4, (os.cpu_count() or 2)))
    pool = ThreadPoolExecutor(max_workers=workers)
    phrases = [p.strip().lower() for p in re.split(r"[,\n]", remove_text or "") if p.strip()]
    opts = {"drop_page_numbers": drop_page_numbers, "top_pct": top_pct, "bottom_pct": bottom_pct,
            "phrases": phrases}
    jobs = []
    try:
        for i in range(total):
            if cancel():
                return None
            page = doc[i]
            ills = page_illustrations(doc, page, watermark, cache)
            if top_pct or bottom_pct:
                ph = page.rect.height
                ills = [(b, im) for b, im in ills
                        if ph * top_pct / 100.0 <= (b.y0 + b.y1) / 2 <= ph * (1 - bottom_pct / 100.0)]
            spans, bullets = pdf_layout(page)
            extra = []
            for phrase in phrases:
                try:
                    extra += page.search_for(phrase)
                except Exception:
                    pass
            im = prepare_page_image(page, dpi, threshold, ills, top_pct, bottom_pct, extra)
            img_path = os.path.join(tmp, f"p{i}.png")
            im.save(img_path)
            fut = pool.submit(ocr_page, img_path, lang, tess_cmd, tessdata)
            jobs.append((fut, ills, spans, bullets, page.rect))
            del im

        pages = []
        for i, (fut, ills, spans, bullets, prect) in enumerate(jobs):
            if cancel():
                return None
            words = fut.result()
            lines = build_lines(words, spans, bullets, zoom, prect, opts)
            pages.append({"lines": lines, "ills": ills, "rect": prect})
            progress(i + 1, total * 2)
            log(f"  page {i + 1}/{total} read")

        if remove_repeated:
            drop = repeated_header_footer_ids(pages)
            if drop:
                log(f"  removed {len(drop)} repeated header/footer lines")
            for p in pages:
                p["lines"] = [ln for ln in p["lines"] if id(ln) not in drop]

        # document-wide text area
        lefts = [ln["l"] for p in pages for ln in p["lines"] if ln["bullet_x"] is None]
        rights = [ln["r"] for p in pages for ln in p["lines"]]
        pw, ph = pages[0]["rect"].width, pages[0]["rect"].height
        ML = pct(lefts, 0.10) if lefts else 36.0
        MR = pct(rights, 0.90) if rights else pw - 36.0
        ML = max(0.0, min(ML, pw * 0.4))
        MR = min(pw, max(MR, pw * 0.6))

        all_items = []
        samples = []
        for pno, p in enumerate(pages):
            paras = group_paragraphs(p["lines"], ML, MR, pw)
            items = [(pr["top"], pr) for pr in paras]
            items += [(b.y0, {"kind": "i", "bbox": b, "im": im, "top": b.y0, "bottom": b.y1}) for b, im in p["ills"]]
            items.sort(key=lambda t: t[0])
            all_items.append([it for _, it in items])
            for ln in p["lines"]:
                for w in ln["items"]:
                    if len(w["text"]) >= 3 and not w["bold"] and re.search(r"[\u0B80-\u0BFF]", w["text"]):
                        samples.append((round(w["size"] * 2) / 2, w["text"], w["r"] - w["l"]))
        scale, glob = calibrate_sizes(samples, font, log)

        tops = [it["top"] for pg in all_items for it in pg[:1]]
        bottoms = [it["bottom"] for pg in all_items for it in pg[-1:]]
        MT = max(10.0, min(tops) if tops else 36.0)
        MB = max(10.0, ph - max(bottoms)) if bottoms else 36.0
        d = new_document(font, pw, ph, ML, pw - MR, MT, MB)
        text_w = MR - ML
        for pno, pg in enumerate(all_items):
            prev_bottom = MT
            for k, it in enumerate(pg):
                first = ("break" if (page_breaks and pno > 0) else "flow") if k == 0 else None
                if it["kind"] == "p":
                    add_paragraph(d, it, prev_bottom, first, font, scale, glob)
                else:
                    add_picture(d, it, ML, text_w, prev_bottom, first)
                prev_bottom = it["bottom"]
            progress(total + pno + 1, total * 2)
        d.save(out_path)
        return out_path
    finally:
        pool.shutdown(wait=False)
        for f in Path(tmp).glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
        try:
            os.rmdir(tmp)
        except Exception:
            pass


def check_tesseract(lang):
    """None if the OCR engine and languages are ready, otherwise a short problem description."""
    tess_cmd, tessdata = find_tesseract()
    if not tess_cmd:
        return "The OCR engine is not installed yet."
    cmd = [tess_cmd, "--list-langs"]
    if tessdata:
        cmd += ["--tessdata-dir", tessdata]
    try:
        res = subprocess.run(cmd, capture_output=True, creationflags=NO_WINDOW)
    except OSError:
        return "The OCR engine could not be started."
    have = set(res.stdout.decode("utf-8", "replace").split())
    missing = [l for l in lang.split("+") if l not in have]
    if missing:
        return "Language data is missing: " + ", ".join(missing)
    return None


# ----------------------------------------------------------------------------
# Window
# ----------------------------------------------------------------------------
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title(f"{APP_NAME} {VERSION}")
    root.geometry("760x720")
    files = []
    q = queue.Queue()
    cancel_flag = threading.Event()

    out_var = tk.StringVar(value="")
    thr_var = tk.IntVar(value=165)
    dpi_var = tk.StringVar(value="300")
    num_var = tk.BooleanVar(value=False)
    rep_var = tk.BooleanVar(value=False)
    txt_var = tk.StringVar(value="")
    pb_var = tk.BooleanVar(value=True)
    font_var = tk.StringVar(value="Nirmala UI")
    top_var = tk.DoubleVar(value=0.0)
    bot_var = tk.DoubleVar(value=0.0)
    lang_var = tk.StringVar(value="tam+eng")

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="1. Choose your PDF files").grid(row=0, column=0, sticky="w")
    lb = tk.Listbox(frm, height=6, selectmode="extended")
    lb.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=4)

    def add_files():
        for f in filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")]):
            if f not in files:
                files.append(f)
                lb.insert("end", f)

    def clear_files():
        files.clear()
        lb.delete(0, "end")

    ttk.Button(frm, text="Add PDFs...", command=add_files).grid(row=2, column=0, sticky="w")
    ttk.Button(frm, text="Clear list", command=clear_files).grid(row=2, column=1, sticky="w")

    ttk.Label(frm, text="2. Save Word files to (empty = next to each PDF)").grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
    ttk.Entry(frm, textvariable=out_var).grid(row=4, column=0, columnspan=2, sticky="ew")
    ttk.Button(frm, text="Browse...", command=lambda: out_var.set(filedialog.askdirectory() or out_var.get())).grid(row=4, column=2)

    opt = ttk.LabelFrame(frm, text="3. Options", padding=6)
    opt.grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)
    ttk.Label(opt, text="Watermark removal strength (higher = removes more, 165 is good):").grid(row=0, column=0, sticky="w")
    ttk.Scale(opt, from_=110, to=215, variable=thr_var, orient="horizontal", length=160).grid(row=0, column=1, padx=6)
    ttk.Label(opt, text="Scan quality (DPI):").grid(row=1, column=0, sticky="w")
    ttk.Combobox(opt, textvariable=dpi_var, values=["200", "300", "400"], width=6, state="readonly").grid(row=1, column=1, sticky="w", padx=6)
    ttk.Label(opt, text="OCR language:").grid(row=2, column=0, sticky="w")
    ttk.Combobox(opt, textvariable=lang_var, values=["tam+eng", "tam", "eng"], width=8, state="readonly").grid(row=2, column=1, sticky="w", padx=6)
    ttk.Label(opt, text="Word font:").grid(row=3, column=0, sticky="w")
    ttk.Combobox(opt, textvariable=font_var, values=list(FONT_FILES), width=12, state="readonly").grid(row=3, column=1, sticky="w", padx=6)
    ttk.Label(opt, text="Ignore top / bottom % of each page (headers, footers):").grid(row=4, column=0, sticky="w")
    sp = ttk.Frame(opt)
    sp.grid(row=4, column=1, sticky="w", padx=6)
    ttk.Spinbox(sp, from_=0, to=30, increment=1, textvariable=top_var, width=4).pack(side="left")
    ttk.Spinbox(sp, from_=0, to=30, increment=1, textvariable=bot_var, width=4).pack(side="left", padx=4)
    ttk.Checkbutton(opt, text="Remove page numbers", variable=num_var).grid(row=5, column=0, sticky="w")
    ttk.Checkbutton(opt, text="Keep the PDF's page breaks (one PDF page = one Word page)", variable=pb_var).grid(row=6, column=0, sticky="w")
    ttk.Checkbutton(opt, text="Remove repeated headers / footers (same text on most pages)", variable=rep_var).grid(row=7, column=0, sticky="w")
    ttk.Label(opt, text="Also remove text containing (separate with commas, e.g. DRAFT, Confidential):").grid(row=8, column=0, columnspan=2, sticky="w")
    ttk.Entry(opt, textvariable=txt_var).grid(row=9, column=0, columnspan=2, sticky="ew", pady=(0, 2))

    pbar = ttk.Progressbar(frm, mode="determinate")
    pbar.grid(row=6, column=0, columnspan=3, sticky="ew")
    logbox = tk.Text(frm, height=8, state="disabled")
    logbox.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=6)
    frm.rowconfigure(7, weight=1)
    frm.columnconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    def log(msg):
        q.put(("log", msg))

    def worker():
        err = check_tesseract(lang_var.get())
        if err:
            q.put(("log", err + ". Downloading it now (one time only)..."))
            err = ensure_engine(lang_var.get(), log=log,
                                progress=lambda d, t: q.put(("prog", (d, t or d))))
            err = err or check_tesseract(lang_var.get())
        if err:
            q.put(("log", "ERROR: " + err))
            q.put(("done", None))
            return
        for n, f in enumerate(list(files), 1):
            if cancel_flag.is_set():
                break
            src = Path(f)
            outdir = Path(out_var.get()) if out_var.get() else src.parent
            outdir.mkdir(parents=True, exist_ok=True)
            dest = outdir / (src.stem + ".docx")
            q.put(("log", f"[{n}/{len(files)}] {src.name}"))
            try:
                convert_pdf(str(src), str(dest), lang=lang_var.get(), dpi=int(dpi_var.get()),
                            threshold=int(thr_var.get()), drop_page_numbers=num_var.get(),
                            page_breaks=pb_var.get(), font=font_var.get(),
                            top_pct=float(top_var.get()), bottom_pct=float(bot_var.get()),
                            remove_repeated=rep_var.get(), remove_text=txt_var.get(),
                            progress=lambda d, t: q.put(("prog", (d, t))),
                            log=log, cancel=cancel_flag.is_set)
                q.put(("log", f"  saved: {dest}"))
            except Exception as e:
                q.put(("log", f"  FAILED: {e}"))
        q.put(("done", None))

    def start():
        if not files:
            messagebox.showinfo(APP_NAME, "Please add at least one PDF first.")
            return
        if check_tesseract(lang_var.get()):
            if not messagebox.askyesno(APP_NAME, "The first run needs to download the OCR engine and "
                                       "language data (about 40 MB, one time only).\n\nDownload now?"):
                return
        cancel_flag.clear()
        go.config(state="disabled")
        threading.Thread(target=worker, daemon=True).start()

    go = ttk.Button(frm, text="Convert to Word", command=start)
    go.grid(row=8, column=0, sticky="w")
    ttk.Button(frm, text="Stop", command=cancel_flag.set).grid(row=8, column=1, sticky="w")

    def poll():
        try:
            while True:
                kind, val = q.get_nowait()
                if kind == "log":
                    logbox.config(state="normal")
                    logbox.insert("end", val + "\n")
                    logbox.see("end")
                    logbox.config(state="disabled")
                elif kind == "prog":
                    pbar["maximum"], pbar["value"] = val[1], val[0]
                elif kind == "done":
                    go.config(state="normal")
                    messagebox.showinfo(APP_NAME, "Finished.")
        except queue.Empty:
            pass
        root.after(150, poll)

    poll()
    root.mainloop()


def main():
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("pdfs", nargs="*")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--lang", default="tam+eng")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--threshold", type=int, default=165)
    ap.add_argument("--drop-page-numbers", action="store_true")
    ap.add_argument("--no-page-breaks", action="store_true")
    ap.add_argument("--font", default="Nirmala UI", choices=list(FONT_FILES))
    ap.add_argument("--remove-repeated", action="store_true", help="remove headers/footers repeated on most pages")
    ap.add_argument("--remove-text", default="", help="comma-separated phrases; lines containing them are removed")
    ap.add_argument("--ignore-top", type=float, default=0.0, help="percent of page height to ignore at the top")
    ap.add_argument("--ignore-bottom", type=float, default=0.0, help="percent of page height to ignore at the bottom")
    a = ap.parse_args()
    if not a.pdfs:
        run_gui()
        return
    if check_tesseract(a.lang):
        print("Downloading the OCR engine (one time only)...")
        err = ensure_engine(a.lang) or check_tesseract(a.lang)
        if err:
            sys.exit(err)
    for f in a.pdfs:
        src = Path(f)
        outdir = Path(a.out) if a.out else src.parent
        outdir.mkdir(parents=True, exist_ok=True)
        dest = outdir / (src.stem + ".docx")
        print(src.name)
        convert_pdf(str(src), str(dest), a.lang, a.dpi, a.threshold,
                    a.drop_page_numbers, not a.no_page_breaks, a.font, a.ignore_top, a.ignore_bottom,
                    a.remove_repeated, a.remove_text)
        print("saved", dest)


if __name__ == "__main__":
    main()
