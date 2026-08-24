#!/usr/bin/env python3
"""
OCR the files that carry no text layer — scanned PDFs, TIFF and JPG drawings.

Uses the macOS Vision framework: on-device, no network, no model download,
German and English out of the box, and considerably better on engineering
drawings than a stock Tesseract install.

Writes one sidecar per file:

    out/ocr/<sha1>.json   {source_path, filename, pages:[{page, text}]}

extract.py picks these up automatically on the next run, so the OCR'd text
joins the same index with the same metadata.

    python ocr.py --list      # what would be processed
    python ocr.py             # run it
    python ocr.py --force     # redo files already done
"""
from __future__ import annotations

import argparse
import hashlib
import json
import platform
import sys
import time
from pathlib import Path

# --------------------------------------------------------------- backends
# macOS  → Vision framework (built in, no download, excellent German)
# other  → RapidOCR (ONNX, pip-installable, no system binary)
IS_MAC = platform.system() == "Darwin"

Quartz = Vision = NSURL = None
_rapid = None

if IS_MAC:
    try:
        import Quartz
        import Vision
        from Foundation import NSURL
    except ImportError:
        IS_MAC = False


def _rapid_engine():
    """Lazily construct the cross-platform OCR engine."""
    global _rapid
    if _rapid is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as e:
            raise SystemExit(
                "No OCR backend available.\n"
                "  macOS  : pip install pyobjc-framework-Vision "
                "pyobjc-framework-Quartz\n"
                "  Windows/Linux: pip install rapidocr-onnxruntime pillow "
                "pymupdf"
            ) from e
        _rapid = RapidOCR()
    return _rapid

ROOT = Path(__file__).resolve().parent.parent
OUT = Path(__file__).resolve().parent / "out"
OCR_DIR = OUT / "ocr"
MANIFEST = OUT / "manifest.json"

LANGS = ["de-DE", "en-US"]
DPI_SCALE = 2.0          # render PDF pages at 2x for legible small type


def sidecar_path(rel_path: str) -> Path:
    h = hashlib.sha1(rel_path.encode("utf-8")).hexdigest()[:16]
    return OCR_DIR / f"{h}.json"


# Vision (and most engines) return nothing at all on very large images.
# Engineering drawings here reach 500+ megapixels, so anything big is tiled.
MAX_PIXELS = 30_000_000
TILE_OVERLAP = 0.06          # keep text that straddles a tile edge


# ------------------------------------------------------------ macOS backend
def _vision_recognize(cg_image) -> str:
    handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(
        cg_image, None)
    req = Vision.VNRecognizeTextRequest.alloc().init()
    req.setRecognitionLevel_(0)              # 0 = accurate
    req.setRecognitionLanguages_(LANGS)
    req.setUsesLanguageCorrection_(True)
    ok, _err = handler.performRequests_error_([req], None)
    if not ok:
        return ""
    lines = []
    for obs in (req.results() or []):
        cand = obs.topCandidates_(1)
        if cand and len(cand):
            lines.append(cand[0].string())
    return "\n".join(lines)


def _vision_tiled(img) -> str:
    """OCR a CGImage, splitting into overlapping tiles when it is huge."""
    import math
    w = Quartz.CGImageGetWidth(img)
    h = Quartz.CGImageGetHeight(img)
    if w * h <= MAX_PIXELS:
        return _vision_recognize(img)

    parts = max(2, min(math.ceil(math.sqrt((w * h) / MAX_PIXELS)), 8))
    tw, th = w // parts, h // parts
    ox, oy = int(tw * TILE_OVERLAP), int(th * TILE_OVERLAP)

    seen, chunks = set(), []
    for r in range(parts):
        for c in range(parts):
            x, y = max(0, c * tw - ox), max(0, r * th - oy)
            cw, ch = min(tw + 2 * ox, w - x), min(th + 2 * oy, h - y)
            if cw <= 0 or ch <= 0:
                continue
            tile = Quartz.CGImageCreateWithImageInRect(
                img, Quartz.CGRectMake(x, y, cw, ch))
            if not tile:
                continue
            for line in _vision_recognize(tile).splitlines():
                s = line.strip()
                if s and s not in seen:
                    seen.add(s)
                    chunks.append(s)
    return "\n".join(chunks)


def _vision_pdf(path: Path) -> list[dict]:
    url = NSURL.fileURLWithPath_(str(path))
    doc = Quartz.CGPDFDocumentCreateWithURL(url)
    if not doc:
        return []
    pages = []
    for i in range(1, Quartz.CGPDFDocumentGetNumberOfPages(doc) + 1):
        page = Quartz.CGPDFDocumentGetPage(doc, i)
        if not page:
            continue
        rect = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)
        w = int(rect.size.width * DPI_SCALE)
        h = int(rect.size.height * DPI_SCALE)
        if w <= 0 or h <= 0 or w * h > 40_000_000:
            continue
        cs = Quartz.CGColorSpaceCreateDeviceRGB()
        ctx = Quartz.CGBitmapContextCreate(
            None, w, h, 8, 0, cs, Quartz.kCGImageAlphaNoneSkipLast)
        Quartz.CGContextSetRGBFillColor(ctx, 1, 1, 1, 1)
        Quartz.CGContextFillRect(ctx, Quartz.CGRectMake(0, 0, w, h))
        Quartz.CGContextScaleCTM(ctx, DPI_SCALE, DPI_SCALE)
        Quartz.CGContextDrawPDFPage(ctx, page)
        text = _vision_tiled(Quartz.CGBitmapContextCreateImage(ctx))
        if text.strip():
            pages.append({"page": str(i), "text": text})
    return pages


def _vision_image(path: Path) -> list[dict]:
    src = Quartz.CGImageSourceCreateWithURL(
        NSURL.fileURLWithPath_(str(path)), None)
    if not src or Quartz.CGImageSourceGetCount(src) == 0:
        return []
    count = Quartz.CGImageSourceGetCount(src)
    out = []
    for i in range(count):                              # multi-page TIFF
        img = Quartz.CGImageSourceCreateImageAtIndex(src, i, None)
        if not img:
            continue
        text = _vision_tiled(img)
        if text.strip():
            out.append({"page": str(i + 1) if count > 1 else "", "text": text})
    return out


# ------------------------------------- cross-platform backend (Windows/Linux)
def _pil_tiled(pil_img) -> str:
    """Same tiling idea, on a PIL image, through RapidOCR."""
    import math

    import numpy as np
    engine = _rapid_engine()

    def run(im) -> list[str]:
        res, _ = engine(np.array(im))
        return [line[1] for line in (res or []) if len(line) > 1]

    w, h = pil_img.size
    if w * h <= MAX_PIXELS:
        return "\n".join(run(pil_img))

    parts = max(2, min(math.ceil(math.sqrt((w * h) / MAX_PIXELS)), 8))
    tw, th = w // parts, h // parts
    ox, oy = int(tw * TILE_OVERLAP), int(th * TILE_OVERLAP)

    seen, chunks = set(), []
    for r in range(parts):
        for c in range(parts):
            x, y = max(0, c * tw - ox), max(0, r * th - oy)
            box = (x, y, min(x + tw + 2 * ox, w), min(y + th + 2 * oy, h))
            for s in run(pil_img.crop(box)):
                s = s.strip()
                if s and s not in seen:
                    seen.add(s)
                    chunks.append(s)
    return "\n".join(chunks)


def _rapid_pdf(path: Path) -> list[dict]:
    """Render each page with PyMuPDF, then OCR the bitmap."""
    import io

    import pymupdf
    from PIL import Image

    pages = []
    with pymupdf.open(path) as doc:
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=pymupdf.Matrix(DPI_SCALE, DPI_SCALE))
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            text = _pil_tiled(img)
            if text.strip():
                pages.append({"page": str(i), "text": text})
    return pages


def _rapid_image(path: Path) -> list[dict]:
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None          # these drawings are legitimately huge
    out = []
    with Image.open(path) as im:
        frames = getattr(im, "n_frames", 1)
        for i in range(frames):            # multi-page TIFF
            im.seek(i)
            text = _pil_tiled(im.convert("RGB"))
            if text.strip():
                out.append({"page": str(i + 1) if frames > 1 else "",
                            "text": text})
    return out


# ------------------------------------------------------------------ dispatch
def ocr_pdf(path: Path) -> list[dict]:
    return _vision_pdf(path) if IS_MAC else _rapid_pdf(path)


def ocr_image(path: Path) -> list[dict]:
    return _vision_image(path) if IS_MAC else _rapid_image(path)


def backend_name() -> str:
    return "macOS Vision" if IS_MAC else "RapidOCR (ONNX)"


def targets() -> list[dict]:
    if not MANIFEST.exists():
        raise SystemExit("Run extract.py first (out/manifest.json missing).")
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    return [r for r in manifest if r["status"] == "needs_ocr"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    rows = targets()
    if args.limit:
        rows = rows[:args.limit]

    if args.list:
        for r in rows:
            print(f"  {r['ext']:<6} {r['bytes']//1024:>7} KB  {r['source_path']}")
        print(f"\n  {len(rows)} files need OCR\n")
        return 0

    OCR_DIR.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    done = skipped = failed = 0
    total_chars = 0

    for i, r in enumerate(rows, 1):
        rel = r["source_path"]
        path = ROOT / rel
        side = sidecar_path(rel)

        if side.exists() and not args.force:
            skipped += 1
            continue
        if not path.exists():
            failed += 1
            continue

        try:
            ext = r["ext"].lower()
            pages = ocr_pdf(path) if ext == ".pdf" else ocr_image(path)
        except Exception as e:
            print(f"  ✗ {rel[:64]} — {type(e).__name__}: {e}", flush=True)
            failed += 1
            continue

        chars = sum(len(p["text"]) for p in pages)
        total_chars += chars
        side.write_text(json.dumps({
            "source_path": rel,
            "filename": r["filename"],
            "pages": pages,
        }, ensure_ascii=False), encoding="utf-8")
        done += 1
        print(f"  [{i:>2}/{len(rows)}] {chars:>7,} chars  {r['filename'][:56]}",
              flush=True)

    print(f"\n  OCR complete — {done} processed, {skipped} already done, "
          f"{failed} failed   [backend: {backend_name()}]")
    print(f"  {total_chars:,} characters recovered in {time.time()-t0:.1f}s")
    print(f"  Sidecars in {OCR_DIR}")
    print("  Re-run extract.py, then index.py, to fold them into the index.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
