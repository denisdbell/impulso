#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OCR via Apple's Vision framework (macOS only) — a drop-in, higher-accuracy
alternative to Tesseract for the DNI photos. Fully LOCAL (the images never
leave the machine), reads HEIC natively, and applies the photo's EXIF
orientation, which phone pictures almost always need.

Used automatically by verify_dni.py when available (see OCR_ENGINE there).

    pip install pyobjc-framework-Vision pyobjc-framework-Quartz
"""

import sys


def available():
    if sys.platform != "darwin":
        return False
    try:
        import Vision, Quartz  # noqa: F401  (import check only)
        return True
    except ImportError:
        return False


def ocr_observations(path):
    """OCR an image file → list of {text, x0, y0, x1, y1} with NORMALIZED
    top-left-origin coordinates (y grows downward), sorted in reading order.
    [] when the file can't be read or no text is recognized. The geometry lets
    callers match a value to the label it sits UNDER — ID cards are laid out in
    columns, so plain line order interleaves unrelated fields."""
    import Vision
    from Foundation import NSURL
    from Quartz import (CGImageSourceCreateWithURL, CGImageSourceCreateImageAtIndex,
                        CGImageSourceCopyPropertiesAtIndex)

    url = NSURL.fileURLWithPath_(str(path))
    src = CGImageSourceCreateWithURL(url, None)
    if src is None:
        return []
    img = CGImageSourceCreateImageAtIndex(src, 0, None)
    if img is None:
        return []
    props = CGImageSourceCopyPropertiesAtIndex(src, 0, None) or {}
    try:
        orientation = int(props.get("Orientation") or 1)   # EXIF 1-8 == CGImagePropertyOrientation
    except (TypeError, ValueError):
        orientation = 1

    req = Vision.VNRecognizeTextRequest.alloc().init()
    req.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    req.setUsesLanguageCorrection_(True)
    try:
        req.setRecognitionLanguages_(["es-AR", "es", "en-US"])
    except Exception:   # noqa: BLE001 — older macOS without es support: keep defaults
        pass

    handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_orientation_options_(
        img, orientation, None)
    ok = handler.performRequests_error_([req], None)
    if isinstance(ok, tuple):       # pyobjc returns (bool, error)
        ok = ok[0]
    if not ok:
        return []

    out = []
    for obs in (req.results() or []):
        cands = obs.topCandidates_(1)
        if not cands or not len(cands):
            continue
        bb = obs.boundingBox()      # normalized, origin BOTTOM-left
        x0, w = bb.origin.x, bb.size.width
        y_bottom, h = bb.origin.y, bb.size.height
        out.append({"text": str(cands[0].string()),
                    "x0": x0, "x1": x0 + w,
                    "y0": 1.0 - (y_bottom + h), "y1": 1.0 - y_bottom})
    out.sort(key=lambda o: (o["y0"], o["x0"]))
    return out


def ocr_image_text(path):
    """OCR an image file → text, lines in reading order (top→bottom, left→right).
    Returns '' when the file can't be read or no text is recognized."""
    return "\n".join(o["text"] for o in ocr_observations(path))


if __name__ == "__main__":
    for p in sys.argv[1:]:
        print("=" * 20, p)
        print(ocr_image_text(p))
