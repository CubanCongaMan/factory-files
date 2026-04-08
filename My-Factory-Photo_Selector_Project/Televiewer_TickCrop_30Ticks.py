#!/usr/bin/env python3
"""
Crop a long televiewer image into segments that each contain at least 30 ruler
 ticks (combined left + inner boundary ruler), equivalent to ~3.0 ft.

Output is hardcoded to ./cropping_test
Only required input: source image path (CLI arg or one prompt).
"""

import os
import re
import sys
from pathlib import Path
from statistics import median

from PIL import Image

# Large televiewer images can exceed PIL's decompression-bomb warning threshold.
Image.MAX_IMAGE_PIXELS = None

RIGHT_CROP_X = 685
LEFT_STRIP_W = 35
TICK_DARK_THRESHOLD = 70
TICK_MIN_DARK_PIXELS = 6
TICK_PEAK_MIN_SEP = 8
TARGET_TICKS = 30
TICKS_PER_FOOT_PER_SIDE = 10  # user rule: 0.1 ft spacing on left ruler


def parse_input_image_path(argv):
    if len(argv) > 1:
        return Path(" ".join(argv[1:]).strip().strip('"'))
    user = input("Enter full path to long televiewer image: ").strip().strip('"')
    return Path(user)


def detect_tick_peaks(gray_img, x_start, x_end, y_start=0, y_end=None):
    pix = gray_img.load()
    width, height = gray_img.size
    x0 = max(0, min(width - 1, x_start))
    x1 = max(x0 + 1, min(width, x_end))
    ys = max(0, min(height - 1, y_start))
    ye = height if y_end is None else max(ys + 1, min(height, y_end))

    scores = []
    for y in range(ys, ye):
        dark = 0
        for x in range(x0, x1):
            if pix[x, y] < TICK_DARK_THRESHOLD:
                dark += 1
        scores.append(dark)

    peaks = []
    for i in range(1, len(scores) - 1):
        if (
            scores[i] >= TICK_MIN_DARK_PIXELS
            and scores[i] >= scores[i - 1]
            and scores[i] >= scores[i + 1]
        ):
            y = ys + i
            if not peaks or (y - peaks[-1]) > TICK_PEAK_MIN_SEP:
                peaks.append(y)
    return peaks


def find_ruler_boundary_x(gray_img):
    width, height = gray_img.size
    pix = gray_img.load()
    y0 = max(0, height // 4)
    y1 = min(height, (3 * height) // 4)
    if y1 <= y0:
        return width // 2

    means = []
    for x in range(width):
        s = 0
        n = 0
        for y in range(y0, y1):
            s += pix[x, y]
            n += 1
        means.append((s / n) if n else 255)

    for x in range(10, width - 10):
        if means[x - 1] > 220 and means[x] < 200:
            return x
    return width // 2


def count_left_ticks(gray_crop):
    """Count tick marks on the left ruler only.

    User criterion: 30 ticks on the left ruler edge per crop.
    """
    w, h = gray_crop.size
    left = detect_tick_peaks(gray_crop, 0, min(LEFT_STRIP_W, w), y_start=0, y_end=h)
    return len(left)


def estimate_tick_spacing(peaks):
    if len(peaks) < 3:
        return None
    diffs = [peaks[i] - peaks[i - 1] for i in range(1, len(peaks))]
    plausible = [d for d in diffs if 8 <= d <= 120]
    if not plausible:
        return None
    return int(round(median(plausible)))


def main(argv):
    source_path = parse_input_image_path(argv)
    if not source_path.exists():
        print(f"ERROR: Image not found: {source_path}")
        return 1

    script_dir = Path(__file__).resolve().parent
    output_dir = script_dir / "cropping_test"
    output_dir.mkdir(exist_ok=True)

    # Clear old test crops in destination.
    for old in output_dir.glob("*.jpg"):
        try:
            old.unlink()
        except OSError:
            pass

    img = Image.open(source_path)
    width, height = img.size
    crop_width = min(RIGHT_CROP_X, width)
    base = re.sub(r"[^A-Za-z0-9._-]", "_", source_path.stem)

    working = img.crop((0, 0, crop_width, height))
    gray = working.convert("L")

    boundary_x = find_ruler_boundary_x(gray)
    left_peaks = detect_tick_peaks(gray, 0, min(LEFT_STRIP_W, crop_width))
    right_peaks = detect_tick_peaks(gray, max(0, boundary_x - LEFT_STRIP_W), min(crop_width, boundary_x + 1))

    spacing = estimate_tick_spacing(left_peaks) or estimate_tick_spacing(right_peaks)
    if spacing is None:
        print("ERROR: Could not detect ruler tick spacing from source image.")
        return 2

    # User rule: one left-side tick = 0.1 ft, so 3.0 ft = 30 ticks.
    target_ticks_one_side = int(3.0 * TICKS_PER_FOOT_PER_SIDE)
    window_height = int(round(spacing * target_ticks_one_side))
    step = max(1, spacing // 2)
    max_expand = spacing * 10

    first_peak = min(left_peaks[0], right_peaks[0]) if (left_peaks and right_peaks) else (left_peaks[0] if left_peaks else (right_peaks[0] if right_peaks else 0))
    start = max(0, first_peak - spacing)

    saved = 0
    scan_guard = 0
    while (start + window_height) <= height and scan_guard < 2000:
        scan_guard += 1
        end = min(height, start + window_height)
        gray_crop = gray.crop((0, start, crop_width, end))
        ticks = count_left_ticks(gray_crop)

        expanded = 0
        while ticks < TARGET_TICKS and end < height and expanded < max_expand:
            end = min(height, end + step)
            expanded += step
            gray_crop = gray.crop((0, start, crop_width, end))
            ticks = count_left_ticks(gray_crop)

        if ticks >= TARGET_TICKS:
            seg = working.crop((0, start, crop_width, end))
            out_name = f"{base}_seg_{saved + 1:03d}_{ticks}leftticks.jpg"
            out_path = output_dir / out_name
            seg.save(out_path, "JPEG", quality=95)
            saved += 1
            start = end
        else:
            # Move down to seek next valid 30-tick window.
            start += step

    print(f"Source: {source_path}")
    print(f"Output folder: {output_dir}")
    print(f"Detected tick spacing: {spacing} px")
    print(f"Estimated 3-ft window height: {window_height} px")
    print(f"Segments written: {saved}")
    print(f"Tick criterion: >= {TARGET_TICKS} ticks on LEFT ruler")

    if saved == 0:
        print("WARNING: No 30-tick crops were produced.")
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
