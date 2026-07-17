from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
from collections import deque
from pathlib import Path

import numpy as np
from PIL import Image, ImageFilter


def render_pdf(pdf: Path, output_dir: Path, dpi: int) -> list[Path]:
    pdftoppm = shutil.which("pdftoppm.exe") or shutil.which("pdftoppm")
    if not pdftoppm:
        # Claude fork: fall back to the Poppler bundled with the Codex runtime on this PC.
        bundled = (
            Path.home()
            / ".cache" / "codex-runtimes" / "codex-primary-runtime"
            / "dependencies" / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe"
        )
        if bundled.exists():
            pdftoppm = str(bundled)
        else:
            raise RuntimeError("pdftoppm was not found on PATH (install Poppler or add it to PATH)")
    pdftoppm_path = Path(pdftoppm)
    if pdftoppm_path.suffix.casefold() == ".cmd":
        bundled_exe = (
            pdftoppm_path.parents[2]
            / "native"
            / "poppler"
            / "Library"
            / "bin"
            / "pdftoppm.exe"
        )
        if bundled_exe.exists():
            pdftoppm = str(bundled_exe)
    output_dir.mkdir(parents=True, exist_ok=True)
    prefix = output_dir / "page"
    subprocess.run(
        [pdftoppm, "-r", str(dpi), "-png", str(pdf), str(prefix)],
        check=True,
        capture_output=True,
        text=True,
    )
    return sorted(output_dir.glob("page-*.png"))


def _components(mask: np.ndarray) -> list[tuple[int, int, int, int]]:
    height, width = mask.shape
    seen = np.zeros_like(mask, dtype=bool)
    boxes: list[tuple[int, int, int, int]] = []
    for y in range(height):
        for x in range(width):
            if not mask[y, x] or seen[y, x]:
                continue
            queue = deque([(y, x)])
            seen[y, x] = True
            min_x = max_x = x
            min_y = max_y = y
            while queue:
                cy, cx = queue.popleft()
                min_x, max_x = min(min_x, cx), max(max_x, cx)
                min_y, max_y = min(min_y, cy), max(max_y, cy)
                for ny in range(max(0, cy - 1), min(height, cy + 2)):
                    for nx in range(max(0, cx - 1), min(width, cx + 2)):
                        if mask[ny, nx] and not seen[ny, nx]:
                            seen[ny, nx] = True
                            queue.append((ny, nx))
            boxes.append((min_x, min_y, max_x + 1, max_y + 1))
    return boxes


def changed_regions(
    current: Image.Image,
    baseline: Image.Image,
    ink_threshold: int = 235,
    block: int = 16,
    margin: int = 48,
) -> list[tuple[int, int, int, int]]:
    # OneNote can publish the same A4 page a few pixels wider on another PC
    # because of device-specific PDF metrics. Normalize small size differences
    # before comparing so a Desktop baseline also works on Lenovo.
    width_delta = abs(current.width - baseline.width) / max(current.width, baseline.width)
    height_delta = abs(current.height - baseline.height) / max(current.height, baseline.height)
    if width_delta <= 0.01 and height_delta <= 0.01 and current.size != baseline.size:
        baseline = baseline.resize(current.size, Image.Resampling.LANCZOS)

    width = max(current.width, baseline.width)
    height = max(current.height, baseline.height)
    current_canvas = Image.new("L", (width, height), 255)
    baseline_canvas = Image.new("L", (width, height), 255)
    current_canvas.paste(current.convert("L"), (0, 0))
    baseline_canvas.paste(baseline.convert("L"), (0, 0))
    cur_ink = np.asarray(current_canvas, dtype=np.uint8) < ink_threshold
    old_ink = np.asarray(baseline_canvas, dtype=np.uint8) < ink_threshold

    # Ignore small antialiasing and stroke-position differences introduced by
    # separate PDF renderings. A real added/removed stroke still has pixels
    # outside the three-pixel neighborhood of the other image.
    def dilate(mask: np.ndarray) -> np.ndarray:
        image = Image.fromarray(mask.astype(np.uint8) * 255)
        return np.asarray(image.filter(ImageFilter.MaxFilter(7))) > 0

    changed = np.logical_or(
        np.logical_and(cur_ink, np.logical_not(dilate(old_ink))),
        np.logical_and(old_ink, np.logical_not(dilate(cur_ink))),
    )

    pad_y = (-height) % block
    pad_x = (-width) % block
    changed = np.pad(changed, ((0, pad_y), (0, pad_x)), constant_values=False)
    pooled = changed.reshape(
        changed.shape[0] // block, block, changed.shape[1] // block, block
    ).sum(axis=(1, 3))
    active = pooled >= 3

    # Join neighboring strokes belonging to one handwritten addition.
    for _ in range(2):
        padded = np.pad(active, 1, constant_values=False)
        active = np.logical_or.reduce(
            [padded[dy : dy + active.shape[0], dx : dx + active.shape[1]]
             for dy in range(3) for dx in range(3)]
        )

    regions = []
    for left, top, right, bottom in _components(active):
        x0 = max(0, left * block - margin)
        y0 = max(0, top * block - margin)
        x1 = min(width, right * block + margin)
        y1 = min(height, bottom * block + margin)
        if (x1 - x0) * (y1 - y0) >= 2_000:
            regions.append((x0, y0, x1, y1))
    return regions


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Render a OneNote PDF and extract regions changed since the last committed baseline."
    )
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--state-dir", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--commit-baseline", action="store_true")
    args = parser.parse_args()

    args.pdf = args.pdf.resolve()
    args.state_dir = args.state_dir.resolve()
    args.output_dir = args.output_dir.resolve()
    baseline_dir = args.state_dir / "baseline"
    args.output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="onenote-render-") as temp:
        temp_dir = Path(temp)
        # The bundled Windows Poppler wrapper cannot reliably open paths that
        # contain Japanese characters, so render an ASCII-named staging copy.
        staged_pdf = temp_dir / "source.pdf"
        shutil.copy2(args.pdf, staged_pdf)
        pages = render_pdf(staged_pdf, temp_dir / "rendered", args.dpi)
        if not pages:
            raise RuntimeError("The PDF rendered no pages")

        manifest: dict[str, object] = {
            "pdf": str(args.pdf),
            "dpi": args.dpi,
            "baseline_exists": baseline_dir.exists(),
            "pages": [],
        }
        total_regions = 0
        for page_number, current_path in enumerate(pages, start=1):
            current = Image.open(current_path).convert("RGB")
            baseline_path = baseline_dir / f"page-{page_number:03d}.png"
            if baseline_path.exists():
                baseline = Image.open(baseline_path).convert("RGB")
                regions = changed_regions(current, baseline)
            else:
                regions = [(0, 0, current.width, current.height)]

            page_entry = {"page": page_number, "regions": []}
            for region_number, box in enumerate(regions, start=1):
                crop_name = f"page-{page_number:03d}-region-{region_number:02d}.png"
                current.crop(box).save(args.output_dir / crop_name)
                page_entry["regions"].append({"box": list(box), "image": crop_name})
                total_regions += 1
            manifest["pages"].append(page_entry)

        manifest["changed_region_count"] = total_regions
        manifest_path = args.output_dir / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

        if args.commit_baseline:
            replacement = args.state_dir / "baseline.new"
            if replacement.exists():
                shutil.rmtree(replacement)
            replacement.mkdir(parents=True, exist_ok=True)
            for page_number, current_path in enumerate(pages, start=1):
                shutil.copy2(current_path, replacement / f"page-{page_number:03d}.png")
            if baseline_dir.exists():
                shutil.rmtree(baseline_dir)
            replacement.replace(baseline_dir)

    print(f"ChangedRegionCount={total_regions}")
    print(f"Manifest={manifest_path}")
    print(f"BaselineCommitted={str(args.commit_baseline).lower()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
