# -*- coding: utf-8 -*-
"""PDF の各ページを PNG に書き出す。

このPCには poppler(pdftoppm) が PATH 上に無いため、pdf2image は使わない。
venv に入っている pypdfium2 で描画する（外部実行ファイル不要）。

    python tools\\render_pdf_to_images.py <input.pdf> <out_dir> [--dpi 300]

書き出したファイルのパスを1行ずつ標準出力に出す。
"""
import argparse
import sys
from pathlib import Path

import pypdfium2 as pdfium


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("out_dir")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--max-pages", type=int, default=40)
    args = ap.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.is_file():
        print(f"ERROR: PDF が見つかりません: {pdf_path}", file=sys.stderr)
        return 2

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    doc = pdfium.PdfDocument(str(pdf_path))
    n = len(doc)
    if n > args.max_pages:
        print(f"WARNING: {n} ページあります。先頭 {args.max_pages} ページのみ描画します。",
              file=sys.stderr)
        n = args.max_pages

    written = []
    for i in range(n):
        img = doc[i].render(scale=args.dpi / 72).to_pil()
        # 読み取りやすさのためグレースケール化（手書きは色情報をほぼ使わない）
        if img.mode not in ("L", "RGB"):
            img = img.convert("RGB")
        out = out_dir / f"page_{i + 1:02d}.png"
        img.save(out, format="PNG", optimize=True)
        written.append(out)

    print(f"PageCount={len(doc)}")
    print(f"Rendered={len(written)}")
    print(f"DPI={args.dpi}")
    for p in written:
        print(f"Image={p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
