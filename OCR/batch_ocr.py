# from pathlib import Path

# root = Path("C:/Users/pavan/OneDrive/Desktop/LAW/1980-2024/pdfs")
# pdfs = sorted(root.rglob("*.pdf"))
# print(f"Found {len(pdfs)} PDFs")
# for p in pdfs[:20]:
#     print(p)

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional

from pypdf import PdfReader


@dataclass
class OCRRecord:
    input_pdf: str
    output_pdf: str
    sidecar_txt: str
    status: str
    pages: int
    file_size_bytes: int
    elapsed_sec: float
    error: Optional[str] = None


def find_pdfs(root: Path) -> List[Path]:
    return sorted([p for p in root.rglob("*.pdf") if p.is_file()])


def safe_relpath(path: Path, root: Path) -> Path:
    return path.resolve().relative_to(root.resolve())


def count_pages(pdf_path: Path) -> int:
    try:
        reader = PdfReader(str(pdf_path))
        return len(reader.pages)
    except Exception:
        return 0


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def is_ocrmypdf_available() -> bool:
    return shutil.which("ocrmypdf") is not None


def run_ocrmypdf(
    input_pdf: Path,
    output_pdf: Path,
    sidecar_txt: Path,
    language: str,
    jobs: int,
    force_ocr: bool,
    redo_ocr: bool,
    timeout_sec: int,
) -> subprocess.CompletedProcess:
    """
    --skip-text: do not re-OCR born-digital pages
    --deskew: helps scanned court docs
    --optimize 1: safe/default
    --sidecar: save extracted OCR text
    """
    cmd = [
        "ocrmypdf",
        "--deskew",
        "--optimize", "1",
        "--jobs", str(jobs),
        "--sidecar", str(sidecar_txt),
        "--language", language,
    ]

    if force_ocr:
        cmd.append("--force-ocr")
    elif redo_ocr:
        cmd.append("--redo-ocr")
    else:
        cmd.append("--skip-text")

    cmd.extend([str(input_pdf), str(output_pdf)])

    return subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=timeout_sec,
        check=False,
    )


def append_jsonl(path: Path, obj: dict) -> None:
    ensure_parent(path)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")


def process_one(
    input_pdf: Path,
    input_root: Path,
    output_root: Path,
    language: str,
    jobs: int,
    force_ocr: bool,
    redo_ocr: bool,
    timeout_sec: int,
    overwrite: bool,
) -> OCRRecord:
    rel = safe_relpath(input_pdf, input_root)
    output_pdf = output_root / rel
    sidecar_txt = output_pdf.with_suffix(".txt")

    ensure_parent(output_pdf)

    if output_pdf.exists() and not overwrite:
        return OCRRecord(
            input_pdf=str(input_pdf),
            output_pdf=str(output_pdf),
            sidecar_txt=str(sidecar_txt),
            status="skipped_exists",
            pages=count_pages(input_pdf),
            file_size_bytes=input_pdf.stat().st_size,
            elapsed_sec=0.0,
            error=None,
        )

    t0 = time.time()
    pages = count_pages(input_pdf)
    size_bytes = input_pdf.stat().st_size

    try:
        result = run_ocrmypdf(
            input_pdf=input_pdf,
            output_pdf=output_pdf,
            sidecar_txt=sidecar_txt,
            language=language,
            jobs=jobs,
            force_ocr=force_ocr,
            redo_ocr=redo_ocr,
            timeout_sec=timeout_sec,
        )

        elapsed = time.time() - t0

        if result.returncode == 0:
            return OCRRecord(
                input_pdf=str(input_pdf),
                output_pdf=str(output_pdf),
                sidecar_txt=str(sidecar_txt),
                status="ok",
                pages=pages,
                file_size_bytes=size_bytes,
                elapsed_sec=round(elapsed, 3),
                error=None,
            )

        error_msg = (result.stderr or result.stdout or "Unknown OCR error").strip()
        return OCRRecord(
            input_pdf=str(input_pdf),
            output_pdf=str(output_pdf),
            sidecar_txt=str(sidecar_txt),
            status="failed",
            pages=pages,
            file_size_bytes=size_bytes,
            elapsed_sec=round(elapsed, 3),
            error=error_msg[:4000],
        )

    except subprocess.TimeoutExpired:
        elapsed = time.time() - t0
        return OCRRecord(
            input_pdf=str(input_pdf),
            output_pdf=str(output_pdf),
            sidecar_txt=str(sidecar_txt),
            status="timeout",
            pages=pages,
            file_size_bytes=size_bytes,
            elapsed_sec=round(elapsed, 3),
            error=f"Timed out after {timeout_sec} seconds",
        )
    except Exception as e:
        elapsed = time.time() - t0
        return OCRRecord(
            input_pdf=str(input_pdf),
            output_pdf=str(output_pdf),
            sidecar_txt=str(sidecar_txt),
            status="failed",
            pages=pages,
            file_size_bytes=size_bytes,
            elapsed_sec=round(elapsed, 3),
            error=str(e),
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Batch OCR all PDFs in a folder")
    parser.add_argument("--input_dir", required=True, help="Root directory containing PDFs")
    parser.add_argument("--output_dir", required=True, help="Root directory for OCRed PDFs")
    parser.add_argument("--language", default="eng", help="Tesseract language, e.g. eng or eng+hin")
    parser.add_argument("--jobs", type=int, default=max(1, os.cpu_count() or 1), help="Parallel OCR jobs per file")
    parser.add_argument("--timeout_sec", type=int, default=3600, help="Per-file timeout")
    parser.add_argument("--force_ocr", action="store_true", help="OCR every page even if text exists")
    parser.add_argument("--redo_ocr", action="store_true", help="Redo OCR on files that already contain OCR text")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite output PDFs if they already exist")
    args = parser.parse_args()

    input_root = Path(args.input_dir).resolve()
    output_root = Path(args.output_dir).resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    if not input_root.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_root}")

    if not is_ocrmypdf_available():
        raise RuntimeError(
            "ocrmypdf is not available on PATH. Install it first: python -m pip install ocrmypdf"
        )

    pdfs = find_pdfs(input_root)
    if not pdfs:
        print(f"No PDFs found under: {input_root}")
        sys.exit(0)

    manifest_path = output_root / "manifest.jsonl"
    failures_path = output_root / "failures.jsonl"
    summary_path = output_root / "summary.json"

    print(f"Found {len(pdfs)} PDFs")
    print(f"Input : {input_root}")
    print(f"Output: {output_root}")

    counts = {
        "total": len(pdfs),
        "ok": 0,
        "failed": 0,
        "timeout": 0,
        "skipped_exists": 0,
    }

    total_pages = 0
    total_elapsed = 0.0

    for i, pdf in enumerate(pdfs, start=1):
        print(f"[{i}/{len(pdfs)}] OCR -> {pdf}")
        rec = process_one(
            input_pdf=pdf,
            input_root=input_root,
            output_root=output_root,
            language=args.language,
            jobs=args.jobs,
            force_ocr=args.force_ocr,
            redo_ocr=args.redo_ocr,
            timeout_sec=args.timeout_sec,
            overwrite=args.overwrite,
        )

        append_jsonl(manifest_path, asdict(rec))

        counts[rec.status] = counts.get(rec.status, 0) + 1
        total_pages += rec.pages
        total_elapsed += rec.elapsed_sec

        if rec.status in {"failed", "timeout"}:
            append_jsonl(failures_path, asdict(rec))

    summary = {
        "input_dir": str(input_root),
        "output_dir": str(output_root),
        "counts": counts,
        "total_pages": total_pages,
        "total_elapsed_sec": round(total_elapsed, 2),
        "avg_sec_per_pdf": round(total_elapsed / max(1, len(pdfs)), 2),
    }

    with summary_path.open("w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)

    print("\nDone.")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()