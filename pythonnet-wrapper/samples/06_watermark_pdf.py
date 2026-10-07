"""Sample 06 — Add a diagonal text watermark to every page of a PDF
(Python.NET wrapper).

If no input PDF is supplied, the sample first asks the worker to render a
tiny `Hello.pdf` and then watermarks that. Pass an explicit --input to
watermark a PDF you already have on disk.
"""
import argparse
import re
from pathlib import Path

import document_sdk  # type: ignore[import-not-found]


def _slug_label(label: str) -> str:
    """Make a label safe to use as part of a file name.

    Strips path separators and any character that is not alphanumeric,
    dash, underscore, or dot. Falls back to ``"label"`` if the result is
    empty.
    """
    slug = re.sub(r"[^A-Za-z0-9._-]+", "-", label).strip("-")
    return slug or "label"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        help="Path to an existing .pdf file. If omitted, a tiny one is created for you.",
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "output",
    )
    parser.add_argument("--label", default="CONFIDENTIAL", help="Watermark text.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    out_dir: Path = args.out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    service = document_sdk.load_service()

    if args.input is None:
        source = out_dir / "Hello.pdf"
        # Match the wording used by sample 01 for the pythonnet track.
        service.create_pdf("Hello from Python.NET and DocIO!", source)
        print("Generated sample PDF:", source)
    else:
        source = args.input.resolve()
        if not source.is_file():
            raise SystemExit(f"Input PDF not found: {source}")

    watermarked = out_dir / f"watermarked-{_slug_label(args.label)}.pdf"
    service.watermark_pdf(source, watermarked, args.label)
    print("Saved:", watermarked)


if __name__ == "__main__":
    main()
