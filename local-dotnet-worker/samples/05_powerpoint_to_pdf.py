"""Sample 05 — Convert a PowerPoint presentation to PDF (local .NET worker).

If no input presentation is supplied, the sample first asks the worker to
build a tiny `input.pptx` for it, so the script can be run end-to-end without
a binary fixture. Pass an explicit --input to convert a presentation you
already have on disk.
"""
import argparse
from pathlib import Path

import document_sdk  # type: ignore[import-not-found]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        help="Path to an existing .pptx file. If omitted, a tiny one is created for you.",
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "output",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    out_dir: Path = args.out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    service = document_sdk.load_service()

    if args.input is None:
        pptx = out_dir / "input.pptx"
        service.create_sample_pptx(pptx)
        print("Generated sample presentation:", pptx)
    else:
        pptx = args.input.resolve()
        if not pptx.is_file():
            raise SystemExit(f"Input presentation not found: {pptx}")

    pdf = out_dir / "slides.pdf"
    service.powerpoint_to_pdf(pptx, pdf)
    print("Saved:", pdf)


if __name__ == "__main__":
    main()
