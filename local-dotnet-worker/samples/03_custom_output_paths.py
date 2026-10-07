"""Sample 03 — Custom output paths and bundle location (local .NET worker).

Demonstrates two things that are common in real applications:
  1. Writing the DOCX and PDF into user-chosen folders instead of ./output.
  2. Pointing the loader at a `dotnet publish` folder that lives outside
     the default `artifacts/` next to `document_sdk.py`.
"""
import argparse
from pathlib import Path

import document_sdk  # type: ignore[import-not-found]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "output" / "custom",
    )
    parser.add_argument(
        "--bundle",
        type=Path,
        help="Folder containing the published DocumentBridge (defaults to ../artifacts).",
    )
    parser.add_argument("--text", default="Generated with custom paths.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    out_dir: Path = args.out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    service = document_sdk.load_service(args.bundle)

    docx_path = out_dir / "Report.docx"
    pdf_path = out_dir / "Report.pdf"
    service.create_docx(args.text, docx_path)
    service.create_pdf(args.text, pdf_path)

    print("Saved DOCX:", docx_path)
    print("Saved PDF: ", pdf_path)


if __name__ == "__main__":
    main()
