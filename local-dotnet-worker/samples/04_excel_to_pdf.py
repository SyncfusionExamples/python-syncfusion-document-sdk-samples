"""Sample 04 — Convert an Excel workbook to PDF (local .NET worker).

If no input workbook is supplied, the sample first asks the worker to build
a tiny `input.xlsx` for it, so the script can be run end-to-end without a
binary fixture. Pass an explicit --input to convert a workbook you already
have on disk.
"""
import argparse
from pathlib import Path

import document_sdk  # type: ignore[import-not-found]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        help="Path to an existing .xlsx file. If omitted, a tiny one is created for you.",
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
        xlsx = out_dir / "input.xlsx"
        service.create_sample_xlsx(xlsx, allow_trial=True)
        print("Generated sample workbook:", xlsx)
    else:
        xlsx = args.input.resolve()
        if not xlsx.is_file():
            raise SystemExit(f"Input workbook not found: {xlsx}")

    pdf = out_dir / "workbook.pdf"
    service.excel_to_pdf(xlsx, pdf, allow_trial=True)
    print("Saved:", pdf)


if __name__ == "__main__":
    main()
