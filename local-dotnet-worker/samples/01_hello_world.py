"""Sample 01 — Hello World (local .NET worker).

The smallest possible usage: import the client, call it twice, print the
output paths. This is the recommended first run to confirm the worker is
published correctly.
"""
from pathlib import Path

import document_sdk  # type: ignore[import-not-found]


def main() -> None:
    service = document_sdk.load_service()
    out_dir = (Path(__file__).resolve().parent.parent / "output").resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    docx_path = out_dir / "Hello.docx"
    pdf_path = out_dir / "Hello.pdf"

    service.create_docx("Hello from the local .NET worker!", docx_path)
    service.create_pdf("Hello from the local .NET worker!", pdf_path)

    print("Saved:", docx_path)
    print("Saved:", pdf_path)


if __name__ == "__main__":
    main()
