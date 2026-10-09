"""Create Word, Excel, and PowerPoint documents and convert them to PDF using
Python.NET (pythonnet).

This module uses Python.NET to host CoreCLR inside the Python process and load
the published DocumentBridge.dll directly. Python can then call DocumentCreator's
methods natively without spawning separate processes.

The public API mirrors the local .NET worker sample (``DocumentService``) so
that the same caller code can be used with either approach.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _load_clr_bridge(bundle: Path) -> type:
    """Initialize CoreCLR with DocumentBridge's runtime configuration, then import DocumentCreator.

    Raises ``FileNotFoundError`` if the worker has not been published yet, and
    re-raises ``ModuleNotFoundError`` if ``pythonnet`` is not installed in the
    active environment.
    """
    bundle = Path(bundle).resolve()
    dll = bundle / "DocumentBridge.dll"
    config = bundle / "DocumentBridge.runtimeconfig.json"
    if not dll.is_file() or not config.is_file():
        raise FileNotFoundError(
            f"Publish DocumentBridge into '{bundle}' first (see README.md)."
        )

    # CoreCLR MUST be configured before importing `clr`.
    from pythonnet import load
    load("coreclr", runtime_config=str(config))

    # Add the assembly folder to sys.path so Python.NET's assembly resolver finds dependencies.
    sys.path.insert(0, str(bundle))
    import clr  # noqa: E402
    clr.AddReference(str(dll))

    from DocumentInterop import DocumentCreator  # type: ignore  # noqa: E402
    return DocumentCreator


class DocumentService:
    """Thin client that hosts CoreCLR once and dispatches calls to DocumentCreator.

    The shape of this class deliberately matches ``local-dotnet-worker.DocumentService``
    so callers can swap between the two approaches without changing their code.
    """

    def __init__(self, creator: type):
        self._creator = creator

    # ---- Word: build from a string -------------------------------------------------

    def create_docx(self, text: str, output: Path) -> None:
        """Create a new Word document and save it as DOCX."""
        self._creator.CreateDocx(str(text), str(output))

    def create_pdf(self, text: str, output: Path) -> None:
        """Create a new Word document in memory and render it directly to PDF."""
        self._creator.CreatePdf(str(text), str(output))

    # ---- Excel / PowerPoint: convert an existing file to PDF -----------------------

    def excel_to_pdf(self, source: Path, output: Path) -> None:
        """Convert an XLSX workbook to PDF using its own print settings."""
        self._creator.ExcelToPdf(str(source), str(output))

    def powerpoint_to_pdf(self, source: Path, output: Path) -> None:
        """Convert a PPTX presentation to PDF via Syncfusion's presentation renderer."""
        self._creator.PowerPointToPdf(str(source), str(output))

    # ---- PDF: overlay a diagonal text watermark on every page ----------------------

    def watermark_pdf(self, source: Path, output: Path, label: str) -> None:
        """Draw a translucent diagonal label on every page of an existing PDF."""
        self._creator.WatermarkPdf(str(source), str(output), str(label))

    # ---- Sample input helpers -------------------------------------------------------

    def create_sample_xlsx(self, output: Path) -> None:
        """Build a tiny XLSX so the samples run end-to-end without a binary fixture."""
        self._creator.CreateSampleXlsx(str(output))

    def create_sample_pptx(self, output: Path) -> None:
        """Build a tiny PPTX so the samples run end-to-end without a binary fixture."""
        self._creator.CreateSamplePptx(str(output))


def load_service(bundle: Path | None = None) -> DocumentService:
    """Load DocumentBridge into the current Python process and return a ready client.

    The first call initializes CoreCLR. A second call with a different
    ``runtime_config`` is not supported by Python.NET — restart the process.
    """
    bundle = Path(bundle or Path(__file__).parent / "artifacts").resolve()
    creator_cls = _load_clr_bridge(bundle)
    creator = creator_cls()
    return DocumentService(creator)


# Each CLI subcommand maps to one DocumentService method. Tuple shape:
# (method-name, input-extension-or-None, output-extension, needs-label).
_OPERATIONS: dict[str, tuple[str, str | None, str, bool]] = {
    "create-docx":        ("create_docx",        None,   ".docx", False),
    "create-pdf":         ("create_pdf",         None,   ".pdf",  False),
    "excel-to-pdf":       ("excel_to_pdf",       ".xlsx", ".pdf", False),
    "powerpoint-to-pdf":  ("powerpoint_to_pdf",  ".pptx", ".pdf", False),
    "watermark-pdf":      ("watermark_pdf",      ".pdf",  ".pdf", True),
    "create-sample-xlsx": ("create_sample_xlsx", None,   ".xlsx", False),
    "create-sample-pptx": ("create_sample_pptx", None,   ".pptx", False),
}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=list(_OPERATIONS))
    parser.add_argument("--input", type=Path, help="Input file (for convert/watermark operations).")
    parser.add_argument("--output", type=Path, help="Output file (default: ./output/Sample.<ext>).")
    parser.add_argument("--text", default="Hello from Python and Syncfusion DocIO!",
                        help="Text for create-docx / create-pdf.")
    parser.add_argument("--label", default="CONFIDENTIAL",
                        help="Label for watermark-pdf.")
    parser.add_argument("--bundle", type=Path, help="Folder containing the published DocumentBridge.")
    args = parser.parse_args()

    method_name, input_ext, output_ext, needs_label = _OPERATIONS[args.operation]
    folder = Path(__file__).resolve().parent
    output = (args.output or folder / "output" / f"Sample{output_ext}").resolve()
    if output.suffix.lower() != output_ext:
        parser.error(f"--output must have extension {output_ext}.")
    if input_ext is not None:
        if args.input is None:
            parser.error(f"--input is required for {args.operation}.")
        source = args.input.resolve()
        if source.suffix.lower() != input_ext:
            parser.error(f"--input must have extension {input_ext}.")
        if not source.is_file():
            parser.error(f"Input file does not exist: {source}")

    service = load_service(args.bundle)
    kwargs: dict = {"output": output}
    if input_ext is not None:
        kwargs["source"] = source
    if method_name in ("create_docx", "create_pdf"):
        kwargs["text"] = args.text
    if needs_label:
        kwargs["label"] = args.label
    getattr(service, method_name)(**kwargs)
    print("Saved:", output)


if __name__ == "__main__":
    main()

