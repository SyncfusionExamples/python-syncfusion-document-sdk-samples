"""Create Word, Excel, and PowerPoint documents and convert them to PDF using a
local .NET worker.

This module launches the published DocumentBridge.dll as a separate process via
Python's built-in `subprocess` module. No pip packages, Python.NET, or a web
server are required. Works with plain Python 3.9+ on any platform that has
the .NET 8 runtime installed.

By default every operation runs in Syncfusion's **trial mode** (no license
required). To use a real license, set the ``SYNCFUSION_LICENSE_KEY`` environment
variable before invoking the worker; the worker picks it up automatically and
Python callers do not need to change.
"""
import argparse
import json
import shutil
import subprocess
from pathlib import Path


class DocumentService:
    """Thin client that starts the .NET worker once per operation."""

    def __init__(self, dll: Path, dotnet: str, timeout: int = 300):
        self._dll = dll
        self._dotnet = dotnet
        # Default is generous because large PPTX/XLSX -> PDF conversions can
        # easily exceed a minute. Callers can override per-instance.
        self._timeout = timeout
        # The .NET worker reads SYNCFUSION_LICENSE_KEY from the environment
        # on first use. If unset, it runs in Syncfusion's trial mode
        # (evaluation watermarks are added to the output).

    def _call(self, operation: str, output: Path, *,
              text: str | None = None, input: Path | None = None,
              label: str | None = None) -> None:
        # The .NET worker reads SYNCFUSION_LICENSE_KEY from the environment,
        # so we do not need to forward a trial flag per call.
        request = {
            "Operation": operation,
            "Text": text,
            "Input": str(input.resolve()) if input is not None else None,
            "Output": str(output.resolve()),
            "Label": label,
        }
        # No shell=True: the request is passed as JSON data on stdin, never
        # interpreted as a command line, so paths/text cannot inject commands.
        try:
            result = subprocess.run(
                [self._dotnet, str(self._dll)],
                input=json.dumps(request),
                text=True,
                encoding="utf-8",
                capture_output=True,
                check=False,
                timeout=self._timeout,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                f"Worker timed out after {self._timeout}s on operation '{operation}'."
            ) from exc
        if result.returncode != 0:
            detail = result.stderr.strip() or f"Worker exited with code {result.returncode}"
            raise RuntimeError(detail)

    # ---- Word: build from a string ----------------------------------------------------

    def create_docx(self, text: str, output: Path) -> None:
        """Create a new Word document and save it as DOCX."""
        self._call("create-docx", output, text=text)

    def create_pdf(self, text: str, output: Path) -> None:
        """Create a new Word document in memory and render it directly to PDF."""
        self._call("create-pdf", output, text=text)

    # ---- Excel / PowerPoint: convert an existing file to PDF -------------------------

    def excel_to_pdf(self, source: Path, output: Path) -> None:
        """Convert an XLSX workbook to PDF using its own print settings."""
        self._call("excel-to-pdf", output, input=source)

    def powerpoint_to_pdf(self, source: Path, output: Path) -> None:
        """Convert a PPTX presentation to PDF via Syncfusion's presentation renderer."""
        self._call("powerpoint-to-pdf", output, input=source)

    # ---- PDF: overlay a diagonal text watermark on every page -------------------------

    def watermark_pdf(self, source: Path, output: Path, label: str) -> None:
        """Draw a translucent diagonal label on every page of an existing PDF."""
        self._call("watermark-pdf", output, input=source, label=label)

    # ---- Sample input helpers --------------------------------------------------------

    def create_sample_xlsx(self, output: Path) -> None:
        """Build a tiny XLSX so the samples run end-to-end without a binary fixture."""
        self._call("create-sample-xlsx", output)

    def create_sample_pptx(self, output: Path) -> None:
        """Build a tiny PPTX so the samples run end-to-end without a binary fixture."""
        self._call("create-sample-pptx", output)


def load_service(bundle: Path | None = None) -> DocumentService:
    """Locate the published worker and the dotnet executable.

    The worker inherits ``SYNCFUSION_LICENSE_KEY`` from the current environment
    if set; the key is never passed as a command-line argument.
    """
    bundle = Path(bundle or Path(__file__).parent / "artifacts").resolve()
    dll = bundle / "DocumentBridge.dll"
    config = bundle / "DocumentBridge.runtimeconfig.json"
    if not dll.is_file() or not config.is_file():
        raise FileNotFoundError(f"Publish DocumentBridge into {bundle} first (see README).")

    dotnet = shutil.which("dotnet")
    if dotnet is None:
        raise RuntimeError("Install the .NET 8 runtime/SDK and put dotnet on PATH.")

    return DocumentService(dll, dotnet)


# Each CLI subcommand maps to one DocumentService method. Tuple shape:
# (method-name, input-extension-or-None, output-extension, needs-label).
_OPERATIONS: dict[str, tuple[str, str | None, str, bool]] = {
    "create-docx":       ("create_docx",       None, ".docx", False),
    "create-pdf":        ("create_pdf",        None, ".pdf",  False),
    "excel-to-pdf":      ("excel_to_pdf",      ".xlsx", ".pdf", False),
    "powerpoint-to-pdf": ("powerpoint_to_pdf", ".pptx", ".pdf", False),
    "watermark-pdf":     ("watermark_pdf",     ".pdf",  ".pdf", True),
    "create-sample-xlsx":("create_sample_xlsx",None,   ".xlsx", False),
    "create-sample-pptx":("create_sample_pptx",None,   ".pptx", False),
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
    parser.add_argument("--bundle", type=Path, help="Folder containing the published worker.")
    parser.add_argument("--timeout", type=int, default=300,
                        help="Per-operation subprocess timeout in seconds (default: 300).")
    args = parser.parse_args()

    method_name, input_ext, output_ext, needs_label = _OPERATIONS[args.operation]
    output = (args.output or Path(__file__).parent / "output" / f"Sample{output_ext}").resolve()
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
