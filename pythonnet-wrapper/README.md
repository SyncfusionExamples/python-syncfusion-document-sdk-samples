# Python.NET Wrapper Sample

Creates a Word document from Python and converts it to PDF by loading the Syncfusion-powered .NET assembly directly **inside** the Python process using [pythonnet](https://pythonnet.github.io/).

## Layout

```text
pythonnet-wrapper/
  document_sdk.py        # Python client (uses pythonnet/clr) — exposes DocumentService
  requirements.txt       # pythonnet>=3.1.0
  NuGet.Config
  DocumentBridge/
    DocumentBridge.csproj    # .NET 8 class library
    DocumentCreator.cs   # DocIO/XlsIO/Presentation/Pdf logic, called directly from Python
  samples/               # Runnable end-to-end examples
    README.md
    01_hello_world.py
    02_batch_invoices.py
    03_custom_output_paths.py
    04_excel_to_pdf.py
    05_powerpoint_to_pdf.py
    06_watermark_pdf.py
```

## Run

```bash
dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts

python -m venv .venv
.venv\Scripts\Activate.ps1   # On Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt

# Build a Word document and save it as DOCX, then render it straight to PDF.
python document_sdk.py create-docx
python document_sdk.py create-pdf

# Convert an existing workbook / presentation to PDF, or watermark a PDF.
python document_sdk.py excel-to-pdf      --input documents/input.xlsx  --output output/workbook.pdf
python document_sdk.py powerpoint-to-pdf --input documents/input.pptx  --output output/slides.pdf
python document_sdk.py watermark-pdf     --input output/workbook.pdf   --output output/watermarked.pdf --label CONFIDENTIAL
```

Replace `win-x64` with the RID that matches your platform (`osx-arm64`,
`osx-x64`, or `linux-x64`). When cross-publishing from Windows, pick the
RID that matches the **target** machine, not the host.

> Every command runs in Syncfusion's **trial mode** by default — no license
> is required, and evaluation watermarks are added to the output documents.
> To produce watermark-free output, set the `SYNCFUSION_LICENSE_KEY`
> environment variable before running the command:
>
> ```bash
> # Windows PowerShell
> $env:SYNCFUSION_LICENSE_KEY = "your-key"
> # bash / zsh
> SYNCFUSION_LICENSE_KEY=your-key python document_sdk.py create-pdf
> ```
>
> The C# worker reads the variable on first use; Python callers do not need
> to pass any flag per call.

## Samples

See [`samples/README.md`](samples/README.md) for a small set of runnable
scripts that build on `document_sdk.load_service()`:

| # | Sample | Demonstrates |
|---|---|---|
| 1 | `01_hello_world.py` | Minimal end-to-end DOCX + PDF generation. |
| 2 | `02_batch_invoices.py` | Loop over records and emit one DOCX each. |
| 3 | `03_custom_output_paths.py` | Custom output folder and non-default bundle path. |
| 4 | `04_excel_to_pdf.py` | Convert an XLSX workbook to PDF. |
| 5 | `05_powerpoint_to_pdf.py` | Convert a PPTX presentation to PDF. |
| 6 | `06_watermark_pdf.py` | Overlay a diagonal text watermark on every page of a PDF. |
