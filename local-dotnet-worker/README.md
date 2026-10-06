# Local .NET Worker Sample

Creates a Word document from Python and converts it to PDF by spawning a local, self-contained .NET process (`DocumentBridge.dll`) via `subprocess`. No Python.NET, no pip packages, no web server.

## Layout

```text
local-dotnet-worker/
  document_sdk.py        # Python client (stdlib only)
  NuGet.Config
  DocumentBridge/
    DocumentBridge.csproj    # .NET 8 executable project
    DocumentCreator.cs   # DocIO/XlsIO/Presentation/Pdf logic
    Program.cs           # stdin JSON -> operation dispatcher
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

# Build a Word document and save it as DOCX, then render it straight to PDF.
python document_sdk.py create-docx --trial
python document_sdk.py create-pdf  --trial

# Convert an existing workbook / presentation to PDF, or watermark a PDF.
python document_sdk.py excel-to-pdf      --input documents/input.xlsx  --output output/workbook.pdf --trial
python document_sdk.py powerpoint-to-pdf --input documents/input.pptx  --output output/slides.pdf  --trial
python document_sdk.py watermark-pdf     --input output/workbook.pdf   --output output/watermarked.pdf --label CONFIDENTIAL --trial
```

Replace `win-x64` with `osx-arm64`, `osx-x64`, or `linux-x64` as needed.

> Each operation runs in a fresh `dotnet` subprocess with a default
> per-operation timeout of **300 seconds**. Override with `--timeout <seconds>`
> on any subcommand for large PPTX/XLSX conversions.

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
