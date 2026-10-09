# Samples — Local .NET Worker

Runnable examples that build on `document_sdk.load_service()`. Each sample is
a single, self-contained Python file.

## Prerequisites

```bash
cd local-dotnet-worker
dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts
```

> Replace `win-x64` with `linux-x64`, `osx-arm64`, or `osx-x64` as needed.

## Available Samples

| Sample | What it does |
|---|---|
| [`01_hello_world.py`](01_hello_world.py) | Minimal end-to-end call: produce one DOCX and one PDF. |
| [`02_batch_invoices.py`](02_batch_invoices.py) | Loop over a list of records and emit a DOCX per record. |
| [`03_custom_output_paths.py`](03_custom_output_paths.py) | Use a custom output directory and a non-default bundle path. |
| [`04_excel_to_pdf.py`](04_excel_to_pdf.py) | Convert an XLSX workbook to PDF (auto-generates sample input if needed). |
| [`05_powerpoint_to_pdf.py`](05_powerpoint_to_pdf.py) | Convert a PPTX presentation to PDF (auto-generates sample input if needed). |
| [`06_watermark_pdf.py`](06_watermark_pdf.py) | Overlay a diagonal text watermark on every page of a PDF. |

## Run

Run any sample as a module from the `local-dotnet-worker` folder:

```bash
python -m samples.01_hello_world
```
