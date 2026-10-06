# Samples — Local .NET Worker

Runnable examples that build on `document_sdk.load_service()`. Each sample is
a single, self-contained Python file. They all share the same prerequisites:

```bash
cd local-dotnet-worker
dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts
# (replace win-x64 with osx-arm64, osx-x64, or linux-x64 as needed)
```

All samples use `--trial` by default so they run without a Syncfusion license.
Remove `--trial` and set `SYNCFUSION_LICENSE_KEY` to use a real license.

| Sample | What it does |
|---|---|
| [`01_hello_world.py`](01_hello_world.py) | Minimal end-to-end call: produce one DOCX and one PDF. |
| [`02_batch_invoices.py`](02_batch_invoices.py) | Loop over a list of records and emit a DOCX per record. |
| [`03_custom_output_paths.py`](03_custom_output_paths.py) | Use `--output` / `--bundle` and run a per-format run. |
| [`04_excel_to_pdf.py`](04_excel_to_pdf.py) | Convert an XLSX workbook to PDF (auto-generates a sample workbook if needed). |
| [`05_powerpoint_to_pdf.py`](05_powerpoint_to_pdf.py) | Convert a PPTX presentation to PDF (auto-generates a sample deck if needed). |
| [`06_watermark_pdf.py`](06_watermark_pdf.py) | Overlay a diagonal text watermark on every page of an existing PDF. |

Run any sample directly (from the approach's root folder):

```bash
cd local-dotnet-worker
python -m samples.01_hello_world
```

> Run samples as modules (`python -m samples.XX`) so the parent folder
> containing `document_sdk.py` is on `sys.path`. A bare `python samples/XX.py`
> would fail with `ModuleNotFoundError: No module named 'document_sdk'`.
