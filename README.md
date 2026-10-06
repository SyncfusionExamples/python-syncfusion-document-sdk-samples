# Python + Syncfusion Document SDK Samples

This repository contains two complete, runnable samples demonstrating how to generate Word (`.docx`) documents and convert them to PDF using the **Syncfusion .NET Document SDK** (DocIO and DocIORenderer) from **Python 3.9+**.

Microsoft Office is **not** required. All document generation and PDF rendering run locally.

## Two Approaches

| Feature | Approach 1: Local .NET Worker | Approach 2: Python.NET Wrapper |
|---|---|---|
| **Location** | [`local-dotnet-worker/`](local-dotnet-worker/) | [`pythonnet-wrapper/`](pythonnet-wrapper/) |
| **How it connects** | Separate process spawned via standard `subprocess` | In-process via Python.NET (`pythonnet` CoreCLR host) |
| **Python requirements** | Python 3.9+, **zero pip packages** | Python 3.12 (or 3.9+ with compatible CPython), `pythonnet>=3.1.0` |
| **.NET requirements** | .NET 8 SDK (build), .NET 8 Runtime | .NET 8 SDK (build), .NET 8 Runtime |
| **IPC mechanism** | JSON over standard input (`stdin`) | Direct managed memory function calls |
| **Best for** | Strict Python environments, batch jobs, CLI tools, maximum isolation | High-frequency calls, latency-sensitive workflows |

---

## Approach Documentation

The two approaches are **independent self-contained samples**. Pick one and read its `README.md`; the table below links to each:

| Approach | Quickstart README | Python client |
|---|---|---|
| Local .NET Worker | [local-dotnet-worker/README.md](local-dotnet-worker/README.md) | [local-dotnet-worker/document_sdk.py](local-dotnet-worker/document_sdk.py) |
| Python.NET Wrapper | [pythonnet-wrapper/README.md](pythonnet-wrapper/README.md) | [pythonnet-wrapper/document_sdk.py](pythonnet-wrapper/document_sdk.py) |

Both clients expose the same public surface (`load_service()` returning a `DocumentService` with `create_docx`, `create_pdf`, `excel_to_pdf`, `powerpoint_to_pdf`, `watermark_pdf`, and two sample-input helpers), so the same caller code works with either approach. For the C# side, see the `DocumentBridge/` folder inside each approach.

---

## Samples

Every approach ships with a small `samples/` folder of runnable scripts that build on `document_sdk.load_service()`:

| Sample | Local Worker | Python.NET |
|---|---|---|
| 01 — Hello World (DOCX + PDF) | [link](local-dotnet-worker/samples/01_hello_world.py) | [link](pythonnet-wrapper/samples/01_hello_world.py) |
| 02 — Batch generation loop | [link](local-dotnet-worker/samples/02_batch_invoices.py) | [link](pythonnet-wrapper/samples/02_batch_invoices.py) |
| 03 — Custom output paths & bundle | [link](local-dotnet-worker/samples/03_custom_output_paths.py) | [link](pythonnet-wrapper/samples/03_custom_output_paths.py) |
| 04 — Excel → PDF | [link](local-dotnet-worker/samples/04_excel_to_pdf.py) | [link](pythonnet-wrapper/samples/04_excel_to_pdf.py) |
| 05 — PowerPoint → PDF | [link](local-dotnet-worker/samples/05_powerpoint_to_pdf.py) | [link](pythonnet-wrapper/samples/05_powerpoint_to_pdf.py) |
| 06 — PDF watermark | [link](local-dotnet-worker/samples/06_watermark_pdf.py) | [link](pythonnet-wrapper/samples/06_watermark_pdf.py) |
| Per-approach index | [local-dotnet-worker/samples/README.md](local-dotnet-worker/samples/README.md) | [pythonnet-wrapper/samples/README.md](pythonnet-wrapper/samples/README.md) |

Run a sample (example, local worker):

```bash
cd local-dotnet-worker
python -m samples.01_hello_world
```

> ℹ️ Samples are run as modules (`python -m samples.XX`) so the parent folder
> containing `document_sdk.py` is on `sys.path`. Running them as plain
> scripts (`python samples/XX.py`) fails with
> `ModuleNotFoundError: No module named 'document_sdk'`.

---

## Quick Tour

### Approach 1: Local .NET Worker (`local-dotnet-worker/`)

```bash
cd local-dotnet-worker

# 1. Publish the .NET worker (Windows x64 example)
dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts

# 2. Run with plain Python (creates output/Sample.docx)
python document_sdk.py create-docx --trial

# 3. Render directly to PDF without an intermediate DOCX
python document_sdk.py create-pdf  --trial

# 4. Convert an existing XLSX/PPTX file to PDF, or watermark a PDF
python document_sdk.py excel-to-pdf      --input documents/input.xlsx --output output/workbook.pdf     --trial
python document_sdk.py powerpoint-to-pdf --input documents/input.pptx --output output/slides.pdf      --trial
python document_sdk.py watermark-pdf     --input output/workbook.pdf  --output output/watermarked.pdf --label CONFIDENTIAL --trial
```

### Approach 2: Python.NET Wrapper (`pythonnet-wrapper/`)

```bash
cd pythonnet-wrapper

# 1. Publish the .NET class library
dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts

# 2. Install pythonnet in a virtual environment
python -m venv .venv
.venv\Scripts\activate      # On Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt

# 3. Run in-process
python document_sdk.py create-docx --trial
python document_sdk.py create-pdf  --trial

# 4. Convert an existing XLSX/PPTX file to PDF, or watermark a PDF
python document_sdk.py excel-to-pdf      --input documents/input.xlsx --output output/workbook.pdf     --trial
python document_sdk.py powerpoint-to-pdf --input documents/input.pptx --output output/slides.pdf      --trial
python document_sdk.py watermark-pdf     --input output/workbook.pdf  --output output/watermarked.pdf --label CONFIDENTIAL --trial
```

---

## Licensing & Evaluation

Both samples support evaluation without a license key by passing the `--trial` flag (evaluation watermarks may appear on the output).

For licensed production use, set the `SYNCFUSION_LICENSE_KEY` environment variable and run without `--trial`:

```bash
# Windows PowerShell
$env:SYNCFUSION_LICENSE_KEY="YOUR_KEY_HERE"

# Linux / macOS
export SYNCFUSION_LICENSE_KEY="YOUR_KEY_HERE"
```

The .NET code automatically reads this variable via `Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY")` and registers it via `SyncfusionLicenseProvider.RegisterLicense(key)`. Keys are never passed on the command line.

---

## Tested Configurations

- **.NET SDK:** .NET 8.0 LTS
- **Syncfusion Packages:** `Syncfusion.DocIORenderer.Net.Core`, `Syncfusion.XlsIORenderer.Net.Core`, `Syncfusion.PresentationRenderer.Net.Core`, `Syncfusion.Pdf.Net.Core` (all 34.2.5)
- **Python:** 3.9.6, 3.12.x
- **Platforms:** Windows 10/11 x64 (`win-x64`), macOS Apple Silicon (`osx-arm64`), Linux Ubuntu x64 (`linux-x64`)

For the C# implementation, see [`local-dotnet-worker/DocumentBridge/`](local-dotnet-worker/DocumentBridge/) and [`pythonnet-wrapper/DocumentBridge/`](pythonnet-wrapper/DocumentBridge/).
