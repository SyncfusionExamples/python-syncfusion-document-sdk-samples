# Python + Syncfusion Document SDK Samples

A simple, fast solution for creating Word (`.docx`), Excel (`.xlsx`), PowerPoint (`.pptx`), and PDF documents from **Python** using the **Syncfusion .NET Document SDK**.

Microsoft Office is **not** required. All document processing runs locally.

---

## Prerequisites

Ensure the following prerequisites are installed:

- **[.NET 10.0 LTS SDK](https://dotnet.microsoft.com/download/dotnet/10.0)**
- **[Python 3.9+](https://www.python.org/downloads/)**
- **Syncfusion .NET Document SDK packages:**
  - `Syncfusion.DocIORenderer.Net.Core`
  - `Syncfusion.XlsIORenderer.Net.Core`
  - `Syncfusion.PresentationRenderer.Net.Core`
  - `Syncfusion.Pdf.Net.Core`

---

## Quickstart

Choose either approach based on your needs:

| Approach | Folder | Best For | Python Packages |
|---|---|---|---|
| **Approach 1: Local .NET Worker** | [`local-dotnet-worker/`](local-dotnet-worker/) | Subprocess isolation, CLI workflows | **Zero pip packages** |
| **Approach 2: Python.NET Wrapper** | [`pythonnet-wrapper/`](pythonnet-wrapper/) | In-process execution, high-performance loops | `pip install pythonnet` (if not installed) |

---

### Approach 1: Local .NET Worker (Zero Pip Dependencies)

1. **Navigate and publish the .NET worker:**
   ```bash
   cd local-dotnet-worker
   dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts
   ```

2. **Generate documents via CLI:**
   ```bash
   python document_sdk.py create-docx
   python document_sdk.py create-pdf
   ```

3. **Run samples:**
   ```bash
   python -m samples.01_hello_world
   python -m samples.02_batch_invoices
   python -m samples.04_excel_to_pdf
   python -m samples.05_powerpoint_to_pdf
   python -m samples.06_watermark_pdf
   ```

---

### Approach 2: Python.NET Wrapper (In-Process Execution)

1. **Navigate and publish the class library:**
   ```bash
   cd pythonnet-wrapper
   dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts
   ```

2. **Install pythonnet (only if not already installed):**
   ```bash
   pip install pythonnet
   ```
   *(Skip this step if `pythonnet` is already installed in your Python environment).*

3. **Generate documents via CLI:**
   ```bash
   python document_sdk.py create-docx
   python document_sdk.py create-pdf
   ```

4. **Run samples:**
   ```bash
   python -m samples.01_hello_world
   python -m samples.02_batch_invoices
   python -m samples.04_excel_to_pdf
   python -m samples.05_powerpoint_to_pdf
   python -m samples.06_watermark_pdf
   ```

> **Platform Note:** If running on non-Windows platforms, replace `-r win-x64` with `-r linux-x64`, `-r osx-arm64`, or `-r osx-x64`.

---

## Samples Overview

Both approaches include the same self-contained samples:

| Sample | Description |
|---|---|
| `01_hello_world.py` | Create a Word document and direct PDF |
| `02_batch_invoices.py` | Generate multiple documents in a loop |
| `03_custom_output_paths.py` | Specify custom output and bundle paths |
| `04_excel_to_pdf.py` | Convert an Excel workbook to PDF |
| `05_powerpoint_to_pdf.py` | Convert a PowerPoint presentation to PDF |
| `06_watermark_pdf.py` | Apply a text watermark to PDF pages |
