# Local .NET Worker Sample

Creates Word, Excel, PowerPoint, and PDF documents from Python by spawning a local .NET process (`DocumentBridge`) via `subprocess`. Requires only standard Python and the .NET 8 runtime — no pip packages needed.

## Quickstart

### 1. Publish the .NET Worker

```bash
dotnet publish DocumentBridge/DocumentBridge.csproj -c Release -r win-x64 --self-contained false -o artifacts
```

> Replace `win-x64` with your target platform: `linux-x64`, `osx-arm64`, or `osx-x64`.

### 2. Run CLI Commands

```bash
# Create DOCX and direct PDF
python document_sdk.py create-docx
python document_sdk.py create-pdf

# Convert Excel or PowerPoint to PDF
python document_sdk.py excel-to-pdf      --input documents/input.xlsx  --output output/workbook.pdf
python document_sdk.py powerpoint-to-pdf --input documents/input.pptx  --output output/slides.pdf

# Watermark a PDF
python document_sdk.py watermark-pdf     --input output/workbook.pdf   --output output/watermarked.pdf --label CONFIDENTIAL
```

### 3. Run Samples

```bash
python -m samples.01_hello_world
python -m samples.02_batch_invoices
python -m samples.03_custom_output_paths
python -m samples.04_excel_to_pdf
python -m samples.05_powerpoint_to_pdf
python -m samples.06_watermark_pdf
```

## Available Samples

| Sample | Description |
|---|---|
| `01_hello_world.py` | Minimal end-to-end DOCX + PDF generation. |
| `02_batch_invoices.py` | Loop over records and emit one DOCX each. |
| `03_custom_output_paths.py` | Custom output folder and non-default bundle path. |
| `04_excel_to_pdf.py` | Convert an XLSX workbook to PDF. |
| `05_powerpoint_to_pdf.py` | Convert a PPTX presentation to PDF. |
| `06_watermark_pdf.py` | Overlay a diagonal text watermark on every page of a PDF. |
