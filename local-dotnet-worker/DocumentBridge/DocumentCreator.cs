using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Licensing;
using Syncfusion.Drawing;

namespace DocumentInterop;

/// <summary>
/// Creates a Word document in memory and saves it either as DOCX or as PDF,
/// converts an existing XLSX/PPTX file to PDF, and overlays a text watermark
/// onto every page of an existing PDF. Each entry point is independent; PDF
/// operations never write an intermediate DOCX/PPTX.
/// </summary>
public class DocumentCreator
{
    private static readonly object s_lock = new();
    private static bool s_licenseRegistered;

    public DocumentCreator()
    {
        ConfigureLicense();
    }

    private static void ConfigureLicense()
    {
        if (s_licenseRegistered) return;
        lock (s_lock)
        {
            if (s_licenseRegistered) return;
            string? key = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
            if (!string.IsNullOrWhiteSpace(key))
            {
                SyncfusionLicenseProvider.RegisterLicense(key);
            }
            s_licenseRegistered = true;
        }
    }

    // Default singleton instance used for static delegations
    private static readonly DocumentCreator s_default = new();

    // ---- Word: build from a string ----------------------------------------------------

    // Build a new Word document and save it directly as DOCX.
    public void CreateDocx(string text, string outputPath)
    {
        using var document = CreateDocument(text);
        using var output = OpenOutput(outputPath);
        document.Save(output, Syncfusion.DocIO.FormatType.Docx);
    }

    public static void CreateDocxStatic(string text, string outputPath) =>
        s_default.CreateDocx(text, outputPath);

    // Build a new Word document and render it straight to PDF via DocIORenderer.
    public void CreatePdf(string text, string outputPath)
    {
        using var document = CreateDocument(text);
        using var renderer = new DocIORenderer();
        using var pdf = renderer.ConvertToPDF(document);
        using var output = OpenOutput(outputPath);
        pdf.Save(output);
    }

    public static void CreatePdfStatic(string text, string outputPath) =>
        s_default.CreatePdf(text, outputPath);

    private static WordDocument CreateDocument(string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);
        var document = new WordDocument();
        var paragraph = document.AddSection().AddParagraph();
        var run = paragraph.AppendText(text);
        run.CharacterFormat.FontName = "Arial";
        run.CharacterFormat.FontSize = 12;
        return document;
    }

    // ---- Excel: convert an existing XLSX file to PDF ----------------------------------

    // Convert an XLSX workbook to PDF using its own print settings.
    public void ExcelToPdf(string inputPath, string outputPath)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("Input workbook not found.", inputPath);
        using var engine = new ExcelEngine();
        engine.Excel.DefaultVersion = ExcelVersion.Xlsx;
        using var stream = File.OpenRead(inputPath);
        var book = engine.Excel.Workbooks.Open(stream);
        try
        {
            using var renderer = new XlsIORenderer();
            using var pdf = renderer.ConvertToPDF(book);
            using var output = OpenOutput(outputPath);
            pdf.Save(output);
        }
        finally { book.Close(); }
    }

    public static void ExcelToPdfStatic(string inputPath, string outputPath) =>
        s_default.ExcelToPdf(inputPath, outputPath);

    // ---- PowerPoint: convert an existing PPTX file to PDF -----------------------------

    // Convert a PPTX presentation to PDF using Syncfusion's presentation renderer.
    public void PowerPointToPdf(string inputPath, string outputPath)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("Input presentation not found.", inputPath);
        using var stream = File.OpenRead(inputPath);
        using var deck = Presentation.Open(stream);
        using var pdf = PresentationToPdfConverter.Convert(deck);
        using var output = OpenOutput(outputPath);
        pdf.Save(output);
    }

    public static void PowerPointToPdfStatic(string inputPath, string outputPath) =>
        s_default.PowerPointToPdf(inputPath, outputPath);

    // ---- PDF: overlay a diagonal text watermark on every page -------------------------

    // Draw a translucent diagonal label on every page of an existing PDF.
    public void WatermarkPdf(string inputPath, string outputPath, string label)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(label);
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("Input PDF not found.", inputPath);
        using var stream = File.OpenRead(inputPath);
        using var document = new PdfLoadedDocument(stream);
        foreach (PdfPageBase page in document.Pages)
        {
            var graphics = page.Graphics;
            var size = graphics.ClientSize;
            // Measure with the largest allowed size, then auto-shrink to fit
            // the page width. Floor at 8pt so very long labels stay visible
            // instead of collapsing to ~1pt and disappearing.
            const float MaxFontSize = 28f;
            const float MinFontSize = 8f;
            var probeFont = new PdfStandardFont(PdfFontFamily.Helvetica, MaxFontSize);
            float measured = probeFont.MeasureString(label).Width;
            float fontSize = MaxFontSize * size.Width * 0.8f / Math.Max(1, measured);
            fontSize = Math.Clamp(fontSize, MinFontSize, MaxFontSize);
            PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, fontSize);
            var state = graphics.Save();
            try
            {
                graphics.SetTransparency(0.25f);
                graphics.TranslateTransform(size.Width / 2, size.Height / 2);
                graphics.RotateTransform(-30);
                graphics.DrawString(label, font, PdfBrushes.Gray,
                    new RectangleF(-size.Width / 2, -30, size.Width, 60),
                    new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle));
            }
            finally { graphics.Restore(state); }
        }
        using var output = OpenOutput(outputPath);
        document.Save(output);
    }

    public static void WatermarkPdfStatic(string inputPath, string outputPath, string label) =>
        s_default.WatermarkPdf(inputPath, outputPath, label);

    // ---- Sample input helpers (used only by the sample scripts) -----------------------

    // Build a tiny XLSX with one cell of text. Lets the samples run end-to-end without
    // shipping a binary fixture in the repository.
    public void CreateSampleXlsx(string outputPath)
    {
        using var engine = new ExcelEngine();
        engine.Excel.DefaultVersion = ExcelVersion.Xlsx;
        var book = engine.Excel.Workbooks.Create();
        try
        {
            var sheet = book.Worksheets[0];
            sheet.Range["A1"].Text = "Sample workbook";
            sheet.Range["A1"].CellStyle.Font.Bold = true;
            sheet.Range["A3"].Text = "Hello from Syncfusion Document SDK!";
            using var output = OpenOutput(outputPath);
            book.SaveAs(output);
        }
        finally { book.Close(); }
    }

    public static void CreateSampleXlsxStatic(string outputPath) =>
        s_default.CreateSampleXlsx(outputPath);

    // Build a tiny PPTX with a single slide. Lets the samples run end-to-end without
    // shipping a binary fixture in the repository.
    public void CreateSamplePptx(string outputPath)
    {
        var deck = Presentation.Create();
        var slide = deck.Slides.Add(SlideLayoutType.Blank);
        var shape = slide.Shapes.AddTextBox(40, 40, 600, 80);
        var paragraph = shape.TextBody.AddParagraph("Hello from Syncfusion Document SDK!");
        paragraph.Font.FontSize = 28;
        paragraph.Font.Bold = true;
        using var output = OpenOutput(outputPath);
        deck.Save(output);
    }

    public static void CreateSamplePptxStatic(string outputPath) =>
        s_default.CreateSamplePptx(outputPath);

    // ---- Common helpers ---------------------------------------------------------------

    private static FileStream OpenOutput(string outputPath)
    {
        string path = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        // FileMode.Create overwrites an existing file at the same path.
        return new FileStream(path, FileMode.Create, FileAccess.Write);
    }
}
