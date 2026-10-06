using System.Text;
using System.Text.Json;
using DocumentInterop;

// Reads a single JSON request from standard input, runs the requested Word/Excel/
// PowerPoint/PDF operation, and returns a nonzero exit code with an error message
// on stderr if anything fails. Python communicates with this worker purely through
// stdin/stdout/exit-code - no HTTP server, no shared memory.
try
{
    // Match the Python side, which encodes the request as UTF-8. Console.In
    // would otherwise default to the system code page (e.g. CP1252 on Windows)
    // and mangle any non-ASCII text passed via --text.
    using var stdin = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
    string requestJson = stdin.ReadToEnd();
    var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
    var request = JsonSerializer.Deserialize<Request>(requestJson, options)
        ?? throw new ArgumentException("A JSON request is required.");
    if (string.IsNullOrWhiteSpace(request.Operation))
        throw new ArgumentException("Request is missing 'Operation'.");
    if (string.IsNullOrWhiteSpace(request.Output))
        throw new ArgumentException("Request is missing 'Output'.");

    switch (request.Operation)
    {
        // Word: build from a string.
        case "create-docx":
            RequireField(request.Text, "Text", "create-docx");
            DocumentCreator.CreateDocx(request.Text!, request.Output, request.AllowTrial);
            break;
        case "create-pdf":
            RequireField(request.Text, "Text", "create-pdf");
            DocumentCreator.CreatePdf(request.Text!, request.Output, request.AllowTrial);
            break;
        // Excel / PowerPoint: convert an existing file to PDF.
        case "excel-to-pdf":
            RequireField(request.Input, "Input", "excel-to-pdf");
            DocumentCreator.ExcelToPdf(request.Input!, request.Output, request.AllowTrial);
            break;
        case "powerpoint-to-pdf":
            RequireField(request.Input, "Input", "powerpoint-to-pdf");
            DocumentCreator.PowerPointToPdf(request.Input!, request.Output, request.AllowTrial);
            break;
        // PDF: overlay a diagonal text watermark on every page.
        case "watermark-pdf":
            RequireField(request.Input, "Input", "watermark-pdf");
            RequireField(request.Label, "Label", "watermark-pdf");
            DocumentCreator.WatermarkPdf(request.Input!, request.Output, request.Label!, request.AllowTrial);
            break;
        // Sample input helpers used by the Python sample scripts.
        case "create-sample-xlsx":
            DocumentCreator.CreateSampleXlsx(request.Output, request.AllowTrial);
            break;
        case "create-sample-pptx":
            DocumentCreator.CreateSamplePptx(request.Output, request.AllowTrial);
            break;
        default:
            throw new ArgumentException("Unknown operation: " + request.Operation);
    }

    return 0;
}
catch (Exception error)
{
    Console.Error.WriteLine($"{error.GetType().Name}: {error.Message}");
    Console.Error.WriteLine(error.StackTrace);
    return 1;
}

static void RequireField(string? value, string name, string operation)
{
    if (string.IsNullOrWhiteSpace(value))
        throw new ArgumentException($"Operation '{operation}' requires a non-empty '{name}' field.");
}

// Shape of the JSON payload sent by document_sdk.py over standard input.
// `Operation` selects the C# entry point. `Input` is required for file -> file
// operations; `Text` is required for create-* operations; `Label` is required
// for watermark-pdf.
internal sealed record Request(
    string Operation,
    string? Text,
    string? Input,
    string Output,
    string? Label,
    bool AllowTrial);
