using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine("Usage: OoxmlValidator <file.docx|file.xlsx|file.pptx> [--verbose] [--json]");
            return 1;
        }

        string filePath = args[0];
        bool verbose = args.Contains("--verbose");
        bool jsonOutput = args.Contains("--json");

        if (!File.Exists(filePath))
        {
            Console.Error.WriteLine($"File not found: {filePath}");
            return 1;
        }

        try
        {
            var extension = Path.GetExtension(filePath).ToLower();
            var errors = extension switch
            {
                ".docx" => ValidateWordDocument(filePath, verbose),
                ".xlsx" => ValidateSpreadsheet(filePath, verbose),
                ".pptx" => ValidatePresentation(filePath, verbose),
                _ => throw new NotSupportedException($"Unsupported file type: {extension}")
            };

            if (jsonOutput)
            {
                var result = new
                {
                    file = filePath,
                    valid = errors.Length == 0,
                    errorCount = errors.Length,
                    errors = errors.Select(e => new
                    {
                        description = e.Description,
                        errorType = e.ErrorType.ToString(),
                        path = e.Path?.XPath,
                        partUri = e.Part?.Uri?.ToString(),
                        node = e.Node?.OuterXml?.Substring(0, Math.Min(200, e.Node?.OuterXml?.Length ?? 0))
                    })
                };
                Console.WriteLine(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
            }
            else
            {
                if (errors.Length == 0)
                {
                    Console.WriteLine($"✓ {filePath} is valid");
                }
                else
                {
                    Console.WriteLine($"✗ {filePath} has {errors.Length} validation error(s):\n");
                    foreach (var error in errors)
                    {
                        Console.WriteLine($"  Error: {error.Description}");
                        Console.WriteLine($"  Type:  {error.ErrorType}");
                        if (error.Path != null)
                            Console.WriteLine($"  Path:  {error.Path.XPath}");
                        if (error.Part != null)
                            Console.WriteLine($"  Part:  {error.Part.Uri}");
                        if (verbose && error.Node != null)
                        {
                            var xml = error.Node.OuterXml;
                            if (xml.Length > 200) xml = xml.Substring(0, 200) + "...";
                            Console.WriteLine($"  XML:   {xml}");
                        }
                        Console.WriteLine();
                    }
                }
            }

            return errors.Length == 0 ? 0 : 2;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            if (verbose)
                Console.Error.WriteLine(ex.StackTrace);
            return 1;
        }
    }

    static ValidationErrorInfo[] ValidateWordDocument(string path, bool verbose)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        return ValidateDocument(doc, verbose);
    }

    static ValidationErrorInfo[] ValidateSpreadsheet(string path, bool verbose)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        return ValidateDocument(doc, verbose);
    }

    static ValidationErrorInfo[] ValidatePresentation(string path, bool verbose)
    {
        using var doc = PresentationDocument.Open(path, false);
        return ValidateDocument(doc, verbose);
    }

    static ValidationErrorInfo[] ValidateDocument(OpenXmlPackage doc, bool verbose)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(doc).ToArray();

        if (verbose)
        {
            Console.WriteLine("Document parts:");
            foreach (var part in doc.Parts)
            {
                Console.WriteLine($"  {part.OpenXmlPart.Uri}");
            }
            Console.WriteLine();
        }

        return errors;
    }
}
