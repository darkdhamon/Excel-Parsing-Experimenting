using Excel_Parsing_Experimenting.Models;
using Excel_Parsing_Experimenting.Services.FitbitParsing;
using Excel_Parsing_Experimenting.ViewModels;
using Microsoft.AspNetCore.Mvc;

namespace Excel_Parsing_Experimenting.Controllers;

public class UploadController : Controller
{
    private readonly FitbitWorkbookParserCatalog _parserCatalog;

    public UploadController(FitbitWorkbookParserCatalog parserCatalog)
    {
        _parserCatalog = parserCatalog;
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public Task<IActionResult> ImportFitbitDataEdr(IFormFile? file, CancellationToken cancellationToken)
    {
        return ImportAsync(WorkbookParserKind.ExcelDataReader, file, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public Task<IActionResult> ImportFitbitDataEpplus(IFormFile? file, CancellationToken cancellationToken)
    {
        return ImportAsync(WorkbookParserKind.Epplus, file, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public Task<IActionResult> ImportFitbitDataOpenXml(IFormFile? file, CancellationToken cancellationToken)
    {
        return ImportAsync(WorkbookParserKind.OpenXml, file, cancellationToken);
    }

    private async Task<IActionResult> ImportAsync(
        WorkbookParserKind parserKind,
        IFormFile? file,
        CancellationToken cancellationToken)
    {
        var parser = _parserCatalog.Get(parserKind);
        var result = new FitbitData();

        if (file is null || file.Length <= 0)
        {
            result.ErrorMessages.Add("Please choose a Fitbit workbook before uploading.");
            return View("ImportFitbitData", BuildViewModel(parser, file?.FileName, result));
        }

        var fileExtension = Path.GetExtension(file.FileName);
        if (!parser.SupportsExtension(fileExtension))
        {
            result.ErrorMessages.Add(
                $"{parser.DisplayName} supports {string.Join(", ", parser.SupportedExtensions)} files. " +
                $"The uploaded file uses {fileExtension}.");

            return View("ImportFitbitData", BuildViewModel(parser, file.FileName, result));
        }

        try
        {
            await using var workbookStream = new MemoryStream();
            await file.CopyToAsync(workbookStream, cancellationToken);
            workbookStream.Position = 0;

            result = parser.Parse(workbookStream);
        }
        catch (Exception exception)
        {
            result.ErrorMessages.Add($"The workbook could not be parsed: {exception.Message}");
        }

        return View("ImportFitbitData", BuildViewModel(parser, file.FileName, result));
    }

    private static FitbitImportResultViewModel BuildViewModel(
        IFitbitWorkbookParser parser,
        string? uploadedFileName,
        FitbitData data)
    {
        return new FitbitImportResultViewModel
        {
            ParserName = parser.DisplayName,
            ParserSummary = parser.Summary,
            LibraryUrl = parser.LibraryUrl,
            SupportedExtensions = parser.SupportedExtensions,
            UploadedFileName = uploadedFileName ?? "No file selected",
            Data = data
        };
    }
}
