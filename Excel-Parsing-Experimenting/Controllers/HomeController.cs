using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Excel_Parsing_Experimenting.Models;
using Excel_Parsing_Experimenting.Services.FitbitParsing;
using Excel_Parsing_Experimenting.ViewModels;

namespace Excel_Parsing_Experimenting.Controllers;

public class HomeController : Controller
{
    private static readonly IReadOnlyDictionary<WorkbookParserKind, string> UploadActions =
        new Dictionary<WorkbookParserKind, string>
        {
            [WorkbookParserKind.ExcelDataReader] = "ImportFitbitDataEdr",
            [WorkbookParserKind.Epplus] = "ImportFitbitDataEpplus",
            [WorkbookParserKind.OpenXml] = "ImportFitbitDataOpenXml"
        };

    private static readonly HashSet<string> AllowedSamples =
    [
        "fitbit_export_20180109.xls",
        "fitbit_export_20180109.xlsx"
    ];

    private readonly FitbitWorkbookParserCatalog _parserCatalog;
    private readonly IWebHostEnvironment _environment;

    public HomeController(FitbitWorkbookParserCatalog parserCatalog, IWebHostEnvironment environment)
    {
        _parserCatalog = parserCatalog;
        _environment = environment;
    }

    public IActionResult Index()
    {
        var viewModel = new HomePageViewModel
        {
            ParserCards = _parserCatalog.All
                .Select(parser => new ParserCardViewModel
                {
                    DisplayName = parser.DisplayName,
                    Summary = parser.Summary,
                    LibraryUrl = parser.LibraryUrl,
                    UploadAction = UploadActions[parser.Kind],
                    SupportedExtensions = parser.SupportedExtensions,
                    Pros = parser.Pros,
                    Cons = parser.Cons
                })
                .ToList()
        };

        return View(viewModel);
    }

    [HttpGet("sample/{fileName}")]
    public IActionResult DownloadSample(string fileName)
    {
        if (!AllowedSamples.Contains(fileName))
        {
            return NotFound();
        }

        var samplePath = Path.GetFullPath(
            Path.Combine(_environment.ContentRootPath, "..", "Sample Upload", fileName));

        if (!System.IO.File.Exists(samplePath))
        {
            return NotFound();
        }

        return PhysicalFile(samplePath, GetContentType(fileName), fileName);
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    private static string GetContentType(string fileName)
    {
        return Path.GetExtension(fileName).ToLowerInvariant() switch
        {
            ".xls" => "application/vnd.ms-excel",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            _ => "application/octet-stream"
        };
    }
}
