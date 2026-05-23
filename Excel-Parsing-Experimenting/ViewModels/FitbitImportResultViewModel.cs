using Excel_Parsing_Experimenting.Models;

namespace Excel_Parsing_Experimenting.ViewModels;

public sealed class FitbitImportResultViewModel
{
    public required string ParserName { get; init; }

    public required string ParserSummary { get; init; }

    public required string LibraryUrl { get; init; }

    public required IReadOnlyList<string> SupportedExtensions { get; init; }

    public required string UploadedFileName { get; init; }

    public FitbitData Data { get; init; } = new();
}
