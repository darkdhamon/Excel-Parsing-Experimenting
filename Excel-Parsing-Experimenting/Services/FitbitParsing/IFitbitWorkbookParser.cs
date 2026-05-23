using Excel_Parsing_Experimenting.Models;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public interface IFitbitWorkbookParser
{
    WorkbookParserKind Kind { get; }

    string DisplayName { get; }

    string Summary { get; }

    string LibraryUrl { get; }

    IReadOnlyList<string> SupportedExtensions { get; }

    IReadOnlyList<string> Pros { get; }

    IReadOnlyList<string> Cons { get; }

    bool SupportsExtension(string fileExtension);

    FitbitData Parse(Stream workbookStream);
}
