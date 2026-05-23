using Excel_Parsing_Experimenting.Models;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public abstract class FitbitWorkbookParserBase : IFitbitWorkbookParser
{
    private readonly FitbitWorkbookMapper _mapper;

    protected FitbitWorkbookParserBase(FitbitWorkbookMapper mapper)
    {
        _mapper = mapper;
    }

    public abstract WorkbookParserKind Kind { get; }

    public abstract string DisplayName { get; }

    public abstract string Summary { get; }

    public abstract string LibraryUrl { get; }

    public abstract IReadOnlyList<string> SupportedExtensions { get; }

    public abstract IReadOnlyList<string> Pros { get; }

    public abstract IReadOnlyList<string> Cons { get; }

    public FitbitData Parse(Stream workbookStream)
    {
        workbookStream.Position = 0;
        return _mapper.Map(ReadWorksheets(workbookStream));
    }

    protected abstract IReadOnlyList<WorksheetData> ReadWorksheets(Stream workbookStream);

    public bool SupportsExtension(string fileExtension)
    {
        return SupportedExtensions.Contains(fileExtension, StringComparer.OrdinalIgnoreCase);
    }
}
