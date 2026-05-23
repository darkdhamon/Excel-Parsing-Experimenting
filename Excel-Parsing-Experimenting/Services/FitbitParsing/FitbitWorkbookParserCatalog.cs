namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public sealed class FitbitWorkbookParserCatalog
{
    private readonly IReadOnlyList<IFitbitWorkbookParser> _orderedParsers;
    private readonly IReadOnlyDictionary<WorkbookParserKind, IFitbitWorkbookParser> _parsersByKind;

    public FitbitWorkbookParserCatalog(IEnumerable<IFitbitWorkbookParser> parsers)
    {
        _orderedParsers = parsers
            .OrderBy(parser => parser.Kind)
            .ToList();

        _parsersByKind = _orderedParsers.ToDictionary(parser => parser.Kind);
    }

    public IReadOnlyList<IFitbitWorkbookParser> All => _orderedParsers;

    public IFitbitWorkbookParser Get(WorkbookParserKind kind)
    {
        return _parsersByKind[kind];
    }
}
