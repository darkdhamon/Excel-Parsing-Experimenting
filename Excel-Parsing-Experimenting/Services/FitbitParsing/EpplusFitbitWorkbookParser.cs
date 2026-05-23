using OfficeOpenXml;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public sealed class EpplusFitbitWorkbookParser : FitbitWorkbookParserBase
{
    public EpplusFitbitWorkbookParser(FitbitWorkbookMapper mapper)
        : base(mapper)
    {
        EpplusLicenseBootstrapper.EnsureConfigured();
    }

    public override WorkbookParserKind Kind => WorkbookParserKind.Epplus;

    public override string DisplayName => "EPPlus";

    public override string Summary =>
        "High-level workbook API with clean .xlsx access and the option to grow into write scenarios.";

    public override string LibraryUrl => "https://github.com/EPPlusSoftware/EPPlus";

    public override IReadOnlyList<string> SupportedExtensions => [".xlsx"];

    public override IReadOnlyList<string> Pros =>
    [
        "Friendly object model for workbook navigation.",
        "Reads and writes .xlsx workbooks.",
        "Now fully maps body, activity, and sleep sheets in this repo."
    ];

    public override IReadOnlyList<string> Cons =>
    [
        "Only supports .xlsx files.",
        "Requires explicit EPPlus license configuration."
    ];

    protected override IReadOnlyList<WorksheetData> ReadWorksheets(Stream workbookStream)
    {
        using var package = new ExcelPackage(workbookStream);
        var worksheets = new List<WorksheetData>(package.Workbook.Worksheets.Count);

        foreach (var worksheet in package.Workbook.Worksheets)
        {
            if (worksheet.Dimension is null)
            {
                worksheets.Add(new WorksheetData(worksheet.Name, []));
                continue;
            }

            var rows = new List<WorksheetRowData>(worksheet.Dimension.End.Row);

            for (var rowNumber = 1; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
            {
                var values = new object?[worksheet.Dimension.End.Column];
                for (var columnNumber = 1; columnNumber <= worksheet.Dimension.End.Column; columnNumber++)
                {
                    values[columnNumber - 1] = worksheet.Cells[rowNumber, columnNumber].Value;
                }

                rows.Add(new WorksheetRowData(rowNumber, values));
            }

            worksheets.Add(new WorksheetData(worksheet.Name, rows));
        }

        return worksheets;
    }
}
