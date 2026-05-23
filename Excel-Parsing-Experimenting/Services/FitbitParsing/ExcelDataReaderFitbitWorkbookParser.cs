using System.Data;
using System.Text;
using ExcelDataReader;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public sealed class ExcelDataReaderFitbitWorkbookParser : FitbitWorkbookParserBase
{
    public ExcelDataReaderFitbitWorkbookParser(FitbitWorkbookMapper mapper)
        : base(mapper)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public override WorkbookParserKind Kind => WorkbookParserKind.ExcelDataReader;

    public override string DisplayName => "ExcelDataReader";

    public override string Summary =>
        "Best format coverage in the comparison, including the legacy .xls Fitbit export.";

    public override string LibraryUrl => "https://github.com/ExcelDataReader/ExcelDataReader";

    public override IReadOnlyList<string> SupportedExtensions => [".xls", ".xlsx"];

    public override IReadOnlyList<string> Pros =>
    [
        "Reads both .xls and .xlsx exports.",
        "Straightforward workbook-to-rows extraction.",
        "Useful as the broadest compatibility baseline."
    ];

    public override IReadOnlyList<string> Cons =>
    [
        "Read-only API focused on data extraction.",
        "Lower-level formatting semantics than EPPlus."
    ];

    protected override IReadOnlyList<WorksheetData> ReadWorksheets(Stream workbookStream)
    {
        using var reader = ExcelReaderFactory.CreateReader(workbookStream);
        var workbook = reader.AsDataSet();
        var worksheets = new List<WorksheetData>(workbook.Tables.Count);

        foreach (DataTable table in workbook.Tables)
        {
            var rows = new List<WorksheetRowData>(table.Rows.Count);

            for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var values = table.Rows[rowIndex]
                    .ItemArray
                    .Select(value => value is DBNull ? null : value)
                    .ToArray();

                rows.Add(new WorksheetRowData(rowIndex + 1, values));
            }

            worksheets.Add(new WorksheetData(table.TableName, rows));
        }

        return worksheets;
    }
}
