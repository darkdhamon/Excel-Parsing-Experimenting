using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public sealed class OpenXmlFitbitWorkbookParser : FitbitWorkbookParserBase
{
    private static readonly HashSet<uint> BuiltInDateFormatIds =
    [
        14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
    ];

    public OpenXmlFitbitWorkbookParser(FitbitWorkbookMapper mapper)
        : base(mapper)
    {
    }

    public override WorkbookParserKind Kind => WorkbookParserKind.OpenXml;

    public override string DisplayName => "Open XML SDK";

    public override string Summary =>
        "Lowest-level comparison path, now with robust shared-string and date handling for the full workbook.";

    public override string LibraryUrl => "https://github.com/dotnet/Open-XML-SDK";

    public override IReadOnlyList<string> SupportedExtensions => [".xlsx"];

    public override IReadOnlyList<string> Pros =>
    [
        "No Excel automation or Office dependency required.",
        "Fine-grained control over workbook structure.",
        "Good reference point for low-level Open XML parsing."
    ];

    public override IReadOnlyList<string> Cons =>
    [
        "More verbose than the other libraries.",
        "Date and style handling must be done manually."
    ];

    protected override IReadOnlyList<WorksheetData> ReadWorksheets(Stream workbookStream)
    {
        using var document = SpreadsheetDocument.Open(workbookStream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("Workbook part was missing.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook was missing.");
        var sheets = workbook.Sheets?.Elements<Sheet>() ?? [];
        var worksheets = new List<WorksheetData>();

        foreach (var sheet in sheets)
        {
            var sheetId = sheet.Id?.Value;
            if (string.IsNullOrWhiteSpace(sheetId))
            {
                continue;
            }

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet was missing.");
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var rows = new List<WorksheetRowData>();

            if (sheetData is not null)
            {
                foreach (var row in sheetData.Elements<Row>())
                {
                    var values = new List<object?>();

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var columnIndex = GetColumnIndex(cell.CellReference?.Value);
                        while (values.Count <= columnIndex)
                        {
                            values.Add(null);
                        }

                        values[columnIndex] = ReadCellValue(cell, workbookPart);
                    }

                    rows.Add(new WorksheetRowData((int)(row.RowIndex?.Value ?? (uint)(rows.Count + 1)), values));
                }
            }

            worksheets.Add(new WorksheetData(sheet.Name?.Value ?? "Unknown", rows));
        }

        return worksheets;
    }

    private static object? ReadCellValue(Cell cell, WorkbookPart workbookPart)
    {
        var rawValue = cell.CellValue?.Text ?? cell.InnerText;
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return null;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return ReadSharedString(rawValue, workbookPart);
        }

        if (cell.DataType?.Value == CellValues.InlineString || cell.DataType?.Value == CellValues.String)
        {
            return cell.InnerText;
        }

        if (cell.DataType?.Value == CellValues.Boolean)
        {
            return rawValue == "1";
        }

        if (cell.DataType?.Value == CellValues.Date)
        {
            return DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var date)
                ? date
                : rawValue;
        }

        if (IsDateFormattedCell(cell, workbookPart) &&
            double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var oaDate))
        {
            return DateTime.FromOADate(oaDate);
        }

        return double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number
            : rawValue;
    }

    private static string ReadSharedString(string rawValue, WorkbookPart workbookPart)
    {
        if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedStringIndex))
        {
            return rawValue;
        }

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        var sharedStringItem = sharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sharedStringIndex);
        if (sharedStringItem is null)
        {
            return rawValue;
        }

        if (sharedStringItem.Text is not null)
        {
            return sharedStringItem.Text.Text ?? string.Empty;
        }

        return string.Concat(sharedStringItem.Descendants<Text>().Select(text => text.Text));
    }

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return 0;
        }

        var columnValue = 0;
        foreach (var character in cellReference)
        {
            if (!char.IsLetter(character))
            {
                break;
            }

            columnValue = (columnValue * 26) + (char.ToUpperInvariant(character) - 'A' + 1);
        }

        return Math.Max(0, columnValue - 1);
    }

    private static bool IsDateFormattedCell(Cell cell, WorkbookPart workbookPart)
    {
        var styleIndex = cell.StyleIndex?.Value;
        if (styleIndex is null)
        {
            return false;
        }

        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        var cellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList();
        if (cellFormats is null || styleIndex.Value >= cellFormats.Count)
        {
            return false;
        }

        var numberFormatId = cellFormats[(int)styleIndex.Value].NumberFormatId?.Value;
        if (numberFormatId is null)
        {
            return false;
        }

        if (BuiltInDateFormatIds.Contains(numberFormatId.Value))
        {
            return true;
        }

        var customFormatCode = stylesheet?.NumberingFormats?
            .Elements<NumberingFormat>()
            .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId.Value)?
            .FormatCode?
            .Value;

        return !string.IsNullOrWhiteSpace(customFormatCode) && LooksLikeDateFormat(customFormatCode);
    }

    private static bool LooksLikeDateFormat(string formatCode)
    {
        var normalized = formatCode.ToLowerInvariant();

        return normalized.Contains("yy", StringComparison.Ordinal) ||
               normalized.Contains("dd", StringComparison.Ordinal) ||
               normalized.Contains("hh", StringComparison.Ordinal) ||
               normalized.Contains("ss", StringComparison.Ordinal) ||
               normalized.Contains("am/pm", StringComparison.Ordinal) ||
               normalized.Contains("[h]", StringComparison.Ordinal);
    }
}
