namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public sealed record WorksheetData(string Name, IReadOnlyList<WorksheetRowData> Rows);

public sealed record WorksheetRowData(int RowNumber, IReadOnlyList<object?> Values)
{
    public object? ValueAt(int index)
    {
        return index < Values.Count ? Values[index] : null;
    }
}
