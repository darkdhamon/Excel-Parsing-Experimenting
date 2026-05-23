using Excel_Parsing_Experimenting.Models;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

public sealed class FitbitWorkbookMapper
{
    public FitbitData Map(IEnumerable<WorksheetData> worksheets)
    {
        var data = new FitbitData();

        foreach (var worksheet in worksheets)
        {
            switch (worksheet.Name.Trim())
            {
                case "Body":
                    MapBodyWorksheet(worksheet, data);
                    break;
                case "Activities":
                    MapActivitiesWorksheet(worksheet, data);
                    break;
                case "Sleep":
                    MapSleepWorksheet(worksheet, data);
                    break;
                default:
                    break;
            }
        }

        return data;
    }

    private static void MapBodyWorksheet(WorksheetData worksheet, FitbitData data)
    {
        foreach (var row in worksheet.Rows)
        {
            if (IsIgnorableRow(row, "Date"))
            {
                continue;
            }

            if (TryCreateBodyEntry(row, out var entry))
            {
                data.BodyDataEntries.Add(entry);
                continue;
            }

            data.ErrorMessages.Add($"Body row {row.RowNumber} could not be parsed and was skipped.");
        }
    }

    private static void MapActivitiesWorksheet(WorksheetData worksheet, FitbitData data)
    {
        foreach (var row in worksheet.Rows)
        {
            if (IsIgnorableRow(row, "Date"))
            {
                continue;
            }

            if (TryCreateActivityEntry(row, out var entry))
            {
                data.ActivityDataEntries.Add(entry);
                continue;
            }

            data.ErrorMessages.Add($"Activities row {row.RowNumber} could not be parsed and was skipped.");
        }
    }

    private static void MapSleepWorksheet(WorksheetData worksheet, FitbitData data)
    {
        foreach (var row in worksheet.Rows)
        {
            if (IsIgnorableRow(row, "Start Time"))
            {
                continue;
            }

            if (TryCreateSleepEntry(row, out var entry))
            {
                data.SleepDataEntries.Add(entry);
                continue;
            }

            data.ErrorMessages.Add($"Sleep row {row.RowNumber} could not be parsed and was skipped.");
        }
    }

    private static bool IsIgnorableRow(WorksheetRowData row, string firstHeaderName)
    {
        if (row.Values.All(FitbitCellConversions.IsBlank))
        {
            return true;
        }

        if (!FitbitCellConversions.TryToString(row.ValueAt(0), out var firstCellValue))
        {
            return false;
        }

        return string.Equals(firstCellValue, firstHeaderName, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryCreateBodyEntry(WorksheetRowData row, out BodyDataEntry entry)
    {
        entry = default!;

        if (!FitbitCellConversions.TryToDateTime(row.ValueAt(0), out var date) ||
            !FitbitCellConversions.TryToDouble(row.ValueAt(1), out var weight) ||
            !FitbitCellConversions.TryToDouble(row.ValueAt(2), out var bmi) ||
            !FitbitCellConversions.TryToDouble(row.ValueAt(3), out var fat))
        {
            return false;
        }

        entry = new BodyDataEntry(date, weight, bmi, fat);
        return true;
    }

    private static bool TryCreateActivityEntry(WorksheetRowData row, out ActivityDataEntry entry)
    {
        entry = default!;

        if (!FitbitCellConversions.TryToDateTime(row.ValueAt(0), out var date) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(1), out var caloriesBurned) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(2), out var steps) ||
            !FitbitCellConversions.TryToDouble(row.ValueAt(3), out var distance) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(4), out var floors) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(5), out var minutesSedentary) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(6), out var minutesLightlyActive) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(7), out var minutesFairlyActive) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(8), out var minutesVeryActive) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(9), out var activityCalories))
        {
            return false;
        }

        entry = new ActivityDataEntry(
            date,
            caloriesBurned,
            steps,
            distance,
            floors,
            minutesSedentary,
            minutesLightlyActive,
            minutesFairlyActive,
            minutesVeryActive,
            activityCalories);

        return true;
    }

    private static bool TryCreateSleepEntry(WorksheetRowData row, out SleepDataEntry entry)
    {
        entry = default!;

        if (!FitbitCellConversions.TryToDateTime(row.ValueAt(0), out var startTime) ||
            !FitbitCellConversions.TryToDateTime(row.ValueAt(1), out var endTime) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(2), out var minutesAsleep) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(3), out var minutesAwake) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(4), out var numberOfAwakenings) ||
            !FitbitCellConversions.TryToInt32(row.ValueAt(5), out var timeInBed) ||
            !FitbitCellConversions.TryToNullableInt32(row.ValueAt(6), out var minutesRemSleep) ||
            !FitbitCellConversions.TryToNullableInt32(row.ValueAt(7), out var minutesLightSleep) ||
            !FitbitCellConversions.TryToNullableInt32(row.ValueAt(8), out var minutesDeepSleep))
        {
            return false;
        }

        entry = new SleepDataEntry(
            startTime,
            endTime,
            minutesAsleep,
            minutesAwake,
            numberOfAwakenings,
            timeInBed,
            minutesRemSleep,
            minutesLightSleep,
            minutesDeepSleep);

        return true;
    }
}
