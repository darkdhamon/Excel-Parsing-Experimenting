using Excel_Parsing_Experimenting.Models;
using Excel_Parsing_Experimenting.Services.FitbitParsing;

namespace Excel_Parsing_Experimenting.Tests;

public sealed class ParserIntegrationTests
{
    private static readonly FitbitWorkbookMapper Mapper = new();
    private static readonly IFitbitWorkbookParser ExcelDataReaderParser = new ExcelDataReaderFitbitWorkbookParser(Mapper);
    private static readonly IFitbitWorkbookParser EpplusParser = new EpplusFitbitWorkbookParser(Mapper);
    private static readonly IFitbitWorkbookParser OpenXmlParser = new OpenXmlFitbitWorkbookParser(Mapper);

    [Fact]
    public void ExcelDataReader_parses_both_sample_formats()
    {
        var xlsResult = ParseWorkbook(ExcelDataReaderParser, "fitbit_export_20180109.xls");
        var xlsxResult = ParseWorkbook(ExcelDataReaderParser, "fitbit_export_20180109.xlsx");

        AssertParsedData(xlsResult);
        AssertParsedData(xlsxResult);
        AssertEquivalent(xlsxResult, xlsResult);
    }

    [Fact]
    public void Epplus_matches_excel_data_reader_for_xlsx_sample()
    {
        var reference = ParseWorkbook(ExcelDataReaderParser, "fitbit_export_20180109.xlsx");
        var candidate = ParseWorkbook(EpplusParser, "fitbit_export_20180109.xlsx");

        AssertParsedData(reference);
        AssertParsedData(candidate);
        AssertEquivalent(reference, candidate);
    }

    [Fact]
    public void OpenXml_matches_excel_data_reader_for_xlsx_sample()
    {
        var reference = ParseWorkbook(ExcelDataReaderParser, "fitbit_export_20180109.xlsx");
        var candidate = ParseWorkbook(OpenXmlParser, "fitbit_export_20180109.xlsx");

        AssertParsedData(reference);
        AssertParsedData(candidate);
        AssertEquivalent(reference, candidate);
    }

    private static FitbitData ParseWorkbook(IFitbitWorkbookParser parser, string fileName)
    {
        var samplePath = Path.Combine(AppContext.BaseDirectory, "Sample Upload", fileName);
        using var stream = File.OpenRead(samplePath);
        return parser.Parse(stream);
    }

    private static void AssertParsedData(FitbitData data)
    {
        Assert.True(data.TotalRowCount > 0);
        Assert.Empty(data.ErrorMessages);
        Assert.NotEmpty(data.BodyDataEntries);
        Assert.NotEmpty(data.ActivityDataEntries);
        Assert.NotEmpty(data.SleepDataEntries);
    }

    private static void AssertEquivalent(FitbitData expected, FitbitData actual)
    {
        Assert.Equal(expected.BodyDataEntries.Count, actual.BodyDataEntries.Count);
        Assert.Equal(expected.ActivityDataEntries.Count, actual.ActivityDataEntries.Count);
        Assert.Equal(expected.SleepDataEntries.Count, actual.SleepDataEntries.Count);

        for (var i = 0; i < expected.BodyDataEntries.Count; i++)
        {
            var expectedEntry = expected.BodyDataEntries[i];
            var actualEntry = actual.BodyDataEntries[i];

            Assert.Equal(expectedEntry.Date, actualEntry.Date);
            AssertNearlyEqual(expectedEntry.Weight, actualEntry.Weight);
            AssertNearlyEqual(expectedEntry.Bmi, actualEntry.Bmi);
            AssertNearlyEqual(expectedEntry.Fat, actualEntry.Fat);
        }

        for (var i = 0; i < expected.ActivityDataEntries.Count; i++)
        {
            var expectedEntry = expected.ActivityDataEntries[i];
            var actualEntry = actual.ActivityDataEntries[i];

            Assert.Equal(expectedEntry.Date, actualEntry.Date);
            Assert.Equal(expectedEntry.CaloriesBurned, actualEntry.CaloriesBurned);
            Assert.Equal(expectedEntry.Steps, actualEntry.Steps);
            AssertNearlyEqual(expectedEntry.Distance, actualEntry.Distance);
            Assert.Equal(expectedEntry.Floors, actualEntry.Floors);
            Assert.Equal(expectedEntry.MinutesSedentary, actualEntry.MinutesSedentary);
            Assert.Equal(expectedEntry.MinutesLightlyActive, actualEntry.MinutesLightlyActive);
            Assert.Equal(expectedEntry.MinutesFairlyActive, actualEntry.MinutesFairlyActive);
            Assert.Equal(expectedEntry.MinutesVeryActive, actualEntry.MinutesVeryActive);
            Assert.Equal(expectedEntry.ActivityCalories, actualEntry.ActivityCalories);
        }

        for (var i = 0; i < expected.SleepDataEntries.Count; i++)
        {
            var expectedEntry = expected.SleepDataEntries[i];
            var actualEntry = actual.SleepDataEntries[i];

            Assert.Equal(expectedEntry.StartTime, actualEntry.StartTime);
            Assert.Equal(expectedEntry.EndTime, actualEntry.EndTime);
            Assert.Equal(expectedEntry.MinutesAsleep, actualEntry.MinutesAsleep);
            Assert.Equal(expectedEntry.MinutesAwake, actualEntry.MinutesAwake);
            Assert.Equal(expectedEntry.NumberOfAwakenings, actualEntry.NumberOfAwakenings);
            Assert.Equal(expectedEntry.TimeInBed, actualEntry.TimeInBed);
            Assert.Equal(expectedEntry.MinutesRemSleep, actualEntry.MinutesRemSleep);
            Assert.Equal(expectedEntry.MinutesLightSleep, actualEntry.MinutesLightSleep);
            Assert.Equal(expectedEntry.MinutesDeepSleep, actualEntry.MinutesDeepSleep);
        }
    }

    private static void AssertNearlyEqual(double expected, double actual, double tolerance = 0.0001d)
    {
        Assert.InRange(actual, expected - tolerance, expected + tolerance);
    }
}
