namespace Excel_Parsing_Experimenting.Models;

public sealed class FitbitData
{
    public List<BodyDataEntry> BodyDataEntries { get; } = [];

    public List<ActivityDataEntry> ActivityDataEntries { get; } = [];

    public List<SleepDataEntry> SleepDataEntries { get; } = [];

    public List<string> ErrorMessages { get; } = [];

    public int TotalRowCount => BodyDataEntries.Count + ActivityDataEntries.Count + SleepDataEntries.Count;

    public bool HasData => TotalRowCount > 0;
}

public sealed record BodyDataEntry(DateTime Date, double Weight, double Bmi, double Fat);

public sealed record ActivityDataEntry(
    DateTime Date,
    int CaloriesBurned,
    int Steps,
    double Distance,
    int Floors,
    int MinutesSedentary,
    int MinutesLightlyActive,
    int MinutesFairlyActive,
    int MinutesVeryActive,
    int ActivityCalories);

public sealed record SleepDataEntry(
    DateTime StartTime,
    DateTime EndTime,
    int MinutesAsleep,
    int MinutesAwake,
    int NumberOfAwakenings,
    int TimeInBed,
    int? MinutesRemSleep,
    int? MinutesLightSleep,
    int? MinutesDeepSleep);
