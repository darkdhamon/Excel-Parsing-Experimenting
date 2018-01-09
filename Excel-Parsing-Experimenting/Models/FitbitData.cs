using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Excel_Parsing_Experimenting.Models
{
    public class FitbitData
    {
        public int ID { get; set; }
        public List<BodyDataEntry> BodyDataEntries { get; set; }
        public List<SleepDataEntry> SleepDataEntries { get; set; }
        public List<ActivityDataEntry> ActivityDataEntries { get; set; }
    }

    public class BodyDataEntry
    {
        public int ID { get; set; }
        public DateTime Date { get; set; }
        public double Weight { get; set; }
        public double BMI { get; set; }
        public double Fat { get; set; }
    }

    public class ActivityDataEntry
    {
        public int ID { get; set; }
        public DateTime Date { get; set; }
        public int CaloriesBurned { get; set; }
        public int Steps { get; set; }
        public double Distance { get; set; }
        public int Floors { get; set; }
        public int MinutesSedentary { get; set; }
        public int MinutesLightlyActive { get; set; }
        public int MinutesFairlyActive { get; set; }
        public int MinutesVeryActive { get; set; }
        public int ActivityCalories { get; set; }
    }

    public class SleepDataEntry
    {
        public int ID { get; set; } 
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public int MinutesAsleep { get; set; }
        public int MinutesAwake { get; set; }
        public int NumberOfAwaakenings { get; set; }
        public int TimeInBed { get; set; }
        public int MinutesREMSleep { get; set; }
        public int MinutesLightSleep { get; set; }
        public int MinutesDeepSleep { get; set; }
    }
}