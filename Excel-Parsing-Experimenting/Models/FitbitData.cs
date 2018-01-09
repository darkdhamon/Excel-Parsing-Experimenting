using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
// ReSharper disable UnusedMember.Global

namespace Excel_Parsing_Experimenting.Models
{
    public class FitbitData
    {
        [NotMapped]
        private List<string> _errorMessages;

        private List<BodyDataEntry> _bodyDataEntries;
        private List<SleepDataEntry> _sleepDataEntries;
        private List<ActivityDataEntry> _activityDataEntries;

        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }

        public List<BodyDataEntry> BodyDataEntries
        {
            get => _bodyDataEntries??(_bodyDataEntries= new List<BodyDataEntry>());
            set => _bodyDataEntries = value;
        }

        public List<SleepDataEntry> SleepDataEntries
        {
            get => _sleepDataEntries??(_sleepDataEntries = new List<SleepDataEntry>());
            set => _sleepDataEntries = value;
        }

        public List<ActivityDataEntry> ActivityDataEntries
        {
            get => _activityDataEntries??(_activityDataEntries = new List<ActivityDataEntry>());
            set => _activityDataEntries = value;
        }

        [NotMapped]
        public List<string> ErrorMessages => _errorMessages??(_errorMessages = new List<string>());
    }

    public class BodyDataEntry
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }
        public DateTime Date { get; set; }
        public double Weight { get; set; }
        public double BMI { get; set; }
        public double Fat { get; set; }
    }

    public class ActivityDataEntry
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
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
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
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