using System;
using System.Data;
using System.Globalization;
using System.Web;
using System.Web.Mvc;
using ExcelDataReader;
using Excel_Parsing_Experimenting.Models;
using static System.Double;

namespace Excel_Parsing_Experimenting.Controllers
{
    public class UploadController : Controller
    {
        public ActionResult ImportFitbitData(HttpPostedFileBase file)
        {
            var viewModel = new FitbitData();
            if (file != null && file.ContentLength > 0)
            {
                var reader = ExcelReaderFactory.CreateReader(file.InputStream);
                var workbook  = reader.AsDataSet();
                foreach (DataTable worksheet in workbook.Tables)
                {
                    switch (worksheet.TableName)
                    {
                        case "Body":
                            foreach (DataRow worksheetRow in worksheet.Rows)
                            {
                                var entry = ParseBody(worksheetRow);
                                if (entry != null)
                                    viewModel.BodyDataEntries.Add(entry);
                            }
                            break;
                        case "Activities":
                            foreach (DataRow worksheetRow in worksheet.Rows)
                            {
                                var entry = ParseActivity(worksheetRow);
                                if (entry != null)
                                    viewModel.ActivityDataEntries.Add(entry);
                            }
                            break;
                        case "Sleep":
                            foreach (DataRow worksheetRow in worksheet.Rows)
                            {
                                var entry = ParseSleep(worksheetRow);
                                if (entry != null)
                                    viewModel.SleepDataEntries.Add(entry);
                            }
                            break;
                        default:
                            viewModel.ErrorMessages.Add($"Unknown worksheet name: {worksheet.TableName}");
                                break;
                    }
                    
                }
            }
            return View(viewModel);
        }

        private SleepDataEntry ParseSleep(DataRow worksheetRow)
        {
            var entry = new SleepDataEntry();
            try
            {
                entry.StartTime = DateTime.Parse(worksheetRow[0].ToString());
                entry.EndTime = DateTime.Parse(worksheetRow[1].ToString());
                entry.MinutesAsleep = int.Parse(worksheetRow[2].ToString(), NumberStyles.AllowThousands);
                entry.MinutesAwake = int.Parse(worksheetRow[3].ToString(), NumberStyles.AllowThousands);
                entry.NumberOfAwaakenings = int.Parse(worksheetRow[4].ToString(), NumberStyles.AllowThousands);
                entry.TimeInBed = int.Parse(worksheetRow[5].ToString(), NumberStyles.AllowThousands);
                entry.MinutesREMSleep = int.Parse(worksheetRow[6].ToString(), NumberStyles.AllowThousands);
                entry.MinutesLightSleep = int.Parse(worksheetRow[7].ToString(), NumberStyles.AllowThousands);
                entry.MinutesDeepSleep = int.Parse(worksheetRow[8].ToString(), NumberStyles.AllowThousands);
            }
            catch
            {
                // ignore error as this is likely a header row.
                return null;
            }
            return entry;
        }

        private ActivityDataEntry ParseActivity(DataRow worksheetRow)
        {
            var entry = new ActivityDataEntry();
            try
            {
                entry.Date = DateTime.Parse(worksheetRow[0].ToString());
                entry.CaloriesBurned = int.Parse(worksheetRow[1].ToString(),NumberStyles.AllowThousands);
                entry.Steps = int.Parse(worksheetRow[2].ToString(), NumberStyles.AllowThousands);
                entry.Distance = Parse(worksheetRow[3].ToString());
                entry.Floors = int.Parse(worksheetRow[4].ToString(), NumberStyles.AllowThousands);
                entry.MinutesSedentary = int.Parse(worksheetRow[5].ToString(), NumberStyles.AllowThousands);
                entry.MinutesLightlyActive = int.Parse(worksheetRow[6].ToString(), NumberStyles.AllowThousands);
                entry.MinutesFairlyActive = int.Parse(worksheetRow[7].ToString(), NumberStyles.AllowThousands);
                entry.MinutesVeryActive = int.Parse(worksheetRow[8].ToString(), NumberStyles.AllowThousands);
                entry.ActivityCalories = int.Parse(worksheetRow[9].ToString(), NumberStyles.AllowThousands);
            }
            catch
            {
                // ignore error as this is likely a header row.
                return null;
            }
            return entry;
        }

        private BodyDataEntry ParseBody(DataRow worksheetRow)
        {
            var entry = new BodyDataEntry();
            try
            {
                entry.Date = DateTime.Parse(worksheetRow[0].ToString());
                entry.Weight = Parse(worksheetRow[1].ToString());
                entry.BMI = Parse(worksheetRow[2].ToString());
                entry.Fat = Parse(worksheetRow[3].ToString());
            }
            catch
            {
                // ignore error as this is likely a header row.
                return null;
            }
            return entry;
        }
    }
}