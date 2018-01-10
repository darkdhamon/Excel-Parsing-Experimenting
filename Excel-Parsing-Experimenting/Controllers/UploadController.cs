using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Excel_Parsing_Experimenting.Models;
using OfficeOpenXml;

namespace Excel_Parsing_Experimenting.Controllers
{
    public class UploadController : Controller
    {
        public ActionResult ImportFitbitDataEDR(HttpPostedFileBase file)
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
            return View("ImportFitbitData",viewModel);
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
                entry.Distance = double.Parse(worksheetRow[3].ToString());
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
                entry.Weight = double.Parse(worksheetRow[1].ToString());
                entry.BMI = double.Parse(worksheetRow[2].ToString());
                entry.Fat = double.Parse(worksheetRow[3].ToString());
            }
            catch
            {
                // ignore error as this is likely a header row.
                return null;
            }
            return entry;
        }

        public ActionResult ImportFitbitDataEep(HttpPostedFileBase file)
        {
            var viewModel = new FitbitData();
            if (file == null || file.ContentLength <= 0) return View("ImportFitbitData", viewModel);
            //var guid = Guid.NewGuid();
            //var targetfolder = HttpContext.Server.MapPath($"~/uploads/excelDoc-{guid}");
            //if (!Directory.Exists(targetfolder))
            //    Directory.CreateDirectory(targetfolder);
            //var targetpath = Path.Combine(targetfolder, file.FileName);
            //file.SaveAs(targetpath);
            //var fileinfo = new FileInfo(targetpath);
            using (var package = new ExcelPackage(file.InputStream))
            {
                var workbook = package.Workbook;
                foreach (var worksheet in workbook.Worksheets)
                {

                    switch (worksheet.Name)
                    {
                        case "Body":
                            for (var i = 2; i <= worksheet.Dimension.End.Row; i++)
                            {
                                var entry = new BodyDataEntry();
                                try
                                {
                                    entry.Date = Convert.ToDateTime(worksheet.Cells[i, 1].Value.ToString());
                                    entry.Weight = Convert.ToDouble(worksheet.Cells[i, 2].Value.ToString());
                                    entry.BMI = Convert.ToDouble(worksheet.Cells[i, 3].Value.ToString());
                                    entry.Fat = Convert.ToDouble(worksheet.Cells[i, 4].Value.ToString());
                                }
                                catch (Exception exception)
                                {
                                    viewModel.ErrorMessages.Add(exception.Message);
                                }
                                viewModel.BodyDataEntries.Add(entry);
                            }
                            break;
                        case "Activities":
                                
                            break;
                        case "Sleep":
                                
                            break;
                        default:
                            viewModel.ErrorMessages.Add($"Unknown worksheet name: {worksheet.Name}");
                            break;
                    }
                }
            }
            return View("ImportFitbitData", viewModel);
        }

        public ActionResult ImportFitbitDataOpenXml(HttpPostedFileBase file)
        {
            var viewModel = new FitbitData();
            if (file == null || file.ContentLength <= 0) return View("ImportFitbitData", viewModel);
            try
            {
                using (var document = SpreadsheetDocument.Open(file.InputStream, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var workbook = workbookPart.Workbook;
                    var worksheets = workbook.Descendants<Sheet>().ToList();
                    foreach (var sheet in worksheets)
                    {
                        var worksheet = ((WorksheetPart) workbookPart.GetPartById(sheet.Id)).Worksheet;
                        var sheetdata = (SheetData) worksheet.ChildElements.GetItem(4); // 4 is the sheet data...
                        switch (sheet.Name.ToString())
                        {
                            case "Body":
                                foreach (var openXmlElement in sheetdata.ChildElements)
                                {
                                    var row = (Row) openXmlElement;
                                    var entry = new BodyDataEntry();
                                    try
                                    {
                                        foreach (var openXmlElement1 in row.ChildElements)
                                        {
                                            var cell = (Cell) openXmlElement1;
                                            SharedStringItem celltext = null;
                                            if (cell.DataType == CellValues.SharedString)
                                            {
                                                celltext = workbookPart.SharedStringTablePart.SharedStringTable
                                                    .Elements<SharedStringItem>()
                                                    .ElementAt(int.Parse(cell.InnerText));
                                            }
                                            if (celltext == null) continue;
                                            if (cell.CellReference.Value.Contains("A"))
                                            {
                                                entry.Date = Convert.ToDateTime(celltext.InnerText);
                                            }
                                            if (cell.CellReference.Value.Contains("B"))
                                            {
                                                entry.Weight = Convert.ToDouble(celltext.InnerText);
                                            }
                                            if (cell.CellReference.Value.Contains("C"))
                                            {
                                                entry.BMI = Convert.ToDouble(celltext.InnerText);
                                            }
                                            if (cell.CellReference.Value.Contains("D"))
                                            {
                                                entry.Fat = Convert.ToDouble(celltext.InnerText);
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        entry = null;
                                    }


                                    if (entry != null)
                                        viewModel.BodyDataEntries.Add(entry);
                                }

                                break;
                            case "Activities":

                                break;
                            case "Sleep":

                                break;
                            default:
                                viewModel.ErrorMessages.Add($"Unknown worksheet name: {sheet.Name}");
                                break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                viewModel.ErrorMessages.Add(e.Message);
            }
            return View("ImportFitbitData",viewModel);
        }
    }
}