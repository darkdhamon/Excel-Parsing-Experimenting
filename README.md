# Excel Parsing Experimenting

`Excel Parsing Experimenting` is a small ASP.NET MVC 5 application used to compare different ways of reading Excel workbooks that were exported from Fitbit. The app accepts a workbook upload, parses data into a shared view model, and renders the results as HTML tables for quick inspection.

## Project Summary

- Framework: ASP.NET MVC 5 on .NET Framework 4.6.1
- App type: classic ASP.NET web application
- Primary goal: experiment with Excel parsing libraries against Fitbit export files
- Persistence: none; uploaded data is parsed in-memory and displayed back to the user
- Sample files included: `.xls` and `.xlsx` Fitbit exports

## What The App Does

The home page presents three separate upload forms. Each form sends the same workbook through a different parser implementation:

1. `ExcelDataReader`
2. `EPPlus`
3. `Open XML SDK`

The parsed results are mapped into a shared `FitbitData` model and displayed in three sections:

- Body data
- Activity data
- Sleep data

## Current Implementation Status

The codebase is intentionally experimental, and the three approaches are not equally complete.

### 1. ExcelDataReader

Controller action: `UploadController.ImportFitbitDataEDR`

- Supports `.xls` and `.xlsx`
- Parses `Body`, `Activities`, and `Sleep` worksheets
- Skips rows that fail parsing, which is how header rows are ignored
- This is the most complete implementation in the repository

### 2. EPPlus

Controller action: `UploadController.ImportFitbitDataEep`

- Supports `.xlsx`
- Currently parses only the `Body` worksheet
- `Activities` and `Sleep` handling are stubbed but not implemented
- Adds parsing exceptions to the view model error list

### 3. Open XML SDK

Controller action: `UploadController.ImportFitbitDataOpenXml`

- Supports `.xlsx`
- Currently attempts to parse only the `Body` worksheet
- Works at a much lower level than the other approaches
- The implementation is more brittle and incomplete than the `ExcelDataReader` path

## Repository Layout

- [Excel-Parsing-Experimenting.sln](./Excel-Parsing-Experimenting.sln)
- [Excel-Parsing-Experimenting/Controllers/UploadController.cs](./Excel-Parsing-Experimenting/Controllers/UploadController.cs)
- [Excel-Parsing-Experimenting/Models/FitbitData.cs](./Excel-Parsing-Experimenting/Models/FitbitData.cs)
- [Excel-Parsing-Experimenting/Views/Home/Index.cshtml](./Excel-Parsing-Experimenting/Views/Home/Index.cshtml)
- [Excel-Parsing-Experimenting/Views/Upload/ImportFitbitData.cshtml](./Excel-Parsing-Experimenting/Views/Upload/ImportFitbitData.cshtml)
- [Sample Upload/fitbit_export_20180109.xls](./Sample%20Upload/fitbit_export_20180109.xls)
- [Excel-Parsing-Experimenting/Content/fitbit_export_20180109.xlsx](./Excel-Parsing-Experimenting/Content/fitbit_export_20180109.xlsx)

## Dependencies

Key NuGet packages used by the project:

- `ExcelDataReader` `3.3.0`
- `ExcelDataReader.DataSet` `3.3.0`
- `EPPlus` `4.1.1`
- `Open-XML-SDK` `2.7.2`
- `Microsoft.AspNet.Mvc` `5.2.3`

The project uses `packages.config`, so NuGet restore is required before the first build if the `packages` directory is not already present.

## Running The Project

### Prerequisites

- Windows
- A Visual Studio installation that supports ASP.NET MVC 5 web applications targeting .NET Framework 4.6.1
- .NET Framework 4.6.1 Developer Pack / Targeting Pack
- ASP.NET web development support / IIS Express

### Open In Visual Studio

Open the solution in your compatible Visual Studio installation:

```powershell
devenv.exe .\Excel-Parsing-Experimenting.sln
```

### Restore And Build From The Command Line

If you prefer the command line, run these from a Developer PowerShell / Developer Command Prompt for Visual Studio:

```powershell
MSBuild.exe .\Excel-Parsing-Experimenting.sln -t:Restore -p:RestorePackagesConfig=true
MSBuild.exe .\Excel-Parsing-Experimenting.sln /p:Configuration=Debug
```

### Run Locally

1. Restore NuGet packages.
2. Build the solution.
3. Start the web project from Visual Studio using IIS Express.
4. Open the home page and upload a Fitbit export workbook.
5. Compare the output produced by each parser path.

## Sample Data

The repo includes Fitbit export samples that are useful for quick manual testing:

- [Sample Upload/fitbit_export_20180109.xls](./Sample%20Upload/fitbit_export_20180109.xls)
- [Excel-Parsing-Experimenting/Content/fitbit_export_20180109.xlsx](./Excel-Parsing-Experimenting/Content/fitbit_export_20180109.xlsx)

## Known Limitations

- The three parser implementations are not functionally equivalent.
- `EPPlus` and `Open XML SDK` do not currently parse activity and sleep worksheets.
- Error handling is lightweight and mostly oriented around skipping headers or recording exception messages.
- The UI is a basic experiment page rather than a polished end-user workflow.
- The application does not store uploads or parsed records.

## Recommended Next Steps

If you want to keep evolving this project, the highest-value improvements would be:

1. Make all three parser implementations cover the same worksheet set.
2. Extract parsing logic into separate services so library behavior can be compared more cleanly.
3. Add automated tests against the included sample exports.
4. Improve validation and user-facing error messages for malformed files.
5. Add performance measurements if the goal is benchmarking as well as correctness.
