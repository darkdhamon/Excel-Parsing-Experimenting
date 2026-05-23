# Excel Parsing Experimenting

`Excel Parsing Experimenting` is now an ASP.NET Core MVC application targeting `.NET 10`.
The repository keeps the original goal of comparing Excel parsing approaches against Fitbit exports, but the implementation has been rebuilt on a modern web stack with fully working parser paths for:

- `ExcelDataReader`
- `EPPlus`
- `Open XML SDK`

## What Changed

- Migrated from ASP.NET MVC 5 on .NET Framework 4.6.1 to ASP.NET Core MVC on `.NET 10`
- Replaced the legacy project system with SDK-style projects
- Added a test project with sample-driven parser regression coverage
- Rebuilt the UI while preserving the upload-and-compare workflow
- Completed full worksheet support for `Body`, `Activities`, and `Sleep` across all three parser implementations

## Current Behavior

The home page presents three upload cards, one for each parser library.
Each upload route maps workbook content into the same shared Fitbit model and renders:

- Body data
- Activity data
- Sleep data

The included sample files currently produce equivalent results across the supported parser routes:

- `31` body rows
- `31` activity rows
- `26` sleep rows

## Parser Notes

### ExcelDataReader

- Supports `.xls` and `.xlsx`
- Used as the broadest compatibility baseline

### EPPlus

- Supports `.xlsx`
- Fully implemented for body, activity, and sleep sheets
- Configured for noncommercial use in this sample app

### Open XML SDK

- Supports `.xlsx`
- Fully implemented for body, activity, and sleep sheets
- Handles shared strings and date-formatted cells directly

## Repository Layout

- [Excel-Parsing-Experimenting.sln](/C:/GitHub/Excel-Parsing-Experimenting/Excel-Parsing-Experimenting.sln)
- [Excel-Parsing-Experimenting/Program.cs](/C:/GitHub/Excel-Parsing-Experimenting/Excel-Parsing-Experimenting/Program.cs)
- [Excel-Parsing-Experimenting/Controllers/UploadController.cs](/C:/GitHub/Excel-Parsing-Experimenting/Excel-Parsing-Experimenting/Controllers/UploadController.cs)
- [Excel-Parsing-Experimenting/Services/FitbitParsing](/C:/GitHub/Excel-Parsing-Experimenting/Excel-Parsing-Experimenting/Services/FitbitParsing)
- [Excel-Parsing-Experimenting/ViewModels](/C:/GitHub/Excel-Parsing-Experimenting/Excel-Parsing-Experimenting/ViewModels)
- [Excel-Parsing-Experimenting.Tests/ParserIntegrationTests.cs](/C:/GitHub/Excel-Parsing-Experimenting/Excel-Parsing-Experimenting.Tests/ParserIntegrationTests.cs)
- [Sample Upload/fitbit_export_20180109.xls](/C:/GitHub/Excel-Parsing-Experimenting/Sample%20Upload/fitbit_export_20180109.xls)
- [Sample Upload/fitbit_export_20180109.xlsx](/C:/GitHub/Excel-Parsing-Experimenting/Sample%20Upload/fitbit_export_20180109.xlsx)

## Requirements

- `.NET SDK 10.0.300` or newer compatible `.NET 10` SDK
- Windows if you want to follow the original repo conventions and sample workflow
- Visual Studio 2026 Insider Preview if you want to open the solution in Visual Studio

## Running Locally

From the repository root:

```powershell
dotnet restore
dotnet run --project .\Excel-Parsing-Experimenting\Excel-Parsing-Experimenting.csproj
```

Then open the local URL shown by Kestrel and upload one of the sample Fitbit exports.

## Running Tests

```powershell
dotnet test .\Excel-Parsing-Experimenting.sln
```

## EPPlus Licensing

This repository configures EPPlus for noncommercial use:

```csharp
ExcelPackage.License.SetNonCommercialPersonal("Excel Parsing Experimenting");
```

If you intend to use this application in a commercial context, replace that setup with the appropriate EPPlus commercial license configuration.
