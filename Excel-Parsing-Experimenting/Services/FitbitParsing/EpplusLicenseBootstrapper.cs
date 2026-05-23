using System.Threading;
using OfficeOpenXml;

namespace Excel_Parsing_Experimenting.Services.FitbitParsing;

internal static class EpplusLicenseBootstrapper
{
    private static int _isConfigured;

    public static void EnsureConfigured()
    {
        if (Interlocked.Exchange(ref _isConfigured, 1) == 1)
        {
            return;
        }

        ExcelPackage.License.SetNonCommercialPersonal("Excel Parsing Experimenting");
    }
}
