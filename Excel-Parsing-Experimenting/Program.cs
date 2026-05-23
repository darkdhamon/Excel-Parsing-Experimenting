using System.Text;
using Excel_Parsing_Experimenting.Services.FitbitParsing;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddSingleton<FitbitWorkbookMapper>();
builder.Services.AddSingleton<IFitbitWorkbookParser, ExcelDataReaderFitbitWorkbookParser>();
builder.Services.AddSingleton<IFitbitWorkbookParser, EpplusFitbitWorkbookParser>();
builder.Services.AddSingleton<IFitbitWorkbookParser, OpenXmlFitbitWorkbookParser>();
builder.Services.AddSingleton<FitbitWorkbookParserCatalog>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseRouting();

app.UseAuthorization();

app.MapStaticAssets();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}")
    .WithStaticAssets();


app.Run();
