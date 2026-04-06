using GenerateDeliveryReports.Components;
using GenerateDeliveryReports.Data.Concrete;
using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Data.Services;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.FileProviders;
using Serilog;

var builder = WebApplication.CreateBuilder(args);

// Serilog
builder.Host.UseSerilog((ctx, cfg) =>
    cfg.WriteTo.File("LogFiles/log.txt", rollingInterval: RollingInterval.Day));

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Bind AppSettings from configuration
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));

// Register data layer services
builder.Services.AddSingleton<IDataProcessor, DataProcessor>();
builder.Services.AddScoped<SprintReportService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseHttpsRedirection();

app.UseAntiforgery();

// Serve dynamically generated files (chart images, PDFs) from wwwroot/downloads
var downloadsPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "downloads");
if (!Directory.Exists(downloadsPath))
    Directory.CreateDirectory(downloadsPath);

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(downloadsPath),
    RequestPath = "/downloads"
});

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
