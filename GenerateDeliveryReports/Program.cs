using GenerateDeliveryReports.Components;
using GenerateDeliveryReports.Data.Concrete;
using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Data.Services;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Options;
using Serilog;
using System.Diagnostics;

var builder = WebApplication.CreateBuilder(args);

// Kestrel: HTTP only, listen on all interfaces so other machines on the LAN can connect
builder.WebHost.UseUrls("http://*:5158");

// Serilog
var logDir = Path.Combine(AppContext.BaseDirectory, "LogFiles");
if (!Directory.Exists(logDir))
    Directory.CreateDirectory(logDir);
builder.Host.UseSerilog((ctx, cfg) =>
    cfg.WriteTo.File(Path.Combine(logDir, "log.txt"), rollingInterval: RollingInterval.Day)
       .WriteTo.Console());

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

// HTTP only -- no HTTPS redirect

app.UseAntiforgery();

// Serve dynamically generated files (chart images, PDFs) from wwwroot/downloads
var downloadsPath = Path.Combine(AppContext.BaseDirectory, "wwwroot", "downloads");
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

app.MapGet("/api/worker-summary", async (IOptions<AppSettings> options) =>
{
    var path = options.Value.WorkerSummaryFilePath;
    if (string.IsNullOrWhiteSpace(path))
        path = Path.Combine(AppContext.BaseDirectory, "wwwroot", "worker-summary.html");

    if (!File.Exists(path))
        return Results.NotFound();

    var html = await File.ReadAllTextAsync(path);
    return Results.Content(html, "text/html");
});

// Auto-launch browser on the local machine when the server starts
app.Lifetime.ApplicationStarted.Register(() =>
{
    try
    {
        Process.Start(new ProcessStartInfo("http://localhost:5158") { UseShellExecute = true });
    }
    catch { /* non-critical */ }
});

app.Run();
