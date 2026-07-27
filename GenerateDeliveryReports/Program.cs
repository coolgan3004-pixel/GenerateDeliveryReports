using GenerateDeliveryReports.Components;
using GenerateDeliveryReports.Data.Concrete;
using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Data.Services;
using GenerateDeliveryReports.Models;
using System.Diagnostics;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using Serilog;

var builder = WebApplication.CreateBuilder(args);

// Entra ID login is only enforced once a tenant/app registration has actually been configured
// (TenantId + ClientId). Until then the app runs exactly as before -- this lets Entra be rolled
// out without locking out local development that hasn't set up an app registration yet.
var azureAdConfigured = !string.IsNullOrWhiteSpace(builder.Configuration["AzureAd:TenantId"])
    && !string.IsNullOrWhiteSpace(builder.Configuration["AzureAd:ClientId"]);

// Kestrel: HTTP only on a fixed port — only for the deployed (Production) instance running
// standalone via Task Scheduler. Dev runs use launchSettings.json to avoid port conflicts.
// Under IIS in-process hosting, UseUrls is skipped entirely (APP_POOL_ID is set by IIS).
var isStandaloneProduction = builder.Environment.IsProduction()
    && string.IsNullOrEmpty(Environment.GetEnvironmentVariable("APP_POOL_ID"));

if (isStandaloneProduction)
{
    builder.WebHost.UseUrls("http://*:5158");
}

// Serilog — log directory comes from CommonFolderPath when configured, falls back to exe folder
builder.Host.UseSerilog((ctx, cfg) =>
{
    var settings = ctx.Configuration.GetSection("AppSettings").Get<AppSettings>() ?? new AppSettings();
    cfg.WriteTo.File(Path.Combine(settings.LogFilesPath, "log.txt"), rollingInterval: RollingInterval.Day)
       .WriteTo.Console();
});

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Bind AppSettings from configuration
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));

// Register data layer services
builder.Services.AddSingleton<IDataProcessor, DataProcessor>();
builder.Services.AddScoped<IEmailService, OutlookEmailService>();
builder.Services.AddScoped<SprintReportService>();
builder.Services.AddScoped<CsatService>();
builder.Services.AddScoped<SprintDashboardService>();
builder.Services.AddScoped<SprintReportStatusService>();
builder.Services.AddSingleton<MeetingMinutesService>();

builder.Services.AddHttpClient<ClaudeApiClient>();
builder.Services.AddSingleton<BriefingCache>();
builder.Services.AddScoped<BriefingGenerator>();

// Entra ID (Microsoft identity platform) authentication
if (azureAdConfigured)
{
    builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
        .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"));

    // Force an explicit account picker on every sign-in. Without this, Windows/Edge's account
    // broker (WAM) can silently substitute a different cached Microsoft account than the one
    // the user intends, which surfaces as AADSTS50197 ("could not find the user").
    builder.Services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, options =>
    {
        options.Prompt = "select_account";
    });

    builder.Services.AddControllersWithViews()
        .AddMicrosoftIdentityUI();

    builder.Services.AddAuthorization(options =>
    {
        options.FallbackPolicy = options.DefaultPolicy; // require an authenticated user everywhere by default
    });
}
else
{
    builder.Services.AddAuthorization();
}

builder.Services.AddCascadingAuthenticationState();

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

if (azureAdConfigured)
{
    app.UseAuthentication();
}
app.UseAuthorization();

app.UseAntiforgery();

// Serve dynamically generated files (chart images, PDFs) from wwwroot/downloads
var downloadsPath = app.Services.GetRequiredService<IOptions<AppSettings>>().Value.TempPath;

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(downloadsPath),
    RequestPath = "/downloads"
});

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

if (azureAdConfigured)
{
    app.MapControllers(); // Microsoft.Identity.Web.UI's Account/SignIn, Account/SignOut endpoints
}

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

app.MapGet("/api/sprint-dashboard", async (IOptions<AppSettings> options) =>
{
    var path = options.Value.SprintDashboardHtmlPath;
    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        return Results.NotFound();

    var html = await File.ReadAllTextAsync(path);
    return Results.Content(html, "text/html");
});

// Auto-launch browser on the local machine when the server starts. Only for the standalone
// Production deployment, which is the only scenario where port 5158 (hardcoded above) is
// actually correct -- in Development the real port comes from launchSettings.json (5159), and
// launchBrowser:true there already opens the browser at the right address.
if (isStandaloneProduction)
{
    app.Lifetime.ApplicationStarted.Register(() =>
    {
        try
        {
            Process.Start(new ProcessStartInfo("http://localhost:5158") { UseShellExecute = true });
        }
        catch { /* non-critical */ }
    });
}

app.Run();
