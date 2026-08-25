using GenerateDeliveryReports.Components;
using GenerateDeliveryReports.Data.Concrete;
using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Data.Services;
using GenerateDeliveryReports.Identity;
using GenerateDeliveryReports.Models;
using System.Diagnostics;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using Serilog;
using Microsoft.AspNetCore.Diagnostics;

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
builder.Services.AddScoped<SprintReportsAvailabilityService>();

// Register email export service
builder.Services.AddScoped<GenerateDeliveryReports.Services.EmailMessageGenerator>();
if (builder.Environment.IsDevelopment())
{
    builder.Services.AddSingleton<MeetingMinutesService>();
}


builder.Services.AddHttpClient<ClaudeApiClient>();
builder.Services.AddSingleton<BriefingCache>();
builder.Services.AddScoped<BriefingGenerator>();

// Entra ID (Microsoft identity platform) authentication -- kept registered for whenever it's
// usable again, but no longer owns the default scheme or the login requirement (see local
// Identity below): the org's tenant-restriction policy blocks sign-in via a personal/external
// tenant, so Entra can't be the enforced mechanism right now.
if (azureAdConfigured)
{
    builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
        .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"));

    // The Identity UI login page lists external providers using the scheme's display name --
    // rename it from the default "OpenIdConnect" to something a user actually recognizes.
    builder.Services.Configure<AuthenticationOptions>(options =>
    {
        if (options.SchemeMap.TryGetValue(OpenIdConnectDefaults.AuthenticationScheme, out var schemeBuilder))
            schemeBuilder.DisplayName = "Login with SSO (Entra)";
    });

    // Force an explicit account picker on every sign-in. Without this, Windows/Edge's account
    // broker (WAM) can silently substitute a different cached Microsoft account than the one
    // the user intends, which surfaces as AADSTS50197 ("could not find the user").
    //
    // SignInScheme: the "Login with SSO (Entra)" button on the Identity login page is Identity
    // UI's own external-provider button -- it posts to /Identity/Account/ExternalLogin, whose
    // callback reads the completed login from the IdentityConstants.ExternalScheme cookie via
    // SignInManager.GetExternalLoginInfoAsync(). AddMicrosoftIdentityWebApp defaults this scheme's
    // SignInScheme to its own "Cookies" scheme instead (it's built to be a standalone auth
    // solution), so that lookup found nothing and failed with "Error loading external login
    // information." Pointing SignInScheme at ExternalScheme is also required for
    // ExternalLoginAuthorizationHandler: it checks UserManager.GetLoginsAsync (AspNetUserLogins
    // rows), which only get populated by Identity's own ExternalLogin flow completing --
    // bypassing that flow entirely would've broken Team Logins access and Register lockdown for
    // real Entra sign-ins too.
    builder.Services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, options =>
    {
        options.Prompt = "select_account";
        options.SignInScheme = IdentityConstants.ExternalScheme;
    });

    builder.Services.AddControllersWithViews()
        .AddMicrosoftIdentityUI();
}

// Local sign-in (ASP.NET Core Identity, SQLite-backed) -- this is the enforced auth mechanism.
// Registered unconditionally: unlike Entra it needs no external tenant, so it's always available.
builder.Services.AddDbContext<ApplicationDbContext>((sp, options) =>
{
    var settings = sp.GetRequiredService<IOptions<AppSettings>>().Value;
    options.UseSqlite(settings.IdentityConnectionString);
});

builder.Services.AddDefaultIdentity<ApplicationUser>(options =>
{
    options.SignIn.RequireConfirmedAccount = false; // no email service wired up for confirmation links
})
    .AddRoles<IdentityRole>()
    .AddEntityFrameworkStores<ApplicationDbContext>();

builder.Services.AddRazorPages();

// Scoped, not Singleton -- it depends on UserManager, which is itself Scoped (avoids capturing
// a scoped DbContext-backed dependency inside a singleton).
builder.Services.AddScoped<IAuthorizationHandler, ExternalLoginAuthorizationHandler>();

builder.Services.AddAuthorization(options =>
{
    options.FallbackPolicy = options.DefaultPolicy; // require an authenticated user everywhere by default
    options.AddPolicy("RequireExternalLogin", policy => policy.Requirements.Add(new ExternalLoginRequirement()));
});

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

app.UseAuthentication();
app.UseAuthorization();

// Explicit middleware, not a Razor Pages authorization convention: Identity's own setup marks
// the whole /Account folder [AllowAnonymous] (it has to -- otherwise nobody could reach Login
// at all), and a page-level AuthorizeAreaPage convention added on top of that did NOT actually
// take effect (verified: Register loaded with no authorization check running at all in the
// logs). This reuses the same RequireExternalLogin policy Team Logins uses, applied directly,
// so self-registration is genuinely disabled -- accounts are only ever created via that page.
app.Use(async (context, next) =>
{
    if (context.Request.Path.StartsWithSegments("/Identity/Account/Register"))
    {
        var authorizationService = context.RequestServices.GetRequiredService<IAuthorizationService>();
        var result = await authorizationService.AuthorizeAsync(context.User, "RequireExternalLogin");
        if (!result.Succeeded)
        {
            context.Response.Redirect("/Identity/Account/AccessDenied");
            return;
        }
    }
    await next();
});

app.UseAntiforgery();

// Initialize the database safely (handles fresh installs and schema updates on existing databases).
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    var logger = scope.ServiceProvider.GetRequiredService<ILogger<Program>>();
    DbInitializer.InitializeSafely(db, logger);
}

// Serve dynamically generated files (chart images, PDFs) from wwwroot/downloads
var downloadsPath = app.Services.GetRequiredService<IOptions<AppSettings>>().Value.TempPath;

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(downloadsPath),
    RequestPath = "/downloads"
});

// AllowAnonymous is required here: the global FallbackPolicy requires authentication on every
// endpoint by default, which would otherwise block the CSS/JS the Login page itself needs to
// render (a chicken-and-egg problem -- you need those assets to be able to log in at all).
app.MapStaticAssets().AllowAnonymous();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.MapRazorPages(); // Identity UI's Login/Register/Logout/Manage pages under /Identity/Account/...

if (azureAdConfigured)
{
    app.MapControllers(); // Microsoft.Identity.Web.UI's Account/SignIn, Account/SignOut endpoints
}

app.MapGet("/api/worker-summary", async (IOptions<AppSettings> options, ILogger<Program> logger) =>
{
    var path = options.Value.WorkerRunHistoryFilePath;
    if (string.IsNullOrWhiteSpace(path))
        path = Path.Combine(AppContext.BaseDirectory, "wwwroot", "CommonFiles", "WorkerHistory", "run-workerhistory.json");

    // Convert JSON path to HTML path
    path = path.Replace(".json", ".html");

    logger.LogInformation($"Worker summary endpoint - Looking for file at: {path}");
    logger.LogInformation($"File exists: {File.Exists(path)}");
    logger.LogInformation($"AppContext.BaseDirectory: {AppContext.BaseDirectory}");

    if (!File.Exists(path))
    {
        logger.LogWarning($"Worker summary file not found at: {path}");
        return Results.NotFound();
    }

    var html = await File.ReadAllTextAsync(path);
    return Results.Content(html, "text/html");
});

app.MapGet("/api/sprint-dashboard", async (IOptions<AppSettings> options) =>
{
    var path = options.Value.ResolvedSprintDashboardHtmlPath;
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

app.UseExceptionHandler(errorApp =>
{
    errorApp.Run(async context =>
    {
        var exceptionHandlerPathFeature = context.Features.Get<IExceptionHandlerPathFeature>();
        var exception = exceptionHandlerPathFeature?.Error;
        
        context.Response.ContentType = "text/html";
        
        var html = $@"
            <html>
            <body style='font-family:Arial;padding:20px;'>
            <h1>ERROR in CSAT Report</h1>
            <p><strong>{exception?.GetType().Name}</strong></p>
            <p>{exception?.Message}</p>
            <pre style='background:#f0f0f0;padding:10px;overflow:auto;'>{exception?.StackTrace}</pre>
            </body>
            </html>";
        
        await context.Response.WriteAsync(html);
    });
});


app.Run();
