using GenerateDeliveryReports.Data.Concrete;
using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Models;
using GenerateDeliveryReports.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .WriteTo.File("LogFiles/workerlog.txt", rollingInterval: RollingInterval.Day)
    .WriteTo.Console()
    .CreateLogger();

try
{
    Log.Information("GenerateDeliveryReports Worker starting.");

    var builder = Host.CreateApplicationBuilder(args);
    builder.Configuration.AddJsonFile("workeremailsettings.json", optional: true, reloadOnChange: false);

    builder.Services.AddSerilog();
    builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));
    builder.Services.AddSingleton<IDataProcessor, DataProcessor>();
    builder.Services.AddSingleton<ReportWorker>();

    var host = builder.Build();

    var worker = host.Services.GetRequiredService<ReportWorker>();
    await worker.RunOnceAsync(CancellationToken.None);

    Log.Information("GenerateDeliveryReports Worker completed.");
}
catch (Exception ex)
{
    Log.Fatal(ex, "Worker terminated unexpectedly.");
    return 1;
}
finally
{
    await Log.CloseAndFlushAsync();
}

return 0;
