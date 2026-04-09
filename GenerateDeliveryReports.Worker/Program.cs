using GenerateDeliveryReports.Data.Concrete;
using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Models;
using GenerateDeliveryReports.Worker;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .WriteTo.File("LogFiles/workerlog.txt", rollingInterval: RollingInterval.Day)
    .WriteTo.Console()
    .CreateLogger();

var builder = Host.CreateApplicationBuilder(args);
builder.Configuration.AddJsonFile("workeremailsettings.json", optional: false, reloadOnChange: true);

// Windows Service support
builder.Services.AddWindowsService(options =>
    options.ServiceName = "GenerateDeliveryReports Worker");

// Serilog
builder.Services.AddSerilog();

// Bind AppSettings
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));

// Register data layer
builder.Services.AddSingleton<IDataProcessor, DataProcessor>();

// Register worker
builder.Services.AddHostedService<ReportWorker>();

var host = builder.Build();
host.Run();
