using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Diagnostics;

namespace GenerateDeliveryReports.Data.Services;

public class SprintDashboardService
{
    private readonly AppSettings _settings;
    private readonly ILogger<SprintDashboardService> _logger;

    public SprintDashboardService(IOptions<AppSettings> options, ILogger<SprintDashboardService> logger)
    {
        _settings = options.Value;
        _logger = logger;
    }

    public (bool exists, DateTime? lastUpdated) GetDashboardInfo()
    {
        var path = _settings.SprintDashboardHtmlPath;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return (false, null);
        return (true, File.GetLastWriteTime(path));
    }

    public bool IsScriptConfigured =>
        !string.IsNullOrWhiteSpace(_settings.SprintDashboardScriptPath) &&
        File.Exists(_settings.SprintDashboardScriptPath);

    public async Task<(bool success, string message)> RunScriptAsync()
    {
        var scriptPath = _settings.SprintDashboardScriptPath;
        if (string.IsNullOrWhiteSpace(scriptPath) || !File.Exists(scriptPath))
            return (false, "Sprint dashboard script path not configured or file not found.");

        var psi = new ProcessStartInfo("python", $"\"{scriptPath}\"")
        {
            WorkingDirectory = Path.GetDirectoryName(scriptPath),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        try
        {
            using var proc = Process.Start(psi)!;
            var stdoutTask = proc.StandardOutput.ReadToEndAsync();
            var stderrTask = proc.StandardError.ReadToEndAsync();
            var completed = await Task.Run(() => proc.WaitForExit(120_000));

            if (!completed)
            {
                proc.Kill();
                return (false, "Script timed out after 2 minutes.");
            }

            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            if (proc.ExitCode != 0)
            {
                _logger.LogError("Sprint dashboard script exited {Code}: {Error}", proc.ExitCode, stderr);
                return (false, string.IsNullOrWhiteSpace(stderr) ? stdout : stderr);
            }

            return (true, stdout);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to run sprint dashboard script.");
            return (false, ex.Message);
        }
    }
}
