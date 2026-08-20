
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace GenerateDeliveryReports.Worker;
public class RunHistoryService
{
    private readonly ILogger<RunHistoryService> _logger;
    private const int MaxRunsToKeep = 30;

public RunHistoryService(ILogger<RunHistoryService> logger)
{
    _logger = logger;
}
    public async Task<List<RunHistory>> AddRunAndGetHistoryAsync(
    List<SprintReportResult> results, 
    string historyFilePath)
{
    try{

        var directory = Path.GetDirectoryName(historyFilePath);
        if (!string.IsNullOrEmpty(directory))
        {
            _logger.LogInformation("Creating directory for run history: {Directory}", directory);
            Directory.CreateDirectory(directory);
        }

    var historyList = File.Exists(historyFilePath)
        ? JsonConvert.DeserializeObject<List<RunHistory>>(
            await File.ReadAllTextAsync(historyFilePath)) ?? new List<RunHistory>()
        : new List<RunHistory>();

    var obj = new RunHistory
    {
        RunTime = DateTime.Now,
        Results = results
    };
                    
    
    // 1. Read existing history from JSON (or create empty list if file doesn't exist)
    // 2. Create new RunHistory object with current DateTime.Now and the results
    // 3. Add it to the list
    // 4. Sort by RunTime descending
    // 5. Keep only first 30 items
    // 6. Serialize and save back to JSON file
    historyList.Add(obj);
    historyList = historyList
        .OrderByDescending(history => history.RunTime)
        .Take(MaxRunsToKeep)
        .ToList();

    _logger.LogInformation("Saving run history to {HistoryFilePath}", historyFilePath);
    await File.WriteAllTextAsync(
        historyFilePath,
        JsonConvert.SerializeObject(historyList, Formatting.Indented));

    return historyList;
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "Error while adding run history.");
        throw;
    }
}

}