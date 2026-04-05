using Microsoft.Extensions.Logging;

namespace GenerateDeliveryReports.Data.Extensions;

public static class FileExtensions
{
    public static FileInfo? GetRecentlyModifiedSimilarFile(this string filePath)
    {
        var directoryPath = Path.GetDirectoryName(filePath);
        if (!Directory.Exists(directoryPath))
            return null;

        var baseFileName = Path.GetFileName(filePath);
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(baseFileName);
        string fileExtension = Path.GetExtension(baseFileName);
        string searchPattern = $"{fileNameWithoutExtension}*{fileExtension}";

        var files = Directory.GetFiles(directoryPath, searchPattern);

        var filteredFiles = files.Where(file =>
        {
            string currentFileName = Path.GetFileNameWithoutExtension(file);
            return currentFileName.StartsWith(fileNameWithoutExtension) &&
                   (currentFileName.Equals(fileNameWithoutExtension, StringComparison.OrdinalIgnoreCase) ||
                    System.Text.RegularExpressions.Regex.IsMatch(currentFileName,
                        $"{System.Text.RegularExpressions.Regex.Escape(fileNameWithoutExtension)}\\s*\\(\\d+\\)"));
        });

        return filteredFiles
            .Select(file => new FileInfo(file))
            .OrderByDescending(f => f.LastWriteTime)
            .FirstOrDefault();
    }
}
