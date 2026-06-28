using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Logging;
using System.Text;
using System.Text.Json;
using Whisper.net;
using Whisper.net.Ggml;

namespace GenerateDeliveryReports.Data.Services;

public class MeetingMinutesService
{
    private readonly string _modelPath;
    private readonly ILogger<MeetingMinutesService> _logger;
    private readonly string _ollamaUrl;

    public MeetingMinutesService(IWebHostEnvironment env, ILogger<MeetingMinutesService> logger)
    {
        _logger = logger;
        _ollamaUrl = "http://localhost:11434";
        var modelDir = Path.Combine(env.ContentRootPath, "WhisperModels");
        Directory.CreateDirectory(modelDir);
        _modelPath = Path.Combine(modelDir, "ggml-base.bin");
    }

    public async Task EnsureModelDownloadedAsync()
    {
        if (File.Exists(_modelPath)) return;

        _logger.LogInformation("Downloading Whisper base model to {Path}…", _modelPath);
        using var modelStream = await WhisperGgmlDownloader.Default.GetGgmlModelAsync(GgmlType.Base);
        using var fileStream = File.Create(_modelPath);
        await modelStream.CopyToAsync(fileStream);
        _logger.LogInformation("Whisper model download complete.");
    }

    public async Task<string> TranscribeAsync(byte[] wavBytes)
    {
        await EnsureModelDownloadedAsync();

        using var factory = WhisperFactory.FromPath(_modelPath);
        using var processor = factory.CreateBuilder()
            .WithLanguage("auto")
            .Build();

        var sb = new StringBuilder();
        using var stream = new MemoryStream(wavBytes);
        await foreach (var segment in processor.ProcessAsync(stream))
            sb.Append(segment.Text).Append(' ');

        return sb.ToString().Trim();
    }

    public async Task<string> GenerateMinutesAsync(string transcript, string model = "llama3")
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };

        var prompt = $"""
            You are a professional meeting minutes writer.
            Transform the following meeting transcript into structured meeting minutes.

            Format your output with these sections (omit any that aren't applicable):
            ## Summary
            ## Attendees
            ## Key Discussion Points
            ## Decisions Made
            ## Action Items
            ## Next Steps

            Transcript:
            {transcript}
            """;

        var body = JsonSerializer.Serialize(new { model, prompt, stream = false });

        HttpResponseMessage response;
        try
        {
            response = await http.PostAsync(
                $"{_ollamaUrl}/api/generate",
                new StringContent(body, Encoding.UTF8, "application/json"));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to reach Ollama at {Url}", _ollamaUrl);
            throw new InvalidOperationException($"Could not connect to Ollama at {_ollamaUrl}. Make sure Ollama is running.", ex);
        }

        var json = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);

        if (doc.RootElement.TryGetProperty("error", out var err))
            throw new InvalidOperationException($"Ollama error: {err.GetString()}");

        return doc.RootElement.GetProperty("response").GetString() ?? string.Empty;
    }

    public bool IsModelDownloaded() => File.Exists(_modelPath);
}
