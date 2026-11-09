using System.Runtime.CompilerServices;
using System.Net.Http.Json;
using System.Text.Json;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.Options;

namespace GenerateDeliveryReports.Data.Services;

/// <summary>Minimal Microsoft.Extensions.AI adapter for a locally hosted Ollama instance.</summary>
public sealed class OllamaChatClient(HttpClient http, IOptions<AutomationReportingOptions> options) : IChatClient
{
    private readonly OllamaOptions _options = options.Value.Llm.Ollama;

    public async Task<ChatResponse> GetResponseAsync(IEnumerable<ChatMessage> messages, ChatOptions? options = null, CancellationToken cancellationToken = default)
    {
        var request = new
        {
            model = options?.ModelId ?? _options.Model,
            stream = false,
            messages = messages.Select(message => new { role = ToOllamaRole(message.Role), content = message.Text ?? string.Empty }),
            options = new { temperature = options?.Temperature ?? _options.Temperature, num_predict = options?.MaxOutputTokens ?? _options.MaxTokens }
        };
        using var response = await http.PostAsJsonAsync("api/chat", request, cancellationToken);
        response.EnsureSuccessStatusCode();
        var payload = await response.Content.ReadFromJsonAsync<OllamaResponse>(cancellationToken: cancellationToken)
            ?? throw new InvalidOperationException("Ollama returned an empty response.");
        return new ChatResponse(new ChatMessage(ChatRole.Assistant, payload.Message?.Content ?? string.Empty));
    }

    public async IAsyncEnumerable<ChatResponseUpdate> GetStreamingResponseAsync(IEnumerable<ChatMessage> messages, ChatOptions? options = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var response = await GetResponseAsync(messages, options, cancellationToken);
        yield return new ChatResponseUpdate(ChatRole.Assistant, response.Text ?? string.Empty);
    }

    public object? GetService(Type serviceType, object? serviceKey = null) => serviceType == typeof(IChatClient) ? this : null;
    public void Dispose() { }
    private static string ToOllamaRole(ChatRole role) => role == ChatRole.System ? "system" : role == ChatRole.Assistant ? "assistant" : "user";
    private sealed class OllamaResponse { public OllamaMessage? Message { get; set; } }
    private sealed class OllamaMessage { public string? Content { get; set; } }
}
