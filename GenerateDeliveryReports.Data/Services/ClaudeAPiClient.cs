
using System.Net.Http;
using System.Text;
using System.Text.Json;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Options;
namespace GenerateDeliveryReports.Data.Services
{
    public class ClaudeApiClient
    {
        private readonly HttpClient _http;
        private readonly AppSettings _settings;

        public ClaudeApiClient(HttpClient http, IOptions<AppSettings> options)
        {
            _http = http;
            _settings = options.Value;
        }
        public async Task<string> GenerateAsync(string systemPrompt, string userMessage)
        {
            

            var allTextBlocks = new List<string>();
            var currentMessages = new List<object> { new { role = "user", content = userMessage } };

            for (int i = 0; i < 5; i++)
            {
                // ... make the call (refactor the HTTP call above into a local method or repeat it)
                var responseJson = await GetContent(systemPrompt, currentMessages );

                var doc = JsonDocument.Parse(responseJson);
                var root = doc.RootElement;
                var stopReason = root.GetProperty("stop_reason").GetString();
                var contentArray = root.GetProperty("content");

                // Collect text blocks
                foreach (var block in contentArray.EnumerateArray())
                {
                    if (block.GetProperty("type").GetString() == "text")
                        allTextBlocks.Add(block.GetProperty("text").GetString() ?? "");
                }

                if (stopReason == "end_turn")
                    break;

                if (stopReason == "pause_turn")
                {
                    // Add assistant's response to message history and resend
                    currentMessages.Add(new { role = "assistant", content = contentArray.Clone() });
                    // rebuild requestBody with updated messages and resend
                }
            }

            return string.Concat(allTextBlocks);
        }

        private async Task<string> GetContent(string systemPrompt, List<object> messages)
        {
            var requestBody = new
            {
                model = _settings.BriefingModel,
                max_tokens = _settings.BriefingMaxTokens,
                system = systemPrompt,
                messages = messages.ToArray(),
                tools = new[]
                {
                    new { type = "web_search_20250305", name = "web_search", max_uses = _settings.BriefingMaxWebSearches }
                }
            };

            var json = JsonSerializer.Serialize(requestBody);
            var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
            request.Headers.Add("x-api-key", _settings.AnthropicApiKey);
            request.Headers.Add("anthropic-version", "2023-06-01");
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _http.SendAsync(request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
    }


    
}