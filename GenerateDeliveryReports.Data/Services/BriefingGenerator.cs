using System;
using System.IO;
using System.Linq;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;


namespace GenerateDeliveryReports.Data.Services
{
    public class BriefingGenerator
    {
        private IOptions<AppSettings> _options;
        private ILogger<BriefingGenerator> _Logger;
        private ClaudeApiClient _claudeApiClient;
        private BriefingCache _briefingCache;

        private const string SystemPrompt = """
                            You are preparing a daily AI and technology briefing for an Engagement Manager at a
                consulting firm with a strong Microsoft technology background.

                Output ONLY clean Markdown for the briefing itself — no preamble, no meta-commentary,
                no closing remarks. Start with: # Daily AI & Tech Briefing — {date}

                Summarize in bullet points the most important recent headlines across these sections:

                - Generative AI & LLMs (model releases, major research, company moves — closed/proprietary
                models like GPT, Claude, Gemini belong here, NOT in the LLM Spotlight section below)
                - Enterprise Tech & SaaS (tools, platforms relevant to consulting)
                - AI in Consulting & Strategy (how AI is reshaping professional services)
                - Emerging Technology (robotics, quantum, biotech, frontier tech)
                - Microsoft & .NET (new .NET releases, C# updates, Azure, Visual Studio, Blazor, MAUI,
                ASP.NET, Microsoft Copilot integrations — highlight anything relevant for teams
                building on the Microsoft stack or advising enterprise clients)

                Then add these three sections:

                ## 🔬 LLM Spotlight (1 open-source model per day)
                Pick one notable OPEN-SOURCE / open-weight LLM — e.g. from families like Llama (Meta),
                Qwen (Alibaba), DeepSeek, Mistral, Gemma (Google), Phi (Microsoft), Kimi (Moonshot),
                GLM (Zhipu), OLMo (AI2), Falcon (TII), or similar. Prefer one released or updated
                recently, and avoid repeating any model/provider listed in the "avoid repeating" list
                given in the user message. Cover:
                - Who built it and license terms (Apache 2.0, MIT, Llama community license, custom, etc.)
                - What makes it distinctive (architecture, parameter count, MoE vs dense, training approach)
                - Key capabilities (context window, modalities, languages, tool use, reasoning)
                - Benchmark performance vs other open models and vs closed frontier models
                - How to access/run it (Hugging Face, Ollama, API providers hosting it, self-host hardware needs)
                - Best-fit use cases, including enterprise/consulting relevance (on-prem/data-residency,
                cost vs closed APIs)

                ## ✅ AI Win of the Day
                One concrete, well-documented real-world case where AI delivered strong, measurable
                results — consulting, enterprise ops, healthcare, finance, legal, or engineering.
                Name the organization, what they did, and the quantified outcome.

                ## ⚠️ AI Struggle of the Day
                One honest, well-documented case where AI failed to meet expectations, caused problems,
                or revealed meaningful limitations (product failure, hallucination incident, bias issue,
                regulatory setback, underdelivered deployment). Be factual and balanced.

                Keep everything sharp and business-relevant.
            """;

        public BriefingGenerator(IOptions<AppSettings> options, ILogger<BriefingGenerator> logger, ClaudeApiClient claudeApiClient,
        BriefingCache briefingCache)
        {
            _options = options;
            _Logger = logger;
            _claudeApiClient = claudeApiClient;
            _briefingCache = briefingCache;
        }

        public async Task GenerateAsync()
        {
            try
            {
                var archiveFolder = _options.Value.ResolvedBriefingArchiveFolder;

                var recentFiles = Directory.GetFiles(archiveFolder, "*.md")
                    .OrderByDescending(f => f)
                    .Take(30)
                    .Select(Path.GetFileNameWithoutExtension)
                    .ToList();


                var today = DateTime.Today.ToString("yyyy-MM-dd");
                var avoidList = recentFiles.Count > 0
                    ? string.Join(", ", recentFiles)
                    : "none";
                var userMessage = $"Generate today's briefing for {today}. Avoid repeating these recent LLM Spotlight picks: {avoidList}.";

                var markdown = await _claudeApiClient.GenerateAsync(SystemPrompt, userMessage);
                var html = Markdig.Markdown.ToHtml(markdown);

                var archivePath = Path.Combine(archiveFolder, $"daily-briefing-{today}.md");

                await File.WriteAllTextAsync(archivePath, markdown);

                _briefingCache.Set(new BriefingCache.BriefingSnapshot(DateTime.Now, html, markdown));
            }
            catch (Exception ex)
            {
                _Logger.LogError(ex, "An error occurred while generating the briefing.");
                throw;
            }
        }
    }

}
