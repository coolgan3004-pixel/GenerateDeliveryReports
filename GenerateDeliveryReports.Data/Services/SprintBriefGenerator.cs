using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace GenerateDeliveryReports.Data.Services;

/// <summary>
/// Native C# port of GenerateSprintDashboard/daily_brief.py -- reads every "&lt;App&gt; - Sprint
/// Metrics.xlsx" workbook under a root folder, derives each team's current-sprint progress from
/// its Status sheet, and renders one self-contained HTML brief. Replaces the former Python
/// subprocess call (SprintDashboardService used to shell out via Process.Start) so the feature
/// works on Azure App Service, which has no Python runtime.
/// </summary>
public static class SprintBriefGenerator
{
    private sealed class StoryRecord
    {
        public string Dev = "";
        public string Id = "";
        public string Desc = "";
        public string StatusRaw = "";
        public string Stage = "";
        public double Weight;
        public double Est;
        public DateTime? DevEnd;
        public bool Stretch;
        public string Flag = "";
        public DateTime? Eta;
        public int Pct;
    }

    private sealed class RollUp
    {
        public int ProgPct;
        public int DonePct;
        public int TimePct;
        public int ProjPct;
        public int DaysLeft;
        public int Blocked;
        public int Off;
        public int Risk;
        public int N;
        public int NDone;
        public int StretchN;
        public int StretchDone;
        public string Verdict = "";
        public bool Overdue;
    }

    private sealed class TeamRecord
    {
        public string Team = "";
        public string Path = "";
        public string App = "";
        public string Sprint = "";
        public DateTime? Start;
        public DateTime? End;
        public List<StoryRecord> Stories = new();
        public string Note = "";
        public bool Ok;
        public RollUp? Roll;
    }

    private static readonly Dictionary<string, string[]> Col = new()
    {
        ["dev"] = new[] { "developer", "owner", "assignee" },
        ["id"] = new[] { "story id", "story id/defect", "id/defect", "defect id", "id" },
        ["desc"] = new[] { "description", "summary", "title" },
        ["status"] = new[] { "status" },
        ["est"] = new[] { "estimated(sp)", "estimated (sp)", "estimated(hrs/sp)", "estimated", "story points", "estimate" },
        ["act"] = new[] { "actual(sp)", "actual (sp)", "actual(hrs/sp)", "actual" },
        ["dstart"] = new[] { "dev start date", "dev start" },
        ["dend"] = new[] { "dev end date", "dev end" },
        ["sprintcol"] = new[] { "sprint" },
        ["main"] = new[] { "main sprint", "main/stretch" },
    };

    private static readonly string[] StageHints =
    {
        "story ", "defect ", "completed", "complete", "review", "tested",
        "testing", "develop", "refine", "approved", "ready for", "blocked",
        "fixed", "peer", "in progress", "qa", "crc", "not started", "to do",
    };

    private static readonly Dictionary<string, string> Clr = new()
    {
        ["On Track"] = "#1a7f37", ["At Risk"] = "#bf8700", ["Off Track"] = "#cf222e",
        ["Blocked"] = "#8250df", ["Done"] = "#57606a", ["No Data"] = "#8c959f",
    };

    private static readonly Dictionary<string, string> Bg = new()
    {
        ["On Track"] = "#e6f4ea", ["At Risk"] = "#fff4d6", ["Off Track"] = "#ffe3e3",
        ["Blocked"] = "#f1e9fc", ["Done"] = "#eef1f4", ["No Data"] = "#eef1f4",
    };

    private static readonly Dictionary<string, int> Ord = new()
    {
        ["Off Track"] = 0, ["At Risk"] = 1, ["Blocked"] = 1, ["On Track"] = 2, ["Done"] = 3, ["No Data"] = 4,
    };

    // Every selector scoped under .dsb-root so the brief can be dropped into any page.
    private const string Style = @"
.dsb-root,.dsb-root *{box-sizing:border-box}
.dsb-root{font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1f2328;
 background:#f6f8fa;padding:0;margin:0}
.dsb-root .wrap{max-width:1180px;margin:0 auto;padding:24px 20px 60px}
.dsb-root h1{font-size:24px;margin:0 0 2px}
.dsb-root .sub{color:#656d76;font-size:13px}
.dsb-root .kpis{display:flex;flex-wrap:wrap;gap:12px;margin:18px 0 26px}
.dsb-root .kpi{background:#fff;border:1px solid #d0d7de;border-radius:10px;padding:11px 16px;min-width:118px}
.dsb-root .kpi b{font-size:22px;display:block;line-height:1.1}
.dsb-root .kpi span{font-size:11.5px;color:#656d76}
.dsb-root .cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:13px;margin-bottom:30px}
.dsb-root .card{background:#fff;border:1px solid #d0d7de;border-left-width:5px;border-radius:10px;padding:13px 15px}
.dsb-root .card h3{margin:0;font-size:15.5px}
.dsb-root .meta{color:#656d76;font-size:11.5px;margin:3px 0 9px}
.dsb-root .m{font-size:11.5px;color:#424a53;margin:3px 0}
.dsb-root h2.sec{font-size:15px;border-bottom:2px solid #d0d7de;padding-bottom:6px;margin:28px 0 14px}
.dsb-root .team{background:#fff;border:1px solid #d0d7de;border-radius:10px;overflow:hidden;margin-bottom:16px}
.dsb-root .th{padding:11px 15px;display:flex;justify-content:space-between;align-items:center;gap:10px;border-bottom:1px solid #eaecef;cursor:pointer}
.dsb-root .th h3{margin:0;font-size:16px}
.dsb-root .th .meta{margin:2px 0 0}
.dsb-root table{border-collapse:collapse;width:100%;font-size:12.5px}
.dsb-root th,.dsb-root td{text-align:left;padding:7px 11px;border-bottom:1px solid #eef0f2;vertical-align:middle}
.dsb-root th{background:#f6f8fa;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;color:#656d76}
.dsb-root .late{color:#cf222e;font-weight:600}
.dsb-root .note{font-size:11.5px;color:#9a6700;background:#fff8e6;padding:4px 10px;border-top:1px solid #f0e3b8}
.dsb-root details>summary{list-style:none}
.dsb-root details>summary::-webkit-details-marker{display:none}
.dsb-root .inact td{color:#656d76}
.dsb-root{scroll-behavior:smooth}
.dsb-root a.cardlink{text-decoration:none;color:inherit;display:block}
.dsb-root .card{transition:box-shadow .12s,transform .12s}
.dsb-root a.cardlink:hover .card{box-shadow:0 3px 12px rgba(0,0,0,.13);transform:translateY(-1px)}
.dsb-root .missing{background:#fff0f0;border:1px solid #ffc9c9;border-left:4px solid #cf222e;color:#86181d;border-radius:8px;padding:9px 13px;margin:14px 0 2px;font-size:12.5px;line-height:1.6}
.dsb-root .m.stretch{color:#8250df}
.dsb-root .stag{background:#f1e9fc;color:#8250df;font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:8px;text-transform:uppercase;letter-spacing:.3px;margin-left:5px}
.dsb-root tr.strow td{background:#faf8ff}
.dsb-root .team:target{box-shadow:0 0 0 2px #0969da;border-color:#0969da}
";

    public static async Task<(bool Success, string Message)> GenerateAsync(string root, string outPath)
    {
        return await Task.Run(() =>
        {
            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("GenerateDeliveryReports");

                var today = DateTime.Today;
                var files = Discover(root);
                var recs = new List<TeamRecord>();
                foreach (var f in files)
                {
                    try
                    {
                        var r = ReadWorkbook(f, today);
                        if (r.Ok)
                            Classify(r, today);
                        recs.Add(r);
                    }
                    catch (Exception ex)
                    {
                        recs.Add(new TeamRecord { Team = TeamName(f), Path = f, Note = $"Read error: {ex.Message}" });
                    }
                }

                var html = Render(recs, today);
                var dir = System.IO.Path.GetDirectoryName(outPath);
                if (!string.IsNullOrEmpty(dir))
                    Directory.CreateDirectory(dir);
                File.WriteAllText(outPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                var (active, inactive) = SplitActive(recs);
                var onTrack = active.Count(r => r.Roll!.Verdict == "On Track");
                var atRisk = active.Count(r => r.Roll!.Verdict == "At Risk");
                var offTrack = active.Count(r => r.Roll!.Verdict == "Off Track");
                var blocked = active.Sum(r => r.Roll!.Blocked);
                var summary = $"Daily Sprint Brief — {today:yyyy-MM-dd}\n" +
                    $"{active.Count} active team(s), {inactive.Count} inactive: {onTrack} on track, " +
                    $"{atRisk} at risk, {offTrack} off track, {blocked} blocked story-owner(s).";
                return (true, summary);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        });
    }

    // ---------------------------------------------------------------- discovery

    private static List<string> Discover(string root)
    {
        var files = new List<string>();
        if (!Directory.Exists(root))
            return files;
        foreach (var p in Directory.EnumerateFiles(root, "*.xlsx", SearchOption.AllDirectories))
        {
            var b = System.IO.Path.GetFileName(p);
            if (b.StartsWith("~$")) continue;
            if (p.ToLowerInvariant().Contains("archive")) continue;
            if (b.ToLowerInvariant().EndsWith("sprint metrics.xlsx")) files.Add(p);
        }
        files.Sort(StringComparer.OrdinalIgnoreCase);
        return files;
    }

    // ---------------------------------------------------------------- helpers

    private static string Norm(object? s)
        => s == null ? "" : Regex.Replace(s.ToString()!.Trim().ToLowerInvariant(), @"\s+", " ");

    private static bool Naf(object? v)
    {
        if (v == null) return false;
        var s = v.ToString()!.Trim();
        return s != "" && !string.Equals(s, "#N/A", StringComparison.OrdinalIgnoreCase)
            && s != "None" && s != "0";
    }

    private static DateTime? ToDate(object? v)
    {
        if (v == null) return null;
        if (v is DateTime dt) return dt.Date;
        if (v is double d)
        {
            try { return DateTime.FromOADate(d).Date; } catch { return null; }
        }
        var s = v.ToString()!.Trim();
        string[] fmts = { "yyyy-MM-dd", "dd-MM-yyyy", "MM/dd/yyyy", "dd/MM/yyyy", "dd-MMM-yyyy", "yyyy-MM-dd HH:mm:ss" };
        foreach (var fmt in fmts)
        {
            if (DateTime.TryParseExact(s, fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
                return parsed.Date;
        }
        return null;
    }

    private static double? Num(object? v)
    {
        if (v == null) return null;
        if (v is double d) return double.IsNaN(d) ? null : d;
        if (double.TryParse(v.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var f))
            return double.IsNaN(f) ? null : f;
        return null;
    }

    private static int WorkingDays(DateTime? d0, DateTime? d1)
    {
        if (d0 == null || d1 == null || d1 < d0) return 0;
        var n = 0;
        var cur = d0.Value;
        while (cur <= d1.Value)
        {
            if (cur.DayOfWeek != DayOfWeek.Saturday && cur.DayOfWeek != DayOfWeek.Sunday) n++;
            cur = cur.AddDays(1);
        }
        return n;
    }

    private static DateTime AddWorkingDays(DateTime start, double k)
    {
        var cur = start;
        var added = 0;
        var target = Math.Max(0, (int)Math.Round(k, MidpointRounding.ToEven));
        while (added < target)
        {
            cur = cur.AddDays(1);
            if (cur.DayOfWeek != DayOfWeek.Saturday && cur.DayOfWeek != DayOfWeek.Sunday) added++;
        }
        return cur;
    }

    private static string TeamName(string path)
    {
        var b = System.IO.Path.GetFileName(path);
        var name = Regex.Replace(b, @"\s*-\s*sprint metrics.*$", "", RegexOptions.IgnoreCase).Trim();
        return string.IsNullOrEmpty(name) ? System.IO.Path.GetFileNameWithoutExtension(b) : name;
    }

    private static int? FindCol(IList<object?> header, string[] keys)
    {
        var cells = header.Select(Norm).ToList();
        foreach (var k in keys)
            for (var i = 0; i < cells.Count; i++)
                if (cells[i] == k) return i;
        foreach (var k in keys)
            for (var i = 0; i < cells.Count; i++)
                if (cells[i].Contains(k) || (k.Contains(cells[i]) && cells[i].Length > 2))
                    return i;
        return null;
    }

    private static (string Stage, double Weight) StageOf(string? status)
    {
        var t = Regex.Replace((status ?? "").Trim().ToLowerInvariant(), @"\s+", " ");
        if (string.IsNullOrEmpty(t)) return ("Not Started", 0.0);
        if (new[] { "blocked", "on hold", "hold", "impediment", "blocker" }.Any(t.Contains))
            return ("Blocked", 0.40);
        if (new[] { "approved", "completed", "complete", "fixed", "closed", "done", "signed off", "sign off", "deployed", "released" }.Any(t.Contains))
            return ("Done", 1.0);
        if (t.Contains("ready for test") || t.Contains("ready to test")) return ("Ready for Test", 0.80);
        if (t.Contains("test") || t.Contains("qa")) return ("Testing", 0.85);
        if (t.Contains("review") || t.Contains("crc")) return ("In Review", 0.75);
        if (new[] { "develop", "in progress", "inprogress", "being fixed", "wip" }.Any(t.Contains))
            return ("In Development", 0.45);
        if (new[] { "refine", "groom", "analysis", "backlog", "not started", "to do", "todo", "ready for dev", "yet to" }.Any(t.Contains))
            return ("Refinement", 0.10);
        return ("In Progress", 0.30);
    }

    // ---------------------------------------------------------------- dashboard row1

    private static Dictionary<string, object?> ParseDashboardHeader(ExcelWorksheet ws)
    {
        var result = new Dictionary<string, object?>();
        var endCol = ws.Dimension?.End.Column ?? 0;
        for (var j = 0; j <= endCol - 2; j++)
        {
            var k = Norm(ws.Cells[1, j + 1].Value);
            if (string.IsNullOrEmpty(k)) continue;
            var v = ws.Cells[1, j + 2].Value;
            if ((k == "current sprint sheet" || k == "sprint sheet") && !result.ContainsKey("sheet"))
                result["sheet"] = v;
            else if (k == "current sprint" && !result.ContainsKey("sprint"))
                result["sprint"] = v;
            else if (k.Contains("sprint start") && !result.ContainsKey("start"))
                result["start"] = v;
            else if (k.Contains("sprint end") && !result.ContainsKey("end"))
                result["end"] = v;
            else if (k.Contains("application name") && !result.ContainsKey("app"))
                result["app"] = v;
        }
        return result;
    }

    private static object? Get(Dictionary<string, object?> d, string key) => d.TryGetValue(key, out var v) ? v : null;

    private static string? LocateStatusSheet(ExcelWorkbook wb, object? sheetHint, object? sprintName)
    {
        var names = wb.Worksheets.Select(w => w.Name).ToList();
        if (Naf(sheetHint))
        {
            var h = sheetHint!.ToString()!.Trim();
            foreach (var sn in names) if (sn.Trim() == h) return sn;
            foreach (var sn in names) if (string.Equals(sn.Trim(), h, StringComparison.OrdinalIgnoreCase)) return sn;
            var prefix = (h.Length >= 14 ? h.Substring(0, 14) : h).ToLowerInvariant();
            foreach (var sn in names) if (sn.Trim().ToLowerInvariant().StartsWith(prefix)) return sn;
        }
        if (Naf(sprintName))
        {
            var sp = sprintName!.ToString()!.Trim().ToLowerInvariant();
            var spPrefix = sp.Length >= 10 ? sp.Substring(0, 10) : sp;
            var cand = names.FirstOrDefault(sn => sn.ToLowerInvariant().Contains("status") && sn.Trim().ToLowerInvariant().StartsWith(spPrefix));
            if (cand != null) return cand;
        }
        return null;
    }

    private static int? DetectStatusCol(List<object?[]> rows, HashSet<int> exclude)
    {
        var ncol = rows.Count == 0 ? 0 : rows.Max(r => r.Length);
        int? best = null;
        var bestScore = 0;
        var upper = Math.Min(rows.Count, 25);
        for (var j = 0; j < ncol; j++)
        {
            if (exclude.Contains(j)) continue;
            var score = 0;
            for (var ri = 1; ri < upper; ri++)
            {
                var r = rows[ri];
                var v = j < r.Length ? Norm(r[j]) : "";
                if (v.Length > 0 && StageHints.Any(v.Contains)) score++;
            }
            if (score > bestScore) { best = j; bestScore = score; }
        }
        return bestScore >= 2 ? best : null;
    }

    // ---------------------------------------------------------------- read one workbook

    private static List<object?[]> ReadAllRows(ExcelWorksheet ws)
    {
        var result = new List<object?[]>();
        if (ws.Dimension == null) return result;
        var endRow = ws.Dimension.End.Row;
        var endCol = ws.Dimension.End.Column;
        for (var r = 1; r <= endRow; r++)
        {
            var row = new object?[endCol];
            for (var c = 1; c <= endCol; c++)
                row[c - 1] = ws.Cells[r, c].Value;
            result.Add(row);
        }
        return result;
    }

    private static TeamRecord ReadWorkbook(string path, DateTime today)
    {
        var rec = new TeamRecord { Team = TeamName(path), Path = path };
        using var package = new ExcelPackage(new FileInfo(path));
        var wb = package.Workbook;
        var dashSheet = wb.Worksheets.FirstOrDefault(w => w.Name == "Dashboard");
        if (dashSheet == null)
        {
            rec.Note = "No Dashboard sheet.";
            return rec;
        }

        var hdr = ParseDashboardHeader(dashSheet);
        rec.App = (Get(hdr, "app")?.ToString() ?? "").Trim();
        rec.Sprint = (Get(hdr, "sprint")?.ToString() ?? "").Trim();
        rec.Start = ToDate(Get(hdr, "start"));
        rec.End = ToDate(Get(hdr, "end"));
        if (!Naf(Get(hdr, "sprint")))
        {
            rec.Note = "No current sprint set in Dashboard (#N/A).";
            return rec;
        }

        var sheetName = LocateStatusSheet(wb, Get(hdr, "sheet"), Get(hdr, "sprint"));
        if (sheetName == null)
        {
            rec.Note = $"Current-sprint status sheet not found ({Get(hdr, "sheet")}).";
            return rec;
        }

        var ws = wb.Worksheets[sheetName];
        var rows = ReadAllRows(ws);
        if (rows.Count == 0)
        {
            rec.Note = "Status sheet empty.";
            return rec;
        }

        var header = rows[0].ToList();
        var ci = Col.ToDictionary(kv => kv.Key, kv => FindCol(header, kv.Value));
        if (ci["status"] == null)
        {
            var ex = new HashSet<int>();
            if (ci["desc"].HasValue) ex.Add(ci["desc"]!.Value);
            if (ci["id"].HasValue) ex.Add(ci["id"]!.Value);
            ci["status"] = DetectStatusCol(rows, ex);
            if (ci["status"] == null && ci["desc"].HasValue)
                ci["status"] = ci["desc"]!.Value + 1;
        }

        var stories = new List<StoryRecord>();
        for (var ri = 1; ri < rows.Count; ri++)
        {
            var r = rows[ri];
            if (!r.Any(Naf)) continue;

            object? G(string key) => ci.TryGetValue(key, out var idx) && idx.HasValue && idx.Value < r.Length ? r[idx.Value] : null;

            var sid = G("id");
            var desc = G("desc");
            var status = G("status");
            if (!Naf(sid) && !Naf(desc)) continue;
            if (sid is string sidStr)
            {
                var low = sidStr.Trim().ToLowerInvariant();
                if (low is "s.no" or "total" or "totals") continue;
            }

            var (stage, weight) = StageOf(status?.ToString());
            var est = Num(G("est"));
            var isStretch = Norm(G("main")) == "stretch" || Norm(G("sprintcol")).Contains("stretch");
            stories.Add(new StoryRecord
            {
                Dev = (G("dev")?.ToString() ?? "").Trim(),
                Id = (sid?.ToString() ?? "").Trim(),
                Desc = (desc?.ToString() ?? "").Trim(),
                StatusRaw = (status?.ToString() ?? "").Trim(),
                Stage = stage,
                Weight = weight,
                Est = est ?? 0.0,
                DevEnd = ToDate(G("dend")),
                Stretch = isStretch,
            });
        }

        rec.Stories = stories;
        rec.Ok = true;
        if (rec.End.HasValue && rec.End.Value < today)
        {
            var days = (today - rec.End.Value).Days;
            rec.Note = $"Sprint ended {days} day(s) ago — data may be stale or in close-out.";
        }
        else if (rec.Start.HasValue && rec.Start.Value > today)
        {
            rec.Note = "Sprint has not started yet.";
        }
        return rec;
    }

    // ---------------------------------------------------------------- analytics

    private static void Classify(TeamRecord rec, DateTime today)
    {
        var allStories = rec.Stories;
        var stories = allStories.Where(s => !s.Stretch).ToList();
        var stretch = allStories.Where(s => s.Stretch).ToList();
        var usePts = stories.Any(s => s.Est != 0);
        double W(StoryRecord s) => usePts ? s.Est : 1.0;
        var total = stories.Sum(s => W(s));
        var done = stories.Where(s => s.Stage == "Done").Sum(s => W(s));
        var weighted = stories.Sum(s => W(s) * s.Weight);
        var prog = total != 0 ? weighted / total : 0.0;
        var doneFrac = total != 0 ? done / total : 0.0;

        var start = rec.Start;
        var end = rec.End;
        var totalWd = WorkingDays(start, end);
        if (totalWd == 0) totalWd = 1;
        var elapsedWd = start.HasValue ? Math.Min(WorkingDays(start, today), totalWd) : totalWd;
        var timeFrac = totalWd != 0 ? (double)elapsedWd / totalWd : 1.0;
        var daysLeft = end.HasValue ? WorkingDays(today, end) : 0;

        double proj;
        if (timeFrac > 0) proj = Math.Min(2.0, prog / timeFrac);
        else proj = prog >= 1 ? 1.0 : 0.0;
        var projPct = (int)Math.Round(proj * 100, MidpointRounding.ToEven);

        var blocked = stories.Count(s => s.Stage == "Blocked");

        foreach (var s in allStories)
        {
            var (flag, eta, pct) = StoryFlag(s, timeFrac, end, daysLeft, today);
            s.Flag = flag;
            s.Eta = eta;
            s.Pct = pct;
        }
        var off = stories.Count(s => s.Flag == "Off Track");
        var risk = stories.Count(s => s.Flag == "At Risk");

        var overdue = end.HasValue && end.Value < today && doneFrac < 0.999;
        string verdict;
        if (total == 0) verdict = "No Data";
        else if (doneFrac >= 0.999) verdict = "On Track";
        else if (overdue) verdict = "Off Track";
        else if (blocked > 0 || off >= 2 || projPct < 85)
            verdict = (off >= 2 || projPct < 70) ? "Off Track" : "At Risk";
        else if (risk > 0 || off > 0 || projPct < 100) verdict = "At Risk";
        else verdict = "On Track";

        rec.Roll = new RollUp
        {
            ProgPct = (int)Math.Round(prog * 100, MidpointRounding.ToEven),
            DonePct = (int)Math.Round(doneFrac * 100, MidpointRounding.ToEven),
            TimePct = (int)Math.Round(timeFrac * 100, MidpointRounding.ToEven),
            ProjPct = projPct,
            DaysLeft = daysLeft,
            Blocked = blocked,
            Off = off,
            Risk = risk,
            N = stories.Count,
            NDone = stories.Count(s => s.Stage == "Done"),
            StretchN = stretch.Count,
            StretchDone = stretch.Count(s => s.Stage == "Done"),
            Verdict = verdict,
            Overdue = overdue,
        };
    }

    private static (string Flag, DateTime? Eta, int Pct) StoryFlag(StoryRecord s, double timeFrac, DateTime? end, int daysLeft, DateTime today)
    {
        if (s.Stage == "Done") return ("Done", end, 100);
        if (s.Stage == "Blocked") return ("Blocked", null, (int)Math.Round(s.Weight * 100, MidpointRounding.ToEven));

        var weight = s.Weight;
        DateTime? eta = (s.DevEnd.HasValue && s.DevEnd.Value >= today) ? s.DevEnd : null;
        if (eta == null && end.HasValue)
        {
            var remaining = Math.Max(0.0, 1 - weight);
            eta = AddWorkingDays(today, Math.Max(1, remaining * Math.Max(daysLeft, 1)));
        }
        var lateByDate = s.DevEnd.HasValue && end.HasValue && s.DevEnd.Value > end.Value;
        var expected = timeFrac;
        string flag;
        if (weight >= expected - 0.10 && !lateByDate && (!end.HasValue || !eta.HasValue || eta.Value <= end.Value))
            flag = "On Track";
        else if (weight >= expected - 0.25 && !lateByDate)
            flag = "At Risk";
        else
            flag = "Off Track";
        if (weight <= 0.10 && timeFrac > 0.55) flag = "Off Track";
        if (lateByDate && flag == "On Track") flag = "At Risk";
        return (flag, eta, (int)Math.Round(weight * 100, MidpointRounding.ToEven));
    }

    // ---------------------------------------------------------------- HTML

    private static string Chip(string flag) =>
        $"<span style=\"background:{Bg[flag]};color:{Clr[flag]};padding:2px 9px;border-radius:11px;font-size:11.5px;font-weight:600;white-space:nowrap\">{flag}</span>";

    private static string Bar(int pct, string color, int w = 120)
    {
        pct = Math.Max(0, Math.Min(100, pct));
        return $"<span style=\"background:#e7ebef;border-radius:6px;height:8px;width:{w}px;display:inline-block;vertical-align:middle\">" +
               $"<span style=\"background:{color};height:8px;border-radius:6px;width:{pct}%;display:inline-block\"></span></span>";
    }

    private static string Esc(object? x) => WebUtility.HtmlEncode(x?.ToString() ?? "");

    private static string Slug(string t)
    {
        var s = Regex.Replace(t.ToLowerInvariant(), "[^a-z0-9]+", "-").Trim('-');
        return string.IsNullOrEmpty(s) ? "team" : s;
    }

    private static (List<TeamRecord> Active, List<TeamRecord> Inactive) SplitActive(List<TeamRecord> recs)
    {
        var inactive = recs.Where(r => !(r.Ok && r.Roll is { N: > 0 })).ToList();
        // OrderBy/ThenBy (stable), not List.Sort (unstable) -- ties on (end date, verdict) must
        // keep discovery order to match the Python script's stable sort() byte-for-byte.
        var active = recs.Where(r => r.Ok && r.Roll is { N: > 0 })
            .OrderBy(r => r.End ?? DateTime.MaxValue)
            .ThenBy(r => Ord[r.Roll!.Verdict])
            .ToList();
        return (active, inactive);
    }

    private static string BuildInner(List<TeamRecord> recs, DateTime today)
    {
        var (active, inactive) = SplitActive(recs);
        var tally = new Dictionary<string, int> { ["On Track"] = 0, ["At Risk"] = 0, ["Off Track"] = 0 };
        foreach (var r in active)
            if (tally.ContainsKey(r.Roll!.Verdict)) tally[r.Roll!.Verdict]++;
        var totStories = active.Sum(r => r.Roll!.N);
        var totDone = active.Sum(r => r.Roll!.NDone);
        var totBlocked = active.Sum(r => r.Roll!.Blocked);

        var p = new StringBuilder();
        p.Append("<div class=\"wrap\">");
        p.Append("<h1>Daily Sprint Status Brief</h1>");
        p.Append($"<div class=\"sub\">{today.ToString("dddd, dd MMMM yyyy", CultureInfo.InvariantCulture)} &middot; {active.Count} active team(s), {inactive.Count} inactive</div>");

        if (inactive.Count > 0)
        {
            var items = inactive.Select(r =>
            {
                var reason = string.IsNullOrEmpty(r.Note) ? "no current sprint set" : r.Note;
                return $"<b>{Esc(r.Team)}</b> &ndash; {Esc(reason)}";
            });
            p.Append($"<div class=\"missing\"><b>&#9888; {inactive.Count} team(s) missing a current sprint:</b> {string.Join("; ", items)}</div>");
        }

        p.Append("<div class=\"kpis\">");
        var kpis = new (string Label, string Val, string Col)[]
        {
            ("On Track", tally["On Track"].ToString(), Clr["On Track"]),
            ("At Risk", tally["At Risk"].ToString(), Clr["At Risk"]),
            ("Off Track", tally["Off Track"].ToString(), Clr["Off Track"]),
            ("Stories done", $"{totDone}/{totStories}", "#1f2328"),
            ("Blocked", totBlocked.ToString(), Clr["Blocked"]),
        };
        foreach (var (label, val, col) in kpis)
            p.Append($"<div class=\"kpi\"><b style=\"color:{col}\">{val}</b><span>{label}</span></div>");
        p.Append("</div>");

        p.Append("<div class=\"cards\">");
        foreach (var r in active)
        {
            var ro = r.Roll!;
            var v = ro.Verdict;
            var appExtra = (!string.IsNullOrEmpty(r.App) && !string.Equals(r.App, r.Team, StringComparison.OrdinalIgnoreCase))
                ? " &middot; " + Esc(r.App) : "";
            var sl = Slug(r.Team);
            var endTxt = r.End.HasValue ? r.End.Value.ToString("dd MMM", CultureInfo.InvariantCulture) : "no date";
            p.Append($"<a class=\"cardlink\" href=\"#detail-{sl}\"><div class=\"card\" style=\"border-left-color:{Clr[v]}\">");
            p.Append($"<div style=\"display:flex;justify-content:space-between;align-items:center;gap:8px\"><h3>{Esc(r.Team)}</h3>{Chip(v)}</div>");
            p.Append($"<div class=\"meta\">{Esc(r.Sprint)}{appExtra} &middot; ends {endTxt}</div>");
            p.Append($"<div class=\"m\">Progress {Bar(ro.ProgPct, Clr[v])} <b>{ro.ProgPct}%</b> &middot; {ro.NDone}/{ro.N} stories done</div>");
            if (ro.StretchN > 0)
                p.Append($"<div class=\"m stretch\">+{ro.StretchN} stretch ({ro.StretchDone} done) &middot; not counted in progress</div>");
            p.Append($"<div class=\"m\">Time elapsed {Bar(ro.TimePct, "#8c959f")} {ro.TimePct}%</div>");
            p.Append($"<div class=\"m\">Projected at sprint end: <b>{ro.ProjPct}%</b> &middot; {ro.DaysLeft} wd left &middot; {ro.Blocked} blocked</div>");
            p.Append("</div></a>");
        }
        p.Append("</div>");

        p.Append("<h2 class=\"sec\">Team detail</h2>");
        foreach (var r in active)
        {
            var ro = r.Roll!;
            var v = ro.Verdict;
            var sub = $"{Esc(r.Sprint)} &middot; {(r.Start.HasValue ? r.Start.Value.ToString("yyyy-MM-dd") : "?")} &rarr; {(r.End.HasValue ? r.End.Value.ToString("yyyy-MM-dd") : "?")}";
            p.Append($"<div class=\"team\" id=\"detail-{Slug(r.Team)}\"><details open><summary>");
            p.Append($"<div class=\"th\"><div><h3>{Esc(r.Team)}</h3><div class=\"meta\">{sub}</div></div>" +
                     $"<div>{Chip(v)} &nbsp;<span class=\"sub\">{ro.ProgPct}% done / {ro.TimePct}% time</span></div></div></summary>");
            if (!string.IsNullOrEmpty(r.Note))
                p.Append($"<div class=\"note\">&#9888; {Esc(r.Note)}</div>");
            p.Append("<table><tr><th>Story</th><th>Description</th><th>Owner</th><th>Stage</th><th>Progress</th><th>Forecast finish</th><th>Status</th></tr>");
            var ordered = r.Stories.OrderBy(s => s.Stretch).ThenBy(s => Ord.GetValueOrDefault(s.Flag, 9)).ToList();
            foreach (var s in ordered)
            {
                string etxt;
                if (s.Flag == "Done") etxt = "&#10003;";
                else if (!s.Eta.HasValue) etxt = "&mdash;";
                else
                {
                    var late = r.End.HasValue && s.Eta.Value > r.End.Value;
                    etxt = $"<span class=\"{(late ? "late" : "")}\">{s.Eta.Value.ToString("dd MMM", CultureInfo.InvariantCulture)}{(late ? " (late)" : "")}</span>";
                }
                var tag = s.Stretch ? " <span class=\"stag\">stretch</span>" : "";
                var rowcls = s.Stretch ? " class=\"strow\"" : "";
                var descTrunc = s.Desc.Length > 70 ? s.Desc[..70] : s.Desc;
                var statusTrunc = s.StatusRaw.Length > 22 ? s.StatusRaw[..22] : s.StatusRaw;
                var idTxt = string.IsNullOrEmpty(s.Id) ? "&ndash;" : Esc(s.Id);
                p.Append($"<tr{rowcls}><td><b>{idTxt}</b>{tag}</td><td>{Esc(descTrunc)}</td><td>{Esc(s.Dev)}</td>" +
                         $"<td>{Esc(s.Stage)} <span class=\"sub\">({Esc(statusTrunc)})</span></td>" +
                         $"<td>{Bar(s.Pct, Clr[s.Flag], 80)} {s.Pct}%</td><td>{etxt}</td><td>{Chip(s.Flag)}</td></tr>");
            }
            p.Append("</table></details></div>");
        }

        p.Append($"<div class=\"sub\" style=\"margin-top:24px\">Generated {DateTime.Now:yyyy-MM-dd HH:mm} " +
                 "&middot; progress inferred from each story's workflow stage; forecast uses recorded Dev End Date " +
                 "where present, else remaining-stage &times; time left.</div>");
        p.Append("</div>");
        return p.ToString();
    }

    private static string Render(List<TeamRecord> recs, DateTime today)
    {
        var inner = BuildInner(recs, today);
        return "<!doctype html><html><head><meta charset=\"utf-8\">" +
               "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">" +
               $"<title>Daily Sprint Status Brief</title><style>{Style}</style></head>" +
               $"<body><div class=\"dsb-root\">{inner}</div></body></html>";
    }
}
