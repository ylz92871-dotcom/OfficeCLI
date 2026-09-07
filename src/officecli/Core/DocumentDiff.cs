// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeCli.Core;

/// <summary>
/// Headless, dependency-free structural diff between two Office documents of
/// the same format. Normalizes each unit (pptx slide / docx paragraph) to text,
/// then aligns old vs new units by edit-distance similarity to classify them as
/// added / removed / changed / unchanged.
///
/// Deliberately pragmatic: this is a "what moved / what changed" report, not a
/// word-processor byte diff, so it stays independent of the heavy renderers and
/// works on documents opened read-only.
/// </summary>
internal static class DocumentDiff
{
    /// <summary>status values: added | removed | changed | unchanged</summary>
    internal sealed record SlideDiff(int? OldIndex, int? NewIndex, string Status,
        double Similarity, List<string> LinesOld, List<string> LinesNew);

    internal sealed record DiffSummary(string Format, int OldCount, int NewCount,
        int Added, int Removed, int Changed, int Unchanged, List<SlideDiff> Slides)
    {
        public bool HasChanges => Added + Removed + Changed > 0;
    }

    private sealed record Unit(int Index, List<string> Lines)
    {
        public string Key => string.Join("\n", Lines).Trim();
    }

    public static DiffSummary Compare(string oldPath, string newPath)
    {
        var oldExt = Path.GetExtension(oldPath).ToLowerInvariant();
        var newExt = Path.GetExtension(newPath).ToLowerInvariant();
        if (oldExt != newExt)
            throw new ArgumentException($"diff requires two files of the same format, got '{oldExt}' and '{newExt}'.");

        var format = oldExt switch
        {
            ".pptx" => "pptx",
            ".docx" => "docx",
            _ => throw new ArgumentException($"diff supports .pptx and .docx only, got '{oldExt}'.")
        };

        var oldUnits = format == "docx" ? ExtractDocx(oldPath) : ExtractPptx(oldPath);
        var newUnits = format == "docx" ? ExtractDocx(newPath) : ExtractPptx(newPath);

        return Align(oldUnits, newUnits, format);
    }

    // ==================== Extraction ====================

    private static List<Unit> ExtractPptx(string path)
    {
        using var doc = PresentationDocument.Open(path, false);
        var presPart = doc.PresentationPart;
        var pres = presPart?.Presentation;
        var units = new List<Unit>();
        if (pres?.SlideIdList == null) return units;

        int idx = 1;
        foreach (var slideId in pres.SlideIdList.Elements<P.SlideId>())
        {
            var slidePart = presPart?.GetPartById(slideId.RelationshipId!) as SlidePart;
            var lines = slidePart?.Slide?.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                .Select(t => Normalize(t.Text))
                .Where(t => t.Length > 0)
                .ToList() ?? new List<string>();
            if (lines.Count == 0) lines = new List<string> { "(empty slide)" };
            units.Add(new Unit(idx++, lines));
        }
        return units;
    }

    private static List<Unit> ExtractDocx(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart?.Document?.Body;
        var units = new List<Unit>();
        if (body == null) return units;

        int idx = 1;
        foreach (var p in body.Elements<Paragraph>())
        {
            var text = Normalize(p.InnerText);
            if (text.Length == 0) continue;
            units.Add(new Unit(idx++, new List<string> { text }));
        }
        return units;
    }

    /// <summary>Collapse whitespace and trim — strip run layout noise so a pure
    /// formatting change doesn't flag a unit as content-modified.</summary>
    private static string Normalize(string s)
    {
        var sb = new System.Text.StringBuilder(s.Length);
        bool ws = false;
        foreach (var ch in s)
        {
            if (char.IsWhiteSpace(ch)) { ws = true; continue; }
            if (ws) { if (sb.Length > 0) sb.Append(' '); ws = false; }
            sb.Append(ch);
        }
        return sb.ToString().Trim();
    }

    // ==================== Alignment ====================

    private static DiffSummary Align(List<Unit> oldUnits, List<Unit> newUnits, string format)
    {
        var usedOld = new bool[oldUnits.Count];
        var slides = new List<SlideDiff>();
        int added = 0, changed = 0, unchanged = 0;

        foreach (var nu in newUnits)
        {
            int bestIdx = -1;
            double bestSim = 0;
            for (int i = 0; i < oldUnits.Count; i++)
            {
                if (usedOld[i]) continue;
                var sim = Similarity(nu.Key, oldUnits[i].Key);
                if (sim > bestSim) { bestSim = sim; bestIdx = i; }
            }

            // No plausible old twin → this unit was added.
            if (bestIdx < 0 || bestSim < 0.4)
            {
                slides.Add(new SlideDiff(null, nu.Index, "added", 0, new List<string>(), nu.Lines));
                added++;
                continue;
            }

            usedOld[bestIdx] = true;
            var status = bestSim >= 0.99 ? "unchanged" : "changed";
            if (status == "changed") changed++; else unchanged++;
            var oldLines = bestSim < 0.99 ? oldUnits[bestIdx].Lines : new List<string>();
            slides.Add(new SlideDiff(oldUnits[bestIdx].Index, nu.Index, status, bestSim, oldLines, nu.Lines));
        }

        // Any old units never matched were removed.
        int removed = 0;
        for (int i = 0; i < oldUnits.Count; i++)
        {
            if (usedOld[i]) continue;
            removed++;
            slides.Add(new SlideDiff(oldUnits[i].Index, null, "removed", 0, oldUnits[i].Lines, new List<string>()));
        }

        slides.Sort((a, b) => (a.NewIndex ?? int.MaxValue).CompareTo(b.NewIndex ?? int.MaxValue));
        return new DiffSummary(format, oldUnits.Count, newUnits.Count, added, removed, changed, unchanged, slides);
    }

    private static double Similarity(string a, string b)
    {
        if (a == b) return 1.0;
        if (a.Length == 0 || b.Length == 0) return 0.0;
        var dist = EditDistance.Damerau(a.ToLowerInvariant(), b.ToLowerInvariant());
        var max = Math.Max(a.Length, b.Length);
        return 1.0 - (double)dist / max;
    }
}