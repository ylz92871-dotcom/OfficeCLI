// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildDiffCommand(Option<bool> jsonOption)
    {
        var oldArg = new Argument<FileInfo>("old") { Description = "Original document path (.pptx or .docx)" };
        var newArg = new Argument<FileInfo>("new") { Description = "Modified document path (same format as 'old')" };
        var diffCommand = new Command("diff", "Compare two pptx/docx files slide-by-slide / paragraph-by-paragraph");
        diffCommand.Add(oldArg);
        diffCommand.Add(newArg);
        diffCommand.Add(jsonOption);
        diffCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var oldPath = result.GetValue(oldArg)!.FullName;
            var newPath = result.GetValue(newArg)!.FullName;
            var summary = DocumentDiff.Compare(oldPath, newPath);
            if (json)
            {
                var dataJson = JsonSerializer.Serialize(summary, AppJsonContext.Default.DiffSummary);
                Console.WriteLine(OutputFormatter.WrapEnvelope(dataJson, success: true));
            }
            else
            {
                Console.WriteLine(OutputFormatter.FormatDiffText(summary));
            }
            return summary.HasChanges ? 1 : 0;
        }, json); });
        return diffCommand;
    }
}