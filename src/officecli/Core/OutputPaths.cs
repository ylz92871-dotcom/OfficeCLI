// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Resolves the destination directory for user-facing generated artifacts
/// (screenshot / html / svg previews) when the caller did not pass an explicit
/// <c>--out</c>. Prioritizes a stable discoverable location next to the document
/// (<c>&lt;docDir&gt;/.trylo/out</c>) over the system temp dir so generated files
/// don't pollute the workspace root or vanish unreachably in temp.
///
/// Priority: <c>--out</c> (handled by callers) → <c>OFFICECLI_OUT_DIR</c> env →
/// <c>TRYLO_OUT_DIR</c> env → <c>&lt;docDir&gt;/.trylo/out</c> (if writable) →
/// <c>Path.GetTempPath()</c>.
/// </summary>
internal static class OutputPaths
{
    private const string OfficeCliEnv = "OFFICECLI_OUT_DIR";
    private const string TryloEnv = "TRYLO_OUT_DIR";

    /// <summary>
    /// Resolve the directory a generated artifact (screenshot/html/svg/pdf)
    /// should be written into when the user did not pass --out.
    /// </summary>
    public static string ResolveDefaultOutputDir(string? docPath)
    {
        foreach (var env in new[] { OfficeCliEnv, TryloEnv })
        {
            var v = Environment.GetEnvironmentVariable(env);
            if (!string.IsNullOrWhiteSpace(v))
                return v.Trim().TrimEnd('\\', '/');
        }

        if (!string.IsNullOrEmpty(docPath))
        {
            try
            {
                var docDir = Path.GetDirectoryName(Path.GetFullPath(docPath));
                if (!string.IsNullOrEmpty(docDir))
                {
                    var candidate = Path.Combine(docDir, ".trylo", "out");
                    // Probe writability before committing to it, so the temp
                    // fallback still works in read-only regions.
                    Directory.CreateDirectory(candidate);
                    return candidate;
                }
            }
            catch
            {
                /* fall through to temp */
            }
        }

        return Path.GetTempPath();
    }

    /// <summary>
    /// Build a full artifact path for a user-facing generated file. When
    /// <paramref name="outArg"/> is provided it wins verbatim (resolved to an
    /// absolute path); otherwise the artifact lands in the default output dir
    /// with a random-token filename (CWE-59-safe, mirrors the historical temp
    /// writers in this codebase).
    /// </summary>
    /// <param name="extension">e.g. ".png", ".html", ".svg".</param>
    public static string ArtifactPath(string? outArg, string docPath, string kind, string extension)
    {
        if (!string.IsNullOrEmpty(outArg)) return Path.GetFullPath(outArg);
        var dir = ResolveDefaultOutputDir(docPath);
        var stem = $"officecli_{kind}_{Path.GetFileNameWithoutExtension(docPath)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}";
        return Path.Combine(dir, stem + extension);
    }
}