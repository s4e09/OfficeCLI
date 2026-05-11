// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Central catalogue of <c>view issues --type</c> accepted values. Single
/// source of truth so the CLI front-end (CommandBuilder.View) and the
/// resident server (ResidentServer.ExecuteView) reject typos identically
/// and the cross-handler protocol documentation cannot drift from what the
/// validator actually accepts.
/// </summary>
public static class IssueSubtypes
{
    public const string FormulaNotEvaluated = "formula_not_evaluated";
    public const string FormulaCacheStale = "formula_cache_stale";
    public const string FieldNotEvaluated = "field_not_evaluated";
    public const string FieldCacheStale = "field_cache_stale";
    public const string SlideFieldNotEvaluated = "slide_field_not_evaluated";
    public const string ChartSeriesRefMissingSheet = "chart_series_ref_missing_sheet";
    public const string ChartCacheStale = "chart_cache_stale";
    public const string DefinedNameBroken = "definedname_broken";
    public const string DefinedNameTargetMissing = "definedname_target_missing";

    /// <summary>Broad IssueType buckets and their single-letter aliases.</summary>
    public static readonly string[] ValidBuckets =
        new[] { "format", "content", "structure", "f", "c", "s" };

    /// <summary>Every subtype the <c>view issues</c> filter accepts by name.</summary>
    public static readonly string[] ValidSubtypes = new[]
    {
        FormulaNotEvaluated, FormulaCacheStale,
        FieldNotEvaluated, FieldCacheStale,
        SlideFieldNotEvaluated,
        ChartSeriesRefMissingSheet, ChartCacheStale,
        DefinedNameBroken, DefinedNameTargetMissing,
    };

    /// <summary>
    /// Validate a user-supplied <c>--type</c> argument. Null and empty pass
    /// through (no filter). Recognised buckets and subtypes (case-insensitive)
    /// pass through. Anything else raises <see cref="CliException"/> with the
    /// full valid list — turning silent typos into a clear failure on both
    /// the CLI front-end and the resident-server fan-out.
    /// </summary>
    public static void Validate(string? issueType)
    {
        if (string.IsNullOrEmpty(issueType)) return;
        var canonical = issueType.ToLowerInvariant();
        foreach (var v in ValidBuckets) if (v == canonical) return;
        foreach (var v in ValidSubtypes) if (v == canonical) return;
        var all = new string[ValidBuckets.Length + ValidSubtypes.Length];
        ValidBuckets.CopyTo(all, 0);
        ValidSubtypes.CopyTo(all, ValidBuckets.Length);
        throw new CliException(
            $"Invalid --type value: '{issueType}'. Valid buckets: format, content, structure (alias f, c, s). Valid subtypes: {string.Join(", ", ValidSubtypes)}.")
        { Code = "invalid_issue_type", ValidValues = all };
    }
}
