// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public enum IssueType
{
    Format,
    Content,
    Structure
}

public enum IssueSeverity
{
    Error,
    Warning,
    Info
}

public class DocumentIssue
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";
    [JsonPropertyName("type")]
    public IssueType Type { get; set; }
    [JsonPropertyName("severity")]
    public IssueSeverity Severity { get; set; }
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";
    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
    [JsonPropertyName("context")]
    public string? Context { get; set; }
    [JsonPropertyName("suggestion")]
    public string? Suggestion { get; set; }
}
