// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildDumpCommand(Option<bool> jsonOption)
    {
        var dumpFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx)" };
        var formatOpt = new Option<string>("--format")
        {
            Description = "Output format (currently: batch)",
            DefaultValueFactory = _ => "batch"
        };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Write output to a file instead of stdout" };

        var dumpCommand = new Command("dump", "Serialize a document into a replayable batch script (round-trip mechanism)");
        dumpCommand.Add(dumpFileArg);
        dumpCommand.Add(formatOpt);
        dumpCommand.Add(outOpt);
        dumpCommand.Add(jsonOption);

        dumpCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(dumpFileArg)!;
            var format = (result.GetValue(formatOpt) ?? "batch").ToLowerInvariant();
            var outPath = result.GetValue(outOpt);

            if (format != "batch")
                throw new CliException($"Unsupported --format: {format}. Valid: batch")
                    { Code = "invalid_format", ValidValues = ["batch"] };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".docx")
                throw new CliException($"dump currently supports .docx only (got {ext})")
                    { Code = "unsupported_format" };

            using var word = new WordHandler(file.FullName, editable: false);
            var items = BatchEmitter.EmitWord(word);

            // Compact JSON (single line) is the canonical batch wire form:
            // `batch run` consumes it directly and AI tooling pipes it through
            // jq/grep without caring about indentation. We previously
            // constructed a JsonSerializerOptions{WriteIndented=true} that was
            // never threaded into Serialize — kept the compact behavior, just
            // dropped the dead options block.
            var output = JsonSerializer.Serialize(items, BatchJsonContext.Default.ListBatchItem);
            // BUG-R4-FUZZ-3: Unix convention — `--out -` means stdout, not a
            // file literally named "-". Without this, running `dump --out -`
            // silently created a `-` file in the cwd (and could pollute the
            // project tree if invoked from inside it).
            if (outPath == "-")
                outPath = null;
            if (outPath != null)
            {
                // The on-disk file is the canonical batch wire form (bare
                // JSON array) so it can feed `batch --input <file>`
                // unchanged — wrapping it in an envelope would break
                // batch consumption.
                File.WriteAllText(outPath, output);
                if (json)
                {
                    // BUG-R6-01: previously stdout returned
                    //   {"success": true, "data": "/tmp/out.json"}
                    // which was indistinguishable in shape from the
                    // no-out form (data is array). Make the file mode's
                    // envelope unambiguous by surfacing structured
                    // metadata under `data` instead of a bare path
                    // string. Callers can detect "data has outputFile" to
                    // disambiguate.
                    var meta = new System.Text.Json.Nodes.JsonObject
                    {
                        ["outputFile"] = outPath,
                        ["itemCount"] = items.Count
                    };
                    Console.WriteLine(OutputFormatter.WrapEnvelope(meta.ToJsonString()));
                }
                else
                    Console.WriteLine(outPath);
            }
            else
            {
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelope(output));
                else
                    Console.WriteLine(output);
            }
            return 0;
        }, json); });

        return dumpCommand;
    }
}
