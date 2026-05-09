// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Sheet-wide range-mutation walker. Used by every operation that needs to
// keep range-bearing OOXML structures in sync after a row/column insert,
// delete, move, or copy: cellRef-shift in the SheetData (still done by the
// caller because it requires direction-specific reverse iteration), then
// every sheet-level structure that anchors on an A1 ref/sqref/range:
//
//   - mergeCells
//   - conditionalFormatting (sqref list)
//   - dataValidations (sqref list)
//   - autoFilter (single ref)
//   - hyperlinks (per-cell anchor)
//   - table ref + autoFilter ref (in TableDefinitionPart)
//   - cell formulas (CellFormula.Text and the shared/array CellFormula.Reference)
//   - workbook-level definedNames text (for refs that target this sheet)
//
// The caller supplies axis-specific mappers; the walker handles the
// per-section iteration, the "drop entry when mapper returns null"
// semantics, the "drop container when last entry vanishes" cascade, and
// the per-part Save() bookkeeping (TableDefinitionPart.Save / Workbook.Save).
//
// Out of scope for this walker (intentionally):
//   - <Columns> width/style metadata (column-only, op-asymmetric — handled
//     directly by the column-shift callers).
//   - SheetData cell/row renumbering (axis-direction-specific reverse
//     iteration — handled directly by callers).
//   - CalcChain invalidation (workbook-level concern handled by callers).

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Apply a per-axis ref/formula rewrite across every range-bearing
    /// structure on a sheet. The per-section semantics (drop entry on null,
    /// drop container when empty, save part) are handled internally so the
    /// caller only supplies the axis-specific mappers.
    /// </summary>
    /// <param name="worksheet">The worksheet part being mutated.</param>
    /// <param name="sheetName">Sheet name; threaded to FormulaRefShifter for
    /// the sheet-scope guard (refs targeting other sheets are left alone).</param>
    /// <param name="refMapper">Per-range rewrite. Returns the new ref string,
    /// or null to drop the entry. Used for mergeCells, sqref lists,
    /// autoFilter, hyperlinks, table refs, and the shared/array formula
    /// <c>ref</c> attribute.</param>
    /// <param name="formulaTextMapper">Per-formula-text rewrite (used for
    /// CellFormula.Text and DefinedName text). Pass null to skip formula
    /// and named-range rewriting (rare — only ops that don't touch
    /// formula content).</param>
    private void ApplySheetRangeMutations(
        WorksheetPart worksheet,
        string sheetName,
        Func<string, string?> refMapper,
        Func<string, string>? formulaTextMapper)
    {
        var ws = GetSheet(worksheet);

        // 1. mergeCells
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells != null)
        {
            foreach (var mc in mergeCells.Elements<MergeCell>().ToList())
            {
                if (mc.Reference?.Value == null) continue;
                var shifted = refMapper(mc.Reference.Value);
                if (shifted == null) mc.Remove();
                else mc.Reference = shifted;
            }
            if (!mergeCells.HasChildren) mergeCells.Remove();
        }

        // 2. conditionalFormatting sqref
        foreach (var cf in ws.Elements<ConditionalFormatting>().ToList())
        {
            if (cf.SequenceOfReferences?.HasValue != true) continue;
            var newRefs = cf.SequenceOfReferences.Items
                .Where(r => r.Value != null)
                .Select(r => refMapper(r.Value!))
                .OfType<string>().ToList();
            if (newRefs.Count == 0) cf.Remove();
            else cf.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
        }

        // 3. dataValidations sqref
        var dvs = ws.GetFirstChild<DataValidations>();
        if (dvs != null)
        {
            foreach (var dv in dvs.Elements<DataValidation>().ToList())
            {
                if (dv.SequenceOfReferences?.HasValue != true) continue;
                var newRefs = dv.SequenceOfReferences.Items
                    .Where(r => r.Value != null)
                    .Select(r => refMapper(r.Value!))
                    .OfType<string>().ToList();
                if (newRefs.Count == 0) dv.Remove();
                else dv.SequenceOfReferences = new ListValue<StringValue>(newRefs.Select(r => new StringValue(r)));
            }
            if (!dvs.HasChildren) dvs.Remove();
        }

        // 4. autoFilter
        var af = ws.GetFirstChild<AutoFilter>();
        if (af?.Reference?.Value != null)
        {
            var shifted = refMapper(af.Reference.Value);
            if (shifted != null) af.Reference = shifted;
            else af.Remove();
        }

        // 5. hyperlinks (per-cell anchor)
        var hyperlinks = ws.GetFirstChild<Hyperlinks>();
        if (hyperlinks != null)
        {
            foreach (var hl in hyperlinks.Elements<Hyperlink>().ToList())
            {
                if (hl.Reference?.Value == null) continue;
                var shifted = refMapper(hl.Reference.Value);
                if (shifted == null) hl.Remove();
                else hl.Reference = shifted;
            }
            if (!hyperlinks.HasChildren) hyperlinks.Remove();
        }

        // 6. tables (separate part, must be saved if mutated)
        foreach (var tablePart in worksheet.TableDefinitionParts)
        {
            var tbl = tablePart.Table;
            if (tbl == null) continue;
            bool tblDirty = false;
            if (tbl.Reference?.Value != null)
            {
                var shifted = refMapper(tbl.Reference.Value);
                if (shifted != null && !string.Equals(shifted, tbl.Reference.Value, StringComparison.Ordinal))
                {
                    tbl.Reference = shifted;
                    tblDirty = true;
                }
            }
            if (tbl.AutoFilter?.Reference?.Value != null)
            {
                var shifted = refMapper(tbl.AutoFilter.Reference.Value);
                if (shifted != null && !string.Equals(shifted, tbl.AutoFilter.Reference.Value, StringComparison.Ordinal))
                {
                    tbl.AutoFilter.Reference = shifted;
                    tblDirty = true;
                }
            }
            if (tblDirty) tbl.Save();
        }

        // 7. cell formulas (text + shared/array ref attribute)
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellFormula == null) continue;
                    if (formulaTextMapper != null && !string.IsNullOrEmpty(cell.CellFormula.Text))
                        cell.CellFormula.Text = formulaTextMapper(cell.CellFormula.Text);
                    if (cell.CellFormula.Reference?.Value != null)
                    {
                        var shifted = refMapper(cell.CellFormula.Reference.Value);
                        if (shifted != null) cell.CellFormula.Reference = shifted;
                        else cell.CellFormula.Remove();
                    }
                }
            }
        }

        // 8. workbook-level definedNames whose text references this sheet.
        // Routed through formulaTextMapper (typically a FormulaRefShifter.*
        // call) so the sheet-scope guard inside the shifter handles "leave
        // refs to other sheets alone".
        if (formulaTextMapper != null)
        {
            var definedNames = GetWorkbook().GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                bool changed = false;
                foreach (var dn in definedNames.Elements<DefinedName>())
                {
                    if (dn.Text == null) continue;
                    var newText = formulaTextMapper(dn.Text);
                    if (!string.Equals(newText, dn.Text, StringComparison.Ordinal))
                    {
                        dn.Text = newText;
                        changed = true;
                    }
                }
                if (changed) GetWorkbook().Save();
            }
        }
    }
}
