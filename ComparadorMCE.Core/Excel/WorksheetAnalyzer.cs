using OfficeOpenXml;
using ComparadorMCE.Core.Models;

namespace ComparadorMCE.Core.Excel
{
    internal sealed class WorksheetAnalyzer
    {
        public List<CauseRow> ExtractCauses(ExcelWorksheet ws)
        {
            var result = new List<CauseRow>();
            var dim = ws.Dimension;
            if (dim == null) return result;

            if (!TryFindCauseHeader(ws, out int headerRow, out var col))
                return result;

            int r = headerRow + 1;

            // lê até encontrar várias linhas vazias (evita parar por “buracos” no meio)
            int emptyStreak = 0;
            const int maxEmptyStreak = 5;

            for (; r <= dim.End.Row; r++)
            {
                bool any = RowHasAny(ws, r, col.Values);
                if (!any)
                {
                    emptyStreak++;
                    if (emptyStreak >= maxEmptyStreak) break;
                    continue;
                }
                emptyStreak = 0;

                // critério: a linha é uma "causa" quando tem valor na coluna V
                var v = Text(ws, r, col["V"]);
                if (string.IsNullOrWhiteSpace(v))
                    continue;

                result.Add(new CauseRow
                {
                    RowIndex = r,
                    V = v,
                    RefDoc = Text(ws, r, col["REFDOC"]),
                    Interface = Text(ws, r, col["INTERFACE"]),
                    Description = Text(ws, r, col["DESCRIPTION"]),
                    Voting = Text(ws, r, col["VOTING"]),
                    TagNumber = Text(ws, r, col["TAGNUMBER"]),
                    Delay = Text(ws, r, col["DELAY"]),
                });
            }

            return result;
        }

        private static string? Text(ExcelWorksheet ws, int r, int c)
        {
            var t = ws.Cells[r, c].Text;
            return string.IsNullOrWhiteSpace(t) ? null : t.Trim();
        }

        private static bool RowHasAny(ExcelWorksheet ws, int row, IEnumerable<int> cols)
        {
            foreach (var c in cols)
            {
                if (!string.IsNullOrWhiteSpace(ws.Cells[row, c].Text))
                    return true;
            }
            return false;
        }

        private static bool TryFindCauseHeader(
            ExcelWorksheet ws,
            out int headerRow,
            out Dictionary<string, int> col)
        {
            headerRow = -1;
            col = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            var dim = ws.Dimension;
            if (dim == null) return false;

            int rowStart = dim.Start.Row;
            int rowEnd = Math.Min(dim.End.Row, rowStart + 300);

            int colStart = dim.Start.Column;
            int colEnd = Math.Min(dim.End.Column, colStart + 80);

            for (int r = rowStart; r <= rowEnd; r++)
            {
                int? vCol = null, refDoc = null, itf = null, desc = null, voting = null, tag = null, delay = null;

                for (int c = colStart; c <= colEnd; c++)
                {
                    var h = NormalizeHeader(ws.Cells[r, c].Text);
                    if (h == "V") vCol = c;
                    else if (h == "REFDOC") refDoc = c;
                    else if (h == "INTERFACE") itf = c;
                    else if (h == "DESCRIPTION") desc = c;
                    else if (h == "VOTING") voting = c;
                    else if (h == "TAGNUMBER") tag = c;
                    else if (h == "DELAY") delay = c;
                }

                // condição mínima para considerar “tabela de causas”
                if (refDoc.HasValue && desc.HasValue)
                {
                    headerRow = r;

                    // V: se não achar header "V", assume 1 coluna à esquerda de REF DOC
                    col["V"] = vCol ?? Math.Max(colStart, refDoc.Value - 1);
                    col["REFDOC"] = refDoc.Value;
                    col["INTERFACE"] = itf ?? (refDoc.Value + 1);
                    col["DESCRIPTION"] = desc.Value;
                    col["VOTING"] = voting ?? (desc.Value + 1);
                    col["TAGNUMBER"] = tag ?? (col["VOTING"] + 1);
                    col["DELAY"] = delay ?? (col["TAGNUMBER"] + 1);

                    return true;
                }
            }

            return false;
        }

        private static string NormalizeHeader(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            // mantém conteúdo, mas normaliza whitespace e case
            var t = text.Trim().ToUpperInvariant();
            t = t.Replace("\r", " ").Replace("\n", " ");
            while (t.Contains("  ")) t = t.Replace("  ", " ");

            // matching por "contains" porque no arquivo real vem "REF DOC <newline> ..."
            if (t == "V") return "V";
            if (t.Contains("REF DOC")) return "REFDOC";
            if (t.Contains("INTERFACE")) return "INTERFACE";
            if (t.Contains("DESCRIPTION")) return "DESCRIPTION";
            if (t.Contains("VOTING")) return "VOTING";
            if (t.Contains("TAG NUMBER")) return "TAGNUMBER";
            if (t == "DELAY" || t.Contains("DELAY")) return "DELAY";

            return t.Replace(" ", "");
        }
    }
}