using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ComparadorMCE.Core.Export
{
    public static class ResumoWriter
    {
        private static bool IsPgSheet(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;

            var m = System.Text.RegularExpressions.Regex.Match(
                name.Trim(),
                @"^PG_(\d{2,4})",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (!m.Success) return false;

            if (!int.TryParse(m.Groups[1].Value, out var n)) return false;

            return n >= 16;
        }


        public static ExcelWorksheet EnsureResumoAtFirst(ExcelPackage package)
        {
            if (package == null) throw new ArgumentNullException(nameof(package));

            var wb = package.Workbook ?? throw new InvalidOperationException("Workbook inválido.");
            const string sheetName = "RESUMO";

            var resumo = wb.Worksheets.FirstOrDefault(w =>
                w.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (resumo == null)
                resumo = wb.Worksheets.Add(sheetName);
            else
                resumo.Cells.Clear();

            var first = wb.Worksheets.First();
            if (!first.Name.Equals(resumo.Name, StringComparison.OrdinalIgnoreCase))
                wb.Worksheets.MoveBefore(resumo.Name, first.Name);

            return resumo;
        }

        public static void WriteResumo(ExcelPackage package)
        {
            var resumo = EnsureResumoAtFirst(package);

            resumo.Cells["A1"].Value = "Planilha";
            resumo.Cells["B1"].Value = "Título da Matriz";
            resumo.Cells["C1"].Value = "V";
            resumo.Cells["D1"].Value = "RefDoc";
            resumo.Cells["E1"].Value = "Interface";
            resumo.Cells["F1"].Value = "Description";
            resumo.Cells["G1"].Value = "Voting";
            resumo.Cells["H1"].Value = "TagNumber";
            resumo.Cells["I1"].Value = "Delay";

            using (var rng = resumo.Cells["A1:I1"])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            int row = 2;

            var analyzer = new ComparadorMCE.Core.Excel.WorksheetAnalyzer();

            foreach (var ws in package.Workbook.Worksheets)
            {
                if (ws.Name.Equals("RESUMO", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!IsPgSheet(ws.Name))
                    continue;

                var titulo = TryExtractMatrixTitle(ws) ?? string.Empty;
                var causes = analyzer.ExtractCauses(ws);

                // se não achou causas, mantém 1 linha só com A/B (útil para rastrear)
                if (causes.Count == 0)
                {
                    resumo.Cells[row, 1].Value = ws.Name;
                    resumo.Cells[row, 2].Value = titulo;
                    row++;
                    continue;
                }

                foreach (var c in causes)
                {
                    resumo.Cells[row, 1].Value = ws.Name;
                    resumo.Cells[row, 2].Value = titulo;

                    resumo.Cells[row, 3].Value = c.V;
                    resumo.Cells[row, 4].Value = c.RefDoc;
                    resumo.Cells[row, 5].Value = c.Interface;
                    resumo.Cells[row, 6].Value = c.Description;
                    resumo.Cells[row, 7].Value = c.Voting;
                    resumo.Cells[row, 8].Value = c.TagNumber;
                    resumo.Cells[row, 9].Value = c.Delay;

                    row++;
                }
            }

            resumo.Column(1).Width = 18;   // Planilha
            resumo.Column(2).Width = 80;   // Título
            resumo.Column(3).Width = 6;    // V
            resumo.Column(4).Width = 20;   // RefDoc
            resumo.Column(5).Width = 22;   // Interface
            resumo.Column(6).Width = 90;   // Description
            resumo.Column(7).Width = 18;   // Voting
            resumo.Column(8).Width = 22;   // TagNumber
            resumo.Column(9).Width = 12;   // Delay

            resumo.View.FreezePanes(2, 1);
        }

        /// <summary>
        /// Extrai o "título da matriz" (ex.: "INPUTS FROM INTEGRATOR"), NÃO o cabeçalho do processo/sistema.
        /// Heurística:
        /// - procurar na faixa típica do título da matriz: linhas 50..58, colunas B..AR
        /// - priorizar células mescladas e com fonte maior (títulos costumam ter fonte maior e negrito)
        /// </summary>
        private static string? TryExtractMatrixTitle(ExcelWorksheet ws)
        {
            // Ajuste fino pode ser feito depois se algum arquivo variar
            const int rowStart = 50;
            const int rowEnd = 58;

            // B..AR (2..44) conforme sua região de matriz (causas começam em B)
            const int colStart = 2;  // B
            const int colEnd = 44;   // AR

            string? bestText = null;
            double bestScore = -1;

            for (int r = rowStart; r <= rowEnd; r++)
            {
                for (int c = colStart; c <= colEnd; c++)
                {
                    var cell = ws.Cells[r, c];

                    object? raw;
                    if (cell.Merge)
                    {
                        var mergedAddress = ws.MergedCells[r, c]; // ex.: "B55:AR56"
                        if (!string.IsNullOrWhiteSpace(mergedAddress))
                        {
                            var addr = new ExcelAddressBase(mergedAddress);
                            raw = ws.Cells[addr.Start.Row, addr.Start.Column].Value; // célula superior-esquerda
                        }
                        else
                        {
                            raw = cell.Value;
                        }
                    }
                    else
                    {
                        raw = cell.Value;
                    }

                    var text = raw?.ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    // Pontuação: prioriza fonte maior e merge (título)
                    // (fallback se estilos não estiverem presentes)
                    var fontSize = cell.Style.Font.Size;
                    var bold = cell.Style.Font.Bold;

                    double score = 0;
                    score += Math.Min(text.Length, 200);  // texto relevante mas limitado
                    score += fontSize * 10;               // fonte maior pesa muito
                    score += bold ? 50 : 0;               // negrito ajuda
                    score += cell.Merge ? 80 : 0;         // merge é forte indicador

                    // Evitar pegar "PROCESS SYSTEM ..." (normalmente muito longo e pode estar fora do range)
                    // Mesmo se aparecer, o score do título cinza costuma ganhar por merge + fonte.
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestText = text;
                    }
                }
            }

            if (bestText != null)
                bestText = bestText.Replace("\r", " ").Replace("\n", " ").Trim();

            return bestText;
        }
    }
}