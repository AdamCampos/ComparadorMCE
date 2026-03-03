using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ComparadorMCE.Core.Export
{
    public static class ResumoWriter
    {
        private static bool IsPgSheet(string name)
    => name.StartsWith("PG_", StringComparison.OrdinalIgnoreCase);
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

            using (var rng = resumo.Cells["A1:B1"])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            int row = 2;

            foreach (var ws in package.Workbook.Worksheets)
            {
                if (ws.Name.Equals("RESUMO", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!IsPgSheet(ws.Name))
                    continue;

                resumo.Cells[row, 1].Value = ws.Name;
                resumo.Cells[row, 2].Value = TryExtractMatrixTitle(ws) ?? string.Empty;

                row++;
            }

            resumo.Column(1).Width = 35;
            resumo.Column(2).Width = 90;
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