using System;
using System.IO;
using OfficeOpenXml;
using ComparadorMCE.Core.Export;

namespace ComparadorMCE
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            ExcelPackage.License.SetNonCommercialPersonal("ComparadorMCE");

            var inputPath = args.Length > 0
                ? args[0]
                : @"C:\Projetos\VisualStudio\ComparadorMCE\resources\P83\I-DE-3010.2P-1200-847-KES-001_E.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Arquivo não encontrado: {inputPath}");
                return 1;
            }

            var outPath = Path.Combine(
                Path.GetDirectoryName(inputPath)!,
                Path.GetFileNameWithoutExtension(inputPath) + "_RESUMO.xlsx"
            );

            using var package = new ExcelPackage(new FileInfo(inputPath));

            ResumoWriter.WriteResumo(package);

            package.SaveAs(new FileInfo(outPath));

            Console.WriteLine($"Gerado: {outPath}");
            return 0;
        }
    }
}