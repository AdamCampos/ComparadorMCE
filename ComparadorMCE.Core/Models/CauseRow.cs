namespace ComparadorMCE.Core.Models
{
    internal sealed class CauseRow
    {
        public int RowIndex { get; init; }

        // Coluna "V" (valor de legenda: P, V, etc.)
        public string? V { get; init; }

        public string? RefDoc { get; init; }
        public string? Interface { get; init; }
        public string? Description { get; init; }
        public string? Voting { get; init; }
        public string? TagNumber { get; init; }
        public string? Delay { get; init; }
    }
}