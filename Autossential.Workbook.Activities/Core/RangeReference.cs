namespace Autossential.Workbook.Activities.Core
{
    internal readonly struct RangeReference
    {
       public RangeReference()
        {
            Start = new CellReference();
            End = new CellReference();
        }

        public RangeReference(CellReference start, CellReference end)
        {
            Start = start;
            End = end;
        }

        public RangeReference(ExcelFileType type, string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                Start = new CellReference();
                End = CellReference.Max(type);
                return;
            }

            var parts = range.Split(':');
            if (parts.Length == 1)
            {
                Start = new CellReference(type, parts[0]);
                End = CellReference.Max(type);
            }
            else if (parts.Length == 2)
            {
                Start = new CellReference(type, parts[0]);
                End = new CellReference(type, parts[1]);
            }
            else
            {
                throw new ArgumentException("Invalid range address format.", nameof(range));
            }
        }
        public CellReference Start { get; }
        public CellReference End { get; }

        public override string ToString()
        {
            return $"{Start}:{End}";
        }

        public bool IsRowInRange(int row) => row >= Start.Row && row <= End.Row;
        public bool IsColInRange(int col) => col >= Start.Col && col <= End.Col;
    }
}
