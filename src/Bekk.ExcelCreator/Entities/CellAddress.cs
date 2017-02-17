using System;

namespace Bekk.ExcelCreator.Entities
{
    public struct CellAddress : IComparable<CellAddress>
    {
        public uint Row { get; }
        public uint Col { get; }

        public CellAddress(uint row, uint col)
        {
            Row = row;
            Col = col;
        }

        public override string ToString()
        {
            return ColumnA1Notation((int)Col) + (Row + 1);
        }

        private string ColumnA1Notation(int col)
        {
            const int @base = 26;
            var digit = col % @base;
            var singleDigit = char.ConvertFromUtf32(digit + 'A');
            if(col < @base) return singleDigit;
            return ColumnA1Notation((col - digit)/@base - 1) + singleDigit;
        }

        public int CompareTo(CellAddress other)
        {
            if (Row == other.Row) return Col.CompareTo(other.Col);
            return Row.CompareTo(other.Row);
        }
    }
}