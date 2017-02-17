using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace Bekk.ExcelCreator.Entities
{
    public struct CellAddress : IComparable<CellAddress>, IEquatable<CellAddress>
    {
        private const int ColBaseNumber = 26;
        public uint Row { get; }
        public uint Col { get; }

        public CellAddress(uint row, uint col)
        {
            Row = row;
            Col = col;
        }

        public CellAddress(string a1)
        {
            try
            {
                if (a1 == null) throw new ArgumentNullException(nameof(a1));
                var match = Regex.Match(a1.ToUpper(), @"(?<Col>[A-Z]+)(?<Row>\d+)");
                var col = match.Groups["Col"];
                Row = uint.Parse(match.Groups["Row"].Value) - 1;
                Col = (uint) ColumnA1Notation(col.Value);
            }
            catch (Exception e)
            {
                throw new ArgumentException($"Cannot parse address \"{a1}\".", nameof(a1), e);
            }
        }

        public override string ToString()
        {
            return ColumnA1Notation((int)Col) + (Row + 1);
        }

        public static CellAddress Parse(string address)
        {
            try
            {
                return new CellAddress(address);
            }
            catch (ArgumentException e)
            {
                throw new FormatException(e.Message, e);
            }
        }

        private static string ColumnA1Notation(int col)
        {
            var digit = col % ColBaseNumber;
            var singleDigit = char.ConvertFromUtf32(digit + 'A');
            if(col < ColBaseNumber) return singleDigit;
            return ColumnA1Notation((col - digit)/ColBaseNumber - 1) + singleDigit;
        }

        private static int ColumnA1Notation(string col)
        {
            if(string.IsNullOrWhiteSpace(col)) throw new ArgumentException("Col cannot be empty", nameof(col));
            var leastDigit = col.Last() - 'A';
            if (col.Length == 1) return leastDigit;
            return leastDigit + (ColumnA1Notation(col.Substring(0, col.Length -1)) + 1) * ColBaseNumber;
        }

        public int CompareTo(CellAddress other)
        {
            if (Row == other.Row) return Col.CompareTo(other.Col);
            return Row.CompareTo(other.Row);
        }

        public bool Equals(CellAddress other)
        {
            return Row == other.Row && Col == other.Col;
        }

        public override int GetHashCode()
        {
            return Row.GetHashCode() + Math.Pow(Col, 2).GetHashCode();
        }

        public override bool Equals(object obj) => obj is CellAddress && Equals((CellAddress) obj);

        public static bool operator ==(CellAddress left, CellAddress right) => left.Equals(right);
        public static bool operator !=(CellAddress left, CellAddress right) => !left.Equals(right);

        public static implicit operator CellAddress(string address) => new CellAddress(address);
    }
}