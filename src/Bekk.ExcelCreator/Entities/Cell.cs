using System.Globalization;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities
{
    abstract class Cell
    {
        public CellAddress Address { get; }

        protected Cell(CellAddress address)
        {
            Address = address;
        }

        public XElement GetElement()
        {
            var ns = new NamespaceDirectory();
            var w = ns.NamespaceMain;
            var c = new XElement(w.GetName("c"));
            c.Add(new XAttribute("r", Address));
            AddValue(c, ns);
            return c;
        }

        protected abstract void AddValue(XElement cell, NamespaceDirectory ns);
    }

    abstract class NumericCell<T> : Cell
    {
        protected T Value { get; }

        public NumericCell(CellAddress address, T value):base(address)
        {
            Value = value;
        }

        protected abstract string ToString(T value); 
        protected override void AddValue(XElement cell, NamespaceDirectory ns)
        {
            cell.Add(new XElement(ns.NamespaceMain.GetName("v"), ToString(Value)));
        }
    } 

    class IntegerCell : NumericCell<int>
    {
        public IntegerCell(CellAddress address, int value) : base(address, value)
        {
        }

        protected override string ToString(int value) => value.ToString(CultureInfo.InvariantCulture);
    }

    class DecimalCell : NumericCell<decimal>
    {
        public DecimalCell(CellAddress address, decimal value) : base(address, value)
        {
        }

        protected override string ToString(decimal value) => value.ToString(CultureInfo.InvariantCulture);
    }
    class DoubleCell : NumericCell<double>
    {
        public DoubleCell(CellAddress address, double value) : base(address, value)
        {
        }

        protected override string ToString(double value) => value.ToString(CultureInfo.InvariantCulture);
    }

    class TextCell : Cell
    {
        private int _value;

        public TextCell(Workbook root, CellAddress address, string value):base(address)
        {
            _value = root.AddString(value);
        }

        protected override void AddValue(XElement cell, NamespaceDirectory ns)
        {
            cell.Add(new XAttribute("t", "s"));
            cell.Add(new XElement(ns.NamespaceMain.GetName("v"), _value));
        }
    }
}