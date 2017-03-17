using System;
using System.Globalization;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Contracts.Formatting;
using Bekk.ExcelBuilder.Entities.Formatting;
using Bekk.ExcelBuilder.Entities.Collections;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities
{
    abstract class Cell : ICell, IHasFormatting, IHasElement
    {
		private ICellFormat format;
        public CellAddress Address { get; }
		public ICellFormat Format { get { return format ?? (format = new CellFormatting()); } set { format = value; } }

        public Func<int> GetFormatId { private get; set; }
        public bool HasFormat => format != null;

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
            if(GetFormatId != null){
                c.Add(new XAttribute("s", GetFormatId()));
            }
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

        public TextCell(SharedStrings strings, CellAddress address, string value):base(address)
        {
            _value = strings.AddString(value);
        }

        protected override void AddValue(XElement cell, NamespaceDirectory ns)
        {
            cell.Add(new XAttribute("t", "s"));
            cell.Add(new XElement(ns.NamespaceMain.GetName("v"), _value));
        }
    }
}