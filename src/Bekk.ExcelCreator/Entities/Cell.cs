using System;
using System.Xml.Linq;
using Bekk.ExcelCreator.Xml;

namespace Bekk.ExcelCreator.Entities
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

    class IntegerCell : Cell
    {
        private readonly int _value;

        public IntegerCell(CellAddress address, int value) : base(address)
        {
            _value = value;
        }

        protected override void AddValue(XElement cell, NamespaceDirectory ns)
        {
            cell.Add(new XElement(ns.NamespaceMain.GetName("v"), _value));
        }
    }
}