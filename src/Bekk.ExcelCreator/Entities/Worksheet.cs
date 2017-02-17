using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Bekk.ExcelCreator.Xml;

namespace Bekk.ExcelCreator.Entities
{
    public class Worksheet
    {
        private readonly Workbook _parent;
        private readonly IDictionary<CellAddress, Cell> _cells = new ConcurrentDictionary<CellAddress, Cell>();
        private readonly NamespaceDirectory _ns;

        public Worksheet(string name, int id, Workbook parent)
        {
            _parent = parent;
            Name = name;
            Id = id;
            _ns = new NamespaceDirectory();
        }

        public string Name { get; }
        public int Id { get; }
        public string RelId => $"rId{Id}";

        public void SetValue(CellAddress address, int value)
        {
            _cells[address] = new IntegerCell(address, value);
        }

        public void SetValue(CellAddress address, decimal value)
        {
            _cells[address] = new DecimalCell(address, value);
        }
        public void SetValue(CellAddress address, double value)
        {
            _cells[address] = new DoubleCell(address, value);
        }

        public XDocument GetDocument()
        {
            var w = _ns.NamespaceMain;
            var worksheet = new XElement(w.GetName("worksheet"));
            var sheetData = new XElement(w.GetName("sheetData"), Rows);
            worksheet.Add(sheetData);
            return new XDocument(_ns.DefaultDeclaration, worksheet);
        }

        private IEnumerable<XElement> Rows => _cells.Keys.Select(a => a.Row).Distinct().OrderBy(r => r).Select(GetRow);

        private XElement GetRow(uint index)
        {
            var w = _ns.NamespaceMain;
            var r = new XElement(w.GetName("row"), new XAttribute("r", index + 1));
            var cells = _cells.Keys.Where(a => a.Row == index).OrderBy(a => a).Select(a => _cells[a]);
            r.Add(cells.Select(c => c.GetElement()));
            return r;
        }
    }
}