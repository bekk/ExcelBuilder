using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Bekk.ExcelCreator.Xml;

namespace Bekk.ExcelCreator.Entities
{
    public class Worksheet
    {
        private readonly Workbook _parent;
        private readonly IList<Cell> _cells = new List<Cell>();
        private NamespaceDirectory _ns;

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

        public void SetValue(uint col, uint row, int value)
        {
            var address = new CellAddress(row, col);
            _cells.Add(new IntegerCell(address, value));
        }

        public XDocument GetDocument()
        {
            var w = _ns.NamespaceMain;
            var worksheet = new XElement(w.GetName("worksheet"));
            var sheetData = new XElement(w.GetName("sheetData"), Rows);
            worksheet.Add(sheetData);
            return new XDocument(_ns.DefaultDeclaration, worksheet);
        }

        private IEnumerable<XElement> Rows => _cells.Select(c => c.Address.Row).Distinct().OrderBy(r => r).Select(GetRow);

        private XElement GetRow(uint index)
        {
            var w = _ns.NamespaceMain;
            var r = new XElement(w.GetName("row"), new XAttribute("r", index+1));
            var cells = _cells.Where(c => c.Address.Row == index).OrderBy(c => c.Address);
            r.Add(cells.Select(c => c.GetElement()));
            return r;
        }
    }
}