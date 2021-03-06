using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Entities.Collections;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities
{
    class Worksheet : IWorksheet, IHasDocument
    {
        private readonly NamespaceDirectory _ns;
        private readonly IWorksheetCellCollection _cells;
        private readonly SharedStrings _strings;

        public Worksheet(string name, int id, IWorksheetCellCollection cells, SharedStrings strings)
        {
            _strings = strings;
            _cells = cells;
            Name = name;
            Id = id;
            _ns = new NamespaceDirectory();
        }

        public string Name { get; }
        public int Id { get; }
        public string RelId => $"rId{Id}";

        public ICell SetValue(CellAddress address, int value)
        {
            return _cells[address] = new IntegerCell(address, value);
        }

        public ICell SetValue(CellAddress address, decimal value)
        {
            return _cells[address] = new DecimalCell(address, value);
        }
        public ICell SetValue(CellAddress address, double value)
        {
            return _cells[address] = new DoubleCell(address, value);
        }

        public ICell SetValue(CellAddress address, string value)
        {
            return _cells[address] = new TextCell(_strings, address, value);
        }

        public XDocument GetDocument()
        {
            var w = _ns.NamespaceMain;
            var worksheet = new XElement(w.GetName("worksheet"));
            var sheetData = new XElement(w.GetName("sheetData"), Rows);
            worksheet.Add(sheetData);
            return new XDocument(_ns.DefaultDeclaration, worksheet);
        }

        private IEnumerable<XElement> Rows => _cells.Rows.Select(row => {
            var w = _ns.NamespaceMain;
            var r = new XElement(w.GetName("row"), new XAttribute("r", row.rowIndex + 1));
            r.Add(row.cells.Select(c=>c.GetElement()));
            return r;
        });
    }
}