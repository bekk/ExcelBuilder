using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Entities.Collections;
using Bekk.ExcelBuilder.Entities.Formatting;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities
{
    class Workbook : IWorkbook, IHasDocument
    {
        private readonly NamespaceDirectory _ns = new NamespaceDirectory();
        private readonly SharedStrings _sharedStrings;
        private readonly WorksheetCollection _sheets;
        private readonly CellRepo _cells;
        private readonly StyleFactory _styles;

        public Workbook()
        {
            _sharedStrings = new SharedStrings();
            _cells = new CellRepo();
            _styles = new StyleFactory(_cells);
            _sheets = new WorksheetCollection(_sharedStrings, _cells);
        }

        public XDocument GetDocument()
        {
            var w = _ns.NamespaceMain;
            var workbook = new XElement(w.GetName("workbook"));
            workbook.Add(new XElement(w.GetName("sheets"), 
                Worksheets.Select(ws => new XElement(w.GetName("sheet"),
                new XAttribute("name", ws.Name),
                new XAttribute(_ns.NamespaceRel.GetName("id"), ws.RelId),
                new XAttribute("sheetId", ws.Id)))));

            return new XDocument(_ns.DefaultDeclaration, workbook);
        }

        public IHasDocument Styles => _styles;
		public IHasDocument SharedStrings => _sharedStrings;

        public IEnumerable<Worksheet> Worksheets
        {
            get
            {
                var sheets = _sheets.GetWorksheets();
                if (!sheets.Any()) return new []{new Worksheet("Empty worksheet", 1, _cells.Create(), _sharedStrings)};
                return sheets;
            }
        }

        IEntityCollection<IWorksheet, string> IWorkbook.Worksheets => _sheets;
    }
}