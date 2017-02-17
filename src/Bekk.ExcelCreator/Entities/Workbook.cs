using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Bekk.ExcelCreator.Xml;

namespace Bekk.ExcelCreator.Entities
{
    public class Workbook
    {
        private readonly IList<Worksheet> _sheets = new List<Worksheet>();

        public Worksheet CreateWorksheet(string name)
        {
            if(_sheets.Select(s=>s.Name).Any(n => n == name)) throw new ArgumentException($"The name {name} is in use", nameof(name));
            var id = _sheets.Any() ? _sheets.Select(s => s.Id).Max() + 1 : 1;
            var sheet = new Worksheet(name, id, this);
            _sheets.Add(sheet);
            return sheet;
        }

        public XDocument GetDocument()
        {
            var ns = new NamespaceDirectory();
            var w = ns.NamespaceMain;
            var workbook = new XElement(w.GetName("workbook"));
            workbook.Add(new XElement(w.GetName("sheets"), 
                Worksheets.Select(ws => new XElement(w.GetName("sheet"),
                new XAttribute("name", ws.Name),
                new XAttribute(ns.NamespaceRel.GetName("id"), ws.RelId),
                new XAttribute("sheetId", ws.Id)))));

            return new XDocument(ns.DefaultDeclaration, workbook);
        }

        public IEnumerable<Worksheet> Worksheets
        {
            get
            {
                if (!_sheets.Any())
                    CreateWorksheet("Empty worksheet");
                return _sheets;
            }
        }
    }
}