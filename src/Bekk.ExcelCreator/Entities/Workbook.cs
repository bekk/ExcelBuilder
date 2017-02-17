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
        private readonly NamespaceDirectory _ns = new NamespaceDirectory();
        private IList<string> _sharedStrings;
 
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
            var w = _ns.NamespaceMain;
            var workbook = new XElement(w.GetName("workbook"));
            workbook.Add(new XElement(w.GetName("sheets"), 
                Worksheets.Select(ws => new XElement(w.GetName("sheet"),
                new XAttribute("name", ws.Name),
                new XAttribute(_ns.NamespaceRel.GetName("id"), ws.RelId),
                new XAttribute("sheetId", ws.Id)))));

            return new XDocument(_ns.DefaultDeclaration, workbook);
        }

        public XDocument GetSharedStringsDocument()
        {
            var w = _ns.NamespaceMain;
            var sst = new XElement(w.GetName("sst"));
            var strings = _sharedStrings ?? Enumerable.Empty<string>();
            sst.Add(new XAttribute("count", strings.Count()));
            sst.Add(new XAttribute("uniqueCount", strings.Count()));
            sst.Add(_sharedStrings
                .Select(txt => 
                new XElement(w.GetName("si"), new XElement(w.GetName("t"), txt))));
            return new XDocument(_ns.DefaultDeclaration, sst);
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

        public int AddString(string text)
        {
            if(_sharedStrings == null) _sharedStrings = new List<string>();
            if(!_sharedStrings.Contains(text)) _sharedStrings.Add(text);
            return _sharedStrings.IndexOf(text);
        }
    }
}