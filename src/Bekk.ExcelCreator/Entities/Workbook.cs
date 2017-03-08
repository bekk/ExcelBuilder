﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Entities.Collections;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities
{
    public class Workbook : IWorkbook
    {
        private readonly NamespaceDirectory _ns = new NamespaceDirectory();
        private IList<string> _sharedStrings;
        private readonly WorksheetCollection _sheets;

        public Workbook()
        {
            _sheets = new WorksheetCollection(this);
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
                var sheets = _sheets.GetWorksheets();
                if (!sheets.Any()) return new []{new Worksheet("Empty worksheet", 1, this)};
                return sheets;
            }
        }

        IEntityCollection<IWorksheet, string> IWorkbook.Worksheets => throw new NotImplementedException();

        public int AddString(string text)
        {
            if(_sharedStrings == null) _sharedStrings = new List<string>();
            if(!_sharedStrings.Contains(text)) _sharedStrings.Add(text);
            return _sharedStrings.IndexOf(text);
        }
    }
}