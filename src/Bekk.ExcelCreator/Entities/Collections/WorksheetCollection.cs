using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Bekk.ExcelBuilder.Contracts;
using System;
using Bekk.ExcelCreator.Entities.Collections;

namespace Bekk.ExcelBuilder.Entities.Collections
{
    class WorksheetCollection : IEntityCollection<IWorksheet, string>
    {
        private readonly Workbook _parent;
        private IList<Worksheet> _sheets = new List<Worksheet>();
        private readonly CellCollection _cells;

        public WorksheetCollection(Workbook parent, CellCollection cells)
        {
            _parent = parent;
            _cells = cells;
        }
        public IEnumerable<Worksheet> GetWorksheets() => _sheets;
        public IEnumerator<IWorksheet> GetEnumerator() => _sheets.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public IWorksheet this[string key] => _sheets.FirstOrDefault(s => s.Name == key) ?? CreateWorksheet(key);

        public IWorksheet First()
        {
            if (_sheets.Any()) return _sheets.First();
            return this["Sheet1"];
        }

        private Worksheet CreateWorksheet(string name)
        {
            if (_sheets.Select(s => s.Name).Any(n => n == name)) throw new ArgumentException($"The name {name} is in use", nameof(name));
            var id = _sheets.Any() ? _sheets.Select(s => s.Id).Max() + 1 : 1;
            var sheet = new Worksheet(name, id, _parent, _cells.Create());
            _sheets.Add(sheet);
            return sheet;
        }
    }
}