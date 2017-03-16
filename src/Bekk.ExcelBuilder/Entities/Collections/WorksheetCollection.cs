using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Bekk.ExcelBuilder.Contracts;
using System;

namespace Bekk.ExcelBuilder.Entities.Collections
{
    class WorksheetCollection : IEntityCollection<IWorksheet, string>
    {
        private readonly SharedStrings _strings;
        private IList<Worksheet> _sheets = new List<Worksheet>();
        private readonly CellRepo _cells;

        public WorksheetCollection(SharedStrings strings, CellRepo cells)
        {
            _strings = strings;
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
            var sheet = new Worksheet(name, id, _cells.Create(), _strings);
            _sheets.Add(sheet);
            return sheet;
        }
    }
}