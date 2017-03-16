using System.Collections.Generic;
using Bekk.ExcelCreator.Entities.Formatting;
using Bekk.ExcelBuilder.Entities;
using System.Collections;
using System.Linq;
using Bekk.ExcelBuilder.Contracts;

namespace Bekk.ExcelCreator.Entities.Collections
{
    class CellCollection:IEnumerable<IHasFormatting>
    {
        private IList<WorksheetCellCollection> _collections;
        public CellCollection()
        {
            _collections = new List<WorksheetCellCollection>();
        }

        public IWorksheetCellCollection Create(){
            var result = new WorksheetCellCollection();
            _collections.Add(result);
            return result;
        }

        public IEnumerator<IHasFormatting> GetEnumerator()
        {
            return _collections.SelectMany(l => l).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        private class WorksheetCellCollection : IWorksheetCellCollection
        {
            private IDictionary<CellAddress,Cell> _cells = new Dictionary<CellAddress, Cell>();

            public Cell this[CellAddress address] { set => _cells[address] = value; }

            public IEnumerable<(uint rowIndex, IEnumerable<Cell> cells)> Rows => 
                _cells.Keys.Select(a => a.Row).Distinct().OrderBy(r => r)
                .Select(r => _cells.Keys.Where(a => a.Row == r).OrderBy(a => a).Select(a => _cells[a]))
                .Select(row => (row.First().Address.Row, row));

            public IEnumerator<Cell> GetEnumerator() => _cells.Values.OrderBy(c => c.Address).GetEnumerator();

            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}