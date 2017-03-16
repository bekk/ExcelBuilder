using System.Collections.Generic;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Entities;

namespace Bekk.ExcelBuilder.Entities.Collections
{
    interface IWorksheetCellCollection : IEnumerable<Cell>
    {
        Cell this[CellAddress address] {set;}
        IEnumerable<(uint rowIndex, IEnumerable<Cell> cells)> Rows {get;}
    }
}