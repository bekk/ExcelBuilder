using System;
using System.Collections.Generic;
using System.Linq;

namespace Bekk.ExcelBuilder.Entities.Formatting.Consolidation
{
    class StyleConsolidator
    {
        public Font[] Fonts { get; private set; }
        public CellFormat[] CellFormats { get; private set; }
        public StyleConsolidator(IEnumerable<IHasFormatting> formattables)
        {
            Consolidate(formattables.Where(f => f.HasFormat).ToList());
        }
        private void Consolidate(IEnumerable<IHasFormatting> formattables)
        {
            var fonts = formattables.Select(f => new Font(f.Format)).ToList();
            fonts.Insert(0, new Font());
            Fonts = fonts.Distinct().ToArray();
            var cellFormats = new List<CellFormat>();
            foreach (var formattable in formattables)
            {
                var format = new CellFormat(formattable, Fonts);
                cellFormats.Add(format);
                formattable.GetFormatId = format.GetId = () => Array.IndexOf(CellFormats, format);
            }
            cellFormats = cellFormats.Distinct().ToList();
            if(cellFormats.Any()) cellFormats.Insert(0, new CellFormat());
            CellFormats = cellFormats.Distinct().ToArray();
        }
    }
}