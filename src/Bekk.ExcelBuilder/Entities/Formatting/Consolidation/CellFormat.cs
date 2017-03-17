using System;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities.Formatting.Consolidation
{
    class CellFormat : IHasElement, IEquatable<CellFormat>
    {
        private int fontId;
        public CellFormat()
        {   
        }
        public CellFormat(IHasFormatting formattable, Font[] fonts)
        {
            fontId = Array.IndexOf(fonts, new Font(formattable.Format));
        }

        public bool Equals(CellFormat other)
        {
            return other?.fontId == fontId;
        }

        public XElement GetElement()
        {
            var ns = new NamespaceDirectory();
            var w = ns.NamespaceMain;
            var xf = new XElement(w.GetName("xf"));
            if(fontId > -1){
                xf.Add(new XAttribute("fontId", fontId));
            }
            xf.Add(new XAttribute("xfId", GetId?.Invoke()??0));
            return xf;
        }

        public Func<int> GetId { private get; set; }
    }
}