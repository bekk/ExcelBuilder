using System;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Contracts.Formatting;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities.Formatting.Consolidation
{
    class Font : IHasElement, IEquatable<Font>
    {
        private TextStyles _textStyle;
        public Font()
        {
            _textStyle = TextStyles.None;
        }
        public Font(ICellFormat format)
        {
            _textStyle = format.TextStyle;
        }

        public XElement GetElement()
        {
            var ns = new NamespaceDirectory();
            var w = ns.NamespaceMain;
            var font = new XElement(w.GetName("font"));
            if(_textStyle.HasFlag(TextStyles.Bold)){
                font.Add(new XElement(w.GetName("b")));
            }
            if(_textStyle.HasFlag(TextStyles.Italic)){
                font.Add(new XElement(w.GetName("i")));
            }
            return font;
        }

        public bool Equals(Font other)
        {
            return other?._textStyle == _textStyle;
        }
    }
}