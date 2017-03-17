using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Entities.Formatting.Consolidation;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities.Formatting
{
    class StyleFactory: IHasDocument
    {
        private IEnumerable<IEnumerable<IHasFormatting>> _formattables;
        public StyleFactory(params IEnumerable<IHasFormatting>[] formattables)
        {
            _formattables = formattables;
        }
        public XDocument GetDocument()
		{
            var styleRepo = new StyleConsolidator(_formattables.SelectMany(f => f));
            if(!styleRepo.CellFormats.Any()) return null;
            var ns = new NamespaceDirectory();
			var w = ns.NamespaceMain;
			var styleSheet = new XElement(w.GetName("styleSheet"));
            Add(styleSheet, styleRepo.Fonts, new XElement(w.GetName("fonts")));
            AddFills(w, styleSheet);
            AddBorders(w, styleSheet);
            Add(styleSheet, styleRepo.CellFormats, new XElement(w.GetName("cellXfs")));
			return new XDocument(ns.DefaultDeclaration, styleSheet);
		}

        private XElement Add(XElement parent, IHasElement[] elements, XElement container)
        {
            var count = (elements?.Length).GetValueOrDefault();
            if(count > 0)
            {
                container.Add(new XAttribute("count", count));
                container.Add(elements.Select(e => e.GetElement()).ToArray());
                parent.Add(container);
            }
            return parent;
        }

        private void AddFills(XNamespace w, XElement parent)
        {
            parent.Add(new XElement(w.GetName("fills"), new XAttribute("count", 1), new XElement(w.GetName("fill"))));
        }
        private void AddBorders(XNamespace w, XElement parent)
        {
            parent.Add(new XElement(w.GetName("borders"), new XAttribute("count", 1), new XElement(w.GetName("border"))));
        }
    }
}
