using System.Collections.Generic;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities.Formatting
{
    class StyleFactory: IHasDocument
    {
        public StyleFactory(params IEnumerable<IHasFormatting>[] formats)
        {
            
        }
        public XDocument GetDocument()
		{
            var ns = new NamespaceDirectory();
			var w = ns.NamespaceMain;
			var styleSheet = new XElement(w.GetName("styleSheet"));
			return new XDocument(ns.DefaultDeclaration, styleSheet);
		} 
    }
}