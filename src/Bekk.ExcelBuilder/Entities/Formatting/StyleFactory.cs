namespace Bekk.ExcelBuilder.Entities.Formatting
{
    class StyleFactory: IHasDocument
    {
        public StyleFactory(params IEnumerable<IHasFormatting>[] formats)
        {
            
        }
        public XDocument GetStylesDocument()
		{
            var ns = new NamespaceDirectory();
			var w = ns.NamespaceMain;
			var styleSheet = new XElement(w.GetName("styleSheet"));
			return new XDocument(ns.DefaultDeclaration, styleSheet);
		} 
    }
}