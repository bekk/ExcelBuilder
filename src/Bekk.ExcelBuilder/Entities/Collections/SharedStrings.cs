using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder.Entities.Collections
{
    public class SharedStrings: IHasDocument
    {
        private List<string> _sharedStrings;

        public int AddString(string text)
        {
            if(_sharedStrings == null) _sharedStrings = new List<string>();
            if(!_sharedStrings.Contains(text)) _sharedStrings.Add(text);
            return _sharedStrings.IndexOf(text);
        }
        		
        public XDocument GetDocument()
        {
            var ns = new NamespaceDirectory();
            var w = ns.NamespaceMain;
            var sst = new XElement(w.GetName("sst"));
            var strings = _sharedStrings ?? Enumerable.Empty<string>();
            sst.Add(new XAttribute("count", strings.Count()));
            sst.Add(new XAttribute("uniqueCount", strings.Count()));
            sst.Add(_sharedStrings
                .Select(txt => 
                new XElement(w.GetName("si"), new XElement(w.GetName("t"), txt))));
            return new XDocument(ns.DefaultDeclaration, sst);
        }        
    }
}