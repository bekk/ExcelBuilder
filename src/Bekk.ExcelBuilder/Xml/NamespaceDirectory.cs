using System.Xml.Linq;

namespace Bekk.ExcelBuilder.Xml
{
    class NamespaceDirectory
    {
        public XDeclaration DefaultDeclaration => new XDeclaration("1.0", "UTF-8", "yes");
        public XNamespace NamespaceMain => XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        public XNamespace NamespaceRel => XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    }
}