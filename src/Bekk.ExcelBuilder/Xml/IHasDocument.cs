using System.Xml.Linq;

namespace Bekk.ExcelBuilder.Xml
{
    public interface IHasDocument
    {
         XDocument GetDocument();
    }
}