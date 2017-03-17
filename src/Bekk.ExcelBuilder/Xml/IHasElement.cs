using System.Xml.Linq;

namespace Bekk.ExcelBuilder.Xml
{
    public interface IHasElement
    {
         XElement GetElement();
    }
}