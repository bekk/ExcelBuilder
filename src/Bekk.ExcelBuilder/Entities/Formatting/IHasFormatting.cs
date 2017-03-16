using System;
using Bekk.ExcelBuilder.Contracts.Formatting;

namespace Bekk.ExcelBuilder.Entities.Formatting
{
    public interface IHasFormatting : ExcelBuilder.Contracts.Formatting.IFormattable
    {
         Func<int> GetFormatId { set; }
    }
}