using System;
namespace Bekk.ExcelBuilder.Contracts.Formatting
{
	public interface IFormattable
	{
		ICellFormat Format { get; set; }
	}
}
