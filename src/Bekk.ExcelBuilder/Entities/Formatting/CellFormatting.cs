using Bekk.ExcelBuilder.Contracts.Formatting;

namespace Bekk.ExcelBuilder.Entities.Formatting
{
	class CellFormatting : ICellFormat
	{
		public TextStyles TextStyle { get;set; }
	}
}
