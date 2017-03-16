using Bekk.ExcelBuilder.Contracts.Formatting;

namespace Bekk.ExcelBuilder.Contracts
{
	public interface ICell : IFormattable
	{
		CellAddress Address { get; }
	}

}