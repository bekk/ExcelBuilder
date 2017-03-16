using System.IO;

namespace Bekk.ExcelBuilder.Contracts
{
	public interface IWorkbookBuilder
	{
		IWorkbook Workbook { get; }
		Stream ToStream();
	}
}
