namespace Bekk.ExcelBuilder.Contracts
{
    public interface IWorksheet
    {
        string Name { get; }
        ICell SetValue(CellAddress address, int value);
        ICell SetValue(CellAddress address, decimal value);
        ICell SetValue(CellAddress address, double value);
        ICell SetValue(CellAddress address, string value);
    }
	
}