namespace Bekk.ExcelBuilder.Contracts
{
    public interface IWorksheet
    {
        string Name { get; }
        void SetValue(CellAddress address, int value);
        void SetValue(CellAddress address, decimal value);
        void SetValue(CellAddress address, double value);
        void SetValue(CellAddress address, string value);
    }
}