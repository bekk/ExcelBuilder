namespace Bekk.ExcelBuilder.Contracts
{
    public interface IWorkbook
    {
        IEntityCollection<IWorksheet, string> Worksheets { get; }
    }
}