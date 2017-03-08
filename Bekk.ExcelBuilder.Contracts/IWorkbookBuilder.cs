using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Bekk.ExcelBuilder.Contracts
{
    public interface IWorkbookBuilder
    {
        IWorkbook Workbook { get; }
        Stream ToStream();
    }
}
