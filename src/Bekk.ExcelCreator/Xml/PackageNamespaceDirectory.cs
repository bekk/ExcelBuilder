﻿using System;

namespace Bekk.ExcelCreator.Xml
{
    class PackageNamespaceDirectory
    {
        public string NsWorkbook => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        public string NsWorksheet => "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        public string NsSharedString => "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        public string NsWorkbookRel => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public string NsWorksheetRel => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        public string NsSharedStringRel => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        public Uri WorkBookUri => new Uri("xl/workbook.xml", UriKind.Relative);
        public Uri SharedStringUri => new Uri("xl/sharedStrings.xml", UriKind.Relative);
        public Uri WorksheetUri(int worksheetId) => new Uri($"xl/worksheets/sheet{worksheetId}.xml", UriKind.Relative);
    }
}