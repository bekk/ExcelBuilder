using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using Bekk.ExcelCreator.Entities;
using Bekk.ExcelCreator.Xml;

namespace Bekk.ExcelCreator
{
    public class WorkBookBuilder
    {
        public Workbook Workbook { get; } = new Workbook();

        public WorkBookBuilder()
        {
        }

        public Stream ToStream()
        {
            var ns = new PackageNamespaceDirectory();
            var memoryStream = new MemoryStream();
            var package = Package.Open(memoryStream, FileMode.Create, FileAccess.ReadWrite);
            var workBookUri = PackUriHelper.CreatePartUri(ns.WorkBookUri);
            var workBook = package.CreatePart(workBookUri, ns.NsWorkbook);
            using (var workBookStream = workBook.GetStream(FileMode.Create, FileAccess.Write))
            {
                Workbook.GetDocument().Save(workBookStream, SaveOptions.OmitDuplicateNamespaces);
            }
            package.CreateRelationship(workBookUri, TargetMode.Internal, ns.NsWorkbookRel, "rId1");

            foreach (var worksheet in Workbook.Worksheets)
            {
                var uri = PackUriHelper.CreatePartUri(ns.WorksheetUri(worksheet.Id));
                var part = package.CreatePart(uri, ns.NsWorksheet);
                using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
                {
                    worksheet.GetDocument().Save(stream, SaveOptions.OmitDuplicateNamespaces);
                }
                workBook.CreateRelationship(uri, TargetMode.Internal, ns.NsWorksheetRel, worksheet.RelId);
            }
            package.Close();
            memoryStream.Position = 0;
            return memoryStream;
        }
    }
}


//Dim xlPackage As Package = Package.Open(Memstr, FileMode.Create, FileAccess.ReadWrite)
//Dim nsWorkbook As String = "application/vnd.openxmlformats-" + "officedocument.spreadsheetml.sheet.main+xml"
//Dim workbookRelationshipType As String = "http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships/" + "officeDocument"
//Dim workBookUri As Uri = PackUriHelper.CreatePartUri(New Uri("xl/workbook.xml", UriKind.Relative))

//'Create the workbook part.
//Dim wbPart As PackagePart = xlPackage .CreatePart(workBookUri, nsWorkbook)

//'Write the workbook XML to the workbook part.
//Dim workbookStream As Stream = wbPart.GetStream(FileMode.Create, FileAccess.Write)
//Dim fs As New FileStream(TempWorkbook, FileMode.Open)
//fs.CopyTo(workbookStream)
//'Create the relationship for the workbook part.
//xlPackage .CreateRelationship(workBookUri, TargetMode.Internal, workbookRelationshipType, "rId1")

//Dim nsWorksheet As String = "application/vnd.openxmlformats-" + "officedocument.spreadsheetml.worksheet+xml"
//Dim worksheetRelationshipType As String = "http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships/worksheet"
//Dim workSheetUri As Uri
//workSheetUri = PackUriHelper.CreatePartUri(New Uri("xl/worksheets/sheet1.xml", UriKind.Relative))

//'Create the workbook part.
//Dim wsPart As PackagePart = xlPackage.CreatePart(workSheetUri, nsWorksheet)

//'Write the workbook XML to the workbook part.
//Dim worksheetStream As Stream = wsPart.GetStream(FileMode.Create, FileAccess.Write)
//Dim fs As New FileStream(TempWorkSheet1, FileMode.Open)
//fs.CopyTo(worksheetStream)
//' xDoc.Save(worksheetStream)

//'Create the relationship for the workbook part.
//Dim wsworkbookPartUri As Uri = PackUriHelper.CreatePartUri(New Uri("xl/workbook.xml", UriKind.Relative))
//Dim wsworkbookPart As PackagePart = xLPackage.GetPart(wsworkbookPartUri)
//wsworkbookPart.CreateRelationship(workSheetUri, TargetMode.Internal, worksheetRelationshipType, "rId1")