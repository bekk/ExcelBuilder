using System;
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

        public Stream ToStream()
        {
            var ns = new PackageNamespaceDirectory();
            var memoryStream = new MemoryStream();
            using (var package = Package.Open(memoryStream, FileMode.Create, FileAccess.ReadWrite))
            {
                var workBook = AddPart(package, null, ns.WorkBookUri, Workbook.GetDocument(), ns.NsWorkbook,
                    ns.NsWorkbookRel, "RiD1");
                AddPart(package, workBook, ns.SharedStringUri, Workbook.GetSharedStringsDocument(), ns.NsSharedString,
                    ns.NsSharedStringRel, "rIdSharedStrings");

                foreach (var worksheet in Workbook.Worksheets)
                {
                    AddPart(package, workBook, ns.WorksheetUri(worksheet.Id), worksheet.GetDocument(), ns.NsWorksheet,
                        ns.NsWorksheetRel, worksheet.RelId);
                }

                package.Close();
            }
            memoryStream.Position = 0;
            return memoryStream;
        }

        private PackagePart AddPart(Package package, PackagePart parent, Uri uri, XDocument document, string ns, string nsRel, string relId)
        {
            if (document == null) return null;
            var partUri = PackUriHelper.CreatePartUri(uri);
            var part = package.CreatePart(partUri, ns);
            using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
            {
                document.Save(stream, SaveOptions.OmitDuplicateNamespaces);
            }
            if (parent != null)
            {
                parent.CreateRelationship(partUri, TargetMode.Internal, nsRel, relId);
            }
            else
            {
                package.CreateRelationship(partUri, TargetMode.Internal, nsRel, relId);
            }
            return part;
        }

    }
}