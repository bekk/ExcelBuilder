﻿using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Entities;
using Bekk.ExcelBuilder.Xml;

namespace Bekk.ExcelBuilder
{
    public class WorkBookBuilder: IWorkbookBuilder
    {
        private readonly Workbook _workbook = new Workbook();
        public IWorkbook Workbook => _workbook;

        public Stream ToStream()
        {
            var wb = _workbook;
            var ns = new PackageNamespaceDirectory();
            var memoryStream = new MemoryStream();
            using (var package = Package.Open(memoryStream, FileMode.Create, FileAccess.ReadWrite))
            {
                var workBook = AddPart(package, null, ns.WorkBookUri, wb, ns.NsWorkbook,
                    ns.NsWorkbookRel, "RiD1");
                AddPart(package, workBook, ns.SharedStringUri, wb.SharedStrings, ns.NsSharedString,
                    ns.NsSharedStringRel, "rIdSharedStrings");
				AddPart(package, workBook, ns.StylesUri, wb.Styles, ns.NsStyles, ns.NsStylesRel, "RiD2");
                foreach (var worksheet in wb.Worksheets)
                {
                    AddPart(package, workBook, ns.WorksheetUri(worksheet.Id), worksheet, ns.NsWorksheet,
                        ns.NsWorksheetRel, worksheet.RelId);
                }

                package.Close();
            }
            memoryStream.Position = 0;
            return memoryStream;
        }

        private PackagePart AddPart(Package package, PackagePart parent, Uri uri, IHasDocument hasDocument, string ns, string nsRel, string relId)
        {
            var document = hasDocument.GetDocument();
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