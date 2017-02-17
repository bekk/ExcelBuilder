using System;
using System.IO;
using Bekk.ExcelCreator.Entities;
using NUnit.Framework;

namespace Bekk.ExcelCreator.Tests
{
    public class WorkBookTests
    {
        private Stream Build()
        {
            var target = new WorkBookBuilder();
            var wb = target.Workbook;
            var sheet = wb.CreateWorksheet("Det første");
            sheet.SetValue(new CellAddress(3,5), 42);
            sheet.SetValue(new CellAddress("A2"), 3.5m);
            sheet.SetValue("A2", 13);
            sheet.SetValue(new CellAddress("A3"), Math.PI);
            sheet.SetValue(new CellAddress("f7"), "Dette er en tekst");
            return target.ToStream();
        }

        [Test]
        public void Main()
        {

            var result = Build();

            Assert.That(result, Has.Length.AtLeast(10));

            var path = Path.GetTempFileName();
            using (var file = File.OpenWrite(path))
            {
                result.CopyTo(file);
            }
            File.Move(path, path+".xlsx");
            Assert.Fail("Wrote to: "+path);
        }
        [Test]
        public void Continous()
        {
            var result = Build();

            Assert.That(result, Has.Length.AtLeast(10));
        }
    }
}
