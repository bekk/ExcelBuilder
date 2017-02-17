

using System.IO;
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
            sheet.SetValue(3, 5, 42);
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
