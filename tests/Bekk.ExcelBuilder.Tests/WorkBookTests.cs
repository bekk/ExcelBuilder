using System;
using System.IO;
using Bekk.ExcelBuilder.Contracts;
using Bekk.ExcelBuilder.Contracts.Formatting;
using Xunit;
using Xunit.Abstractions;

namespace Bekk.ExcelBuilder.Tests
{
    public class WorkBookTests
    {
		readonly ITestOutputHelper output;

		public WorkBookTests(ITestOutputHelper output)
		{
			this.output = output;
		}
		private Stream Build()
        {
            var target = new WorkBookBuilder();
            var wb = target.Workbook;
            var sheet = wb.Worksheets["Det første"];
            sheet.SetValue(new CellAddress(3,5), 42);
            sheet.SetValue(new CellAddress("A2"), 3.5m);
            sheet.SetValue("A2", 13);
            sheet.SetValue(new CellAddress("A3"), Math.PI);
            sheet.SetValue(new CellAddress("f7"), "Dette er en tekst").Format.TextStyle = TextStyle.Bold|TextStyle.Italic;
            return target.ToStream();
        }

        [Fact]
        public void Main()
        {

            var result = Build();

			Assert.True(result.Length >= 9);

            var path = Path.GetTempFileName();
            using (var file = File.OpenWrite(path))
            {
                result.CopyTo(file);
            }
            File.Move(path, path+".xlsx");
			output.WriteLine("Wrote to: " + path);
        }
    }
}
